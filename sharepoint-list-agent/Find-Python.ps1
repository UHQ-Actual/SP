<#
Find-Python.ps1 - Locate a genuinely usable Python 3 on Windows.

PATH is the least reliable place to look. On a real machine you hit:
  * C:\...\WindowsApps\python.exe          - a 0-byte App Execution Alias stub
                                             that opens the Microsoft Store
  * C:\Windows\system32\python3            - another stub
  * registry entries pointing at folders   - e.g. a Tcl runtime, or blank
    that contain no python.exe

So this script TRUSTS NOTHING. It gathers candidates from every place Python
hides, then validates each one by actually running it and reading back its
version, venv support and pip support. Only survivors are reported.

Usage:
    powershell -ExecutionPolicy Bypass -File .\Find-Python.ps1
    powershell -ExecutionPolicy Bypass -File .\Find-Python.ps1 -All
    powershell -ExecutionPolicy Bypass -File .\Find-Python.ps1 -Json
    powershell -ExecutionPolicy Bypass -File .\Find-Python.ps1 -Deep -MinVersion 3.11

    # from another script:
    $py = & "$PSScriptRoot\Find-Python.ps1" -Quiet
    & $py -m venv .venv

Caching: the winning path is cached under %LOCALAPPDATA%\spla\python-cache.json
so repeat runs (setup.ps1 calls this every time) skip the sweep. A cached path
is a CLAIM, never a fact - it is re-validated exactly like any other candidate
before it is used, and ignored entirely with -Refresh, -All or -Deep.

Exit codes: 0 = a usable interpreter was found, 1 = none.
Windows PowerShell 5.1 compatible.
#>
[CmdletBinding()]
param(
    [version]$MinVersion = "3.10",
    [string[]]$Hint,            # try these exact paths first (wins if usable)
    [switch]$All,               # show rejected candidates too
    [switch]$Json,              # machine-readable output
    [switch]$Quiet,             # print only the winning path (for $() capture)
    [switch]$Deep,              # also scan common roots on disk (slower)
    [switch]$Refresh,           # ignore the cache and re-sweep
    [switch]$NoOneDriveScan     # skip the OneDrive tree walk entirely
)

$ErrorActionPreference = "Stop"

$CacheFile   = Join-Path $env:LOCALAPPDATA "spla\python-cache.json"
$CacheMaxAge = [TimeSpan]::FromDays(7)

# A hint means the caller knows better than any cache we wrote yesterday.
$HasHint = ($Hint -and $Hint.Count -gt 0) -or [bool]$env:SPLA_PYTHON
$UseCache = -not ($Refresh -or $All -or $Deep -or $HasHint)

# Opt out of the OneDrive walk without editing the command line (used by tests,
# and useful on a laptop whose OneDrive is enormous).
if ($env:SPLA_NO_ONEDRIVE_SCAN -and $env:SPLA_NO_ONEDRIVE_SCAN -ne "0") { $NoOneDriveScan = $true }

# ------------------------------------------------------------------ helpers
$candidates = New-Object System.Collections.ArrayList

function Add-Candidate([string]$Path, [string]$Source) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    try { $full = [System.IO.Path]::GetFullPath($Path) } catch { return }
    $null = $candidates.Add([pscustomobject]@{ Path = $full; Source = $Source })
}

# Add <parent>\<dirFilter>\<relExe> without ever handing a user-supplied path
# to the wildcard parser. -LiteralPath on the parent, -Filter for the pattern:
# a profile folder literally named "u[1]" matches nothing under -Path globbing
# but works fine here.
function Add-ChildCandidates([string]$Parent, [string]$DirFilter, [string]$RelExe, [string]$Source) {
    if ([string]::IsNullOrWhiteSpace($Parent)) { return }
    try {
        if (-not (Test-Path -LiteralPath $Parent -PathType Container)) { return }
        foreach ($d in (Get-ChildItem -LiteralPath $Parent -Filter $DirFilter -Directory -Force -ErrorAction SilentlyContinue)) {
            Add-Candidate (Join-Path $d.FullName $RelExe) $Source
        }
    } catch { Write-Verbose "scan failed for ${Parent}\${DirFilter}: $_" }
}

# Every native invocation in this script goes through here. Under
# $ErrorActionPreference='Stop' a process that writes to stderr raises
# NativeCommandError and kills the script - which is exactly what a
# half-broken interpreter does, so the one thing we are here to survive would
# be the one thing that takes us down. Returns @{ Code; Lines } and never
# throws. Lines are trimmed and blanks dropped, so caller-side parsing does
# not have to cope with launcher chatter or a stray BOM line.
function Invoke-Native([string]$Exe, [string[]]$Arguments) {
    $res = @{ Code = -1; Lines = @() }
    $old = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $out = & $Exe @Arguments 2>$null
        $res.Code  = $LASTEXITCODE
        $res.Lines = @($out | Where-Object { $null -ne $_ -and "$_".Trim() -ne "" } | ForEach-Object { "$_".Trim() })
    } catch {
        $res.Code = -1
        $res.Lines = @()
    } finally {
        $ErrorActionPreference = $old
    }
    return $res
}

# A bounded directory walk.
#
# Get-ChildItem -Recurse in PowerShell 5.1 follows junctions and symlinks, so
# on a user profile it loops (C:\Users\All Users -> C:\ProgramData) and inside
# OneDrive it can wander into cloud-only trees. Unbounded, that turns a
# "find python" call into a minutes-long disk crawl on every setup.ps1 run. So
# we walk it ourselves: explicit queue, depth cap, visited-directory cap,
# wall-clock budget, reparse points skipped outright, and a no-descend list for
# the folders that are always huge and never contain an interpreter root.
function Invoke-BoundedWalk {
    param(
        [string]$Root,
        [int]$MaxDepth = 4,
        [int]$MaxDirs  = 1500,
        [int]$BudgetMs = 4000,
        [scriptblock]$OnDir
    )
    if ([string]::IsNullOrWhiteSpace($Root)) { return }
    try { if (-not (Test-Path -LiteralPath $Root -PathType Container)) { return } } catch { return }

    # Never descend into these. '.conda'/'.venv'/'envs' are handled by $OnDir
    # one level deep, so there is nothing below them we need.
    $noDescend = @(
        '.conda', '.venv', 'envs', 'venv', 'node_modules', '.git', '__pycache__',
        'site-packages', 'Lib', 'lib', 'libs', 'Scripts', 'include', 'share',
        'pkgs', 'conda-meta', '.vscode', '.idea', '.vs', 'WinSxS', 'Windows',
        '$RECYCLE.BIN', 'System Volume Information', 'AppData'
    )

    $sw      = [System.Diagnostics.Stopwatch]::StartNew()
    $visited = 0
    $queue   = New-Object System.Collections.Queue
    $queue.Enqueue([pscustomobject]@{ Path = $Root; Depth = 0 })

    while ($queue.Count -gt 0) {
        if ($visited -ge $MaxDirs -or $sw.ElapsedMilliseconds -ge $BudgetMs) {
            Write-Verbose "walk of '$Root' stopped early: $visited dirs, $($sw.ElapsedMilliseconds) ms"
            break
        }
        $node = $queue.Dequeue()
        $visited++

        $kids = $null
        try { $kids = Get-ChildItem -LiteralPath $node.Path -Directory -Force -ErrorAction SilentlyContinue } catch { continue }
        if (-not $kids) { continue }

        foreach ($k in $kids) {
            # Junctions and symlinks are how a bounded walk becomes an
            # unbounded one. Skip them without looking inside.
            if ($k.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }

            try { & $OnDir $k } catch { Write-Verbose "OnDir failed for $($k.FullName): $_" }

            if (($node.Depth + 1) -lt $MaxDepth -and $noDescend -notcontains $k.Name) {
                $queue.Enqueue([pscustomobject]@{ Path = $k.FullName; Depth = $node.Depth + 1 })
            }
        }
    }
}

# ---------------------------------------------------------------- validation
# Sentinel-tagged output. Printing bare values and reading $lines[0] breaks the
# moment a corporate sitecustomize.py, a PYTHONWARNINGS message or a conda
# activation banner writes to stdout first - and on a managed machine one of
# those usually does. Match the tags wherever they land instead.
$probe = "import sys;print('SPLA_V=%d.%d.%d' % sys.version_info[:3]);print('SPLA_X=' + sys.executable)"

function Get-Tagged([string[]]$Lines, [string]$Tag) {
    $hit = @($Lines | Where-Object { $_ -like "$Tag=*" } | Select-Object -Last 1)
    if ($hit.Count -eq 0) { return $null }
    return $hit[0].Substring($Tag.Length + 1).Trim()
}

function Test-Candidate($cand) {
    $r = [pscustomobject]@{
        Path = $cand.Path; Source = $cand.Source; Version = $null
        HasVenv = $false; HasPip = $false; Usable = $false; Reason = ""
    }
    # -LiteralPath throughout: work paths carry spaces, '-' and '!', and a bare
    # -Path would try to interpret bracket characters as wildcards.
    if (-not (Test-Path -LiteralPath $cand.Path -PathType Leaf)) { $r.Reason = "no such file"; return $r }

    $note = ""
    try {
        $item = Get-Item -LiteralPath $cand.Path -Force
        # The App Execution Alias stubs are 0-byte reparse points. Running one
        # pops the Microsoft Store and can hang, so reject on size BEFORE
        # executing.
        if ($item.Length -eq 0) { $r.Reason = "0-byte Store alias stub"; return $r }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and $cand.Path -like "*\WindowsApps\*") {
            $r.Reason = "Store App Execution Alias stub"; return $r
        }
        # A OneDrive cloud-only placeholder still reports a size; running it
        # forces a download, which is slow but works. Note it so the user knows.
        if ($item.Attributes -band [System.IO.FileAttributes]::Offline) {
            $note = " (OneDrive placeholder - hydrates on first run)"
        }
    } catch { $r.Reason = "unreadable"; return $r }

    $run = Invoke-Native $cand.Path @("-c", $probe)
    if ($run.Code -ne 0 -or $run.Lines.Count -eq 0) { $r.Reason = "did not run$note"; return $r }

    $vRaw = Get-Tagged $run.Lines "SPLA_V"
    $xRaw = Get-Tagged $run.Lines "SPLA_X"
    if (-not $vRaw) { $r.Reason = "no version in output$note"; return $r }
    try { $r.Version = [version]$vRaw } catch { $r.Reason = "unparsable version '$vRaw'$note"; return $r }

    # Canonical sys.executable, but only if it is really there - a venv's
    # python.exe can report a base prefix that has since been uninstalled.
    if ($xRaw -and (Test-Path -LiteralPath $xRaw -PathType Leaf)) { $r.Path = $xRaw }

    if ($r.Version -lt $MinVersion) { $r.Reason = "too old (need $MinVersion+)$note"; return $r }

    $r.HasVenv = ((Invoke-Native $r.Path @("-c", "import venv")).Code -eq 0)
    $r.HasPip  = ((Invoke-Native $r.Path @("-m", "pip", "--version")).Code -eq 0)

    $r.Usable = $true
    if (-not $r.HasVenv)     { $r.Reason = "usable but no venv module$note" }
    elseif (-not $r.HasPip)  { $r.Reason = "usable but no pip$note" }
    else                     { $r.Reason = "ok$note" }
    return $r
}

# ------------------------------------------------------------------- ranking
# Best = usable, has venv+pip, newest version, prefer machine-wide installs.
function Get-Rank($r) {
    $score = 0
    if ($r.Usable) { $score += 1000 }
    # An explicitly hinted interpreter wins outright when it works - the user
    # knows where their environment is better than any heuristic does.
    if ($r.Source -like "hint*")      { $score += 500 }
    if ($r.HasVenv)                   { $score += 400 }
    if ($r.HasPip)                    { $score += 200 }
    if ($r.Source -like "OneDrive*")  { $score += 60 }
    if ($r.Path -like "*\WindowsApps\*") { $score -= 150 }
    if ($env:ProgramFiles -and $r.Path -like "$env:ProgramFiles*") { $score += 40 }
    if ($r.Version) { $score += ($r.Version.Major * 10 + $r.Version.Minor) }
    return $score
}

# --------------------------------------------------------------------- cache
function Read-PythonCache {
    if (-not $UseCache) { return $null }
    try {
        if (-not (Test-Path -LiteralPath $CacheFile -PathType Leaf)) { return $null }
        $raw = Get-Content -LiteralPath $CacheFile -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        $c = $raw | ConvertFrom-Json
        if (-not $c.path) { return $null }
        # Age it out so a new interpreter is eventually noticed even though the
        # cached one still validates fine.
        if ($c.stamped) {
            try {
                if (((Get-Date) - [datetime]::Parse($c.stamped, $null, [Globalization.DateTimeStyles]::RoundtripKind)) -gt $CacheMaxAge) {
                    Write-Verbose "cache expired"
                    return $null
                }
            } catch { return $null }
        }
        return $c
    } catch { Write-Verbose "cache unreadable: $_"; return $null }
}

function Write-PythonCache($best) {
    if (-not $best) { return }
    try {
        $dir = Split-Path -Parent $CacheFile
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [pscustomobject]@{
            path    = $best.Path
            version = $best.Version.ToString()
            stamped = (Get-Date).ToString("o")
        } | ConvertTo-Json | Set-Content -LiteralPath $CacheFile -Encoding UTF8
    } catch { Write-Verbose "cache write failed: $_" }
}

# -------------------------------------------------------------------- output
function Write-Findings {
    param($Best, $Ordered, [switch]$FromCache)

    # Cache from every mode, not just the interactive one - setup.ps1 calls
    # this with -Quiet, and that is the caller that runs it most.
    if ($Best -and -not $FromCache) { Write-PythonCache $Best }

    if ($Quiet) {
        if ($Best) { Write-Output $Best.Path; exit 0 }
        exit 1
    }

    if ($Json) {
        $payload = [pscustomobject]@{
            best       = if ($Best) { $Best.Path } else { $null }
            version    = if ($Best) { $Best.Version.ToString() } else { $null }
            cached     = [bool]$FromCache
            candidates = @($Ordered | ForEach-Object {
                [pscustomobject]@{ path = $_.Path; source = $_.Source
                    version = if ($_.Version) { $_.Version.ToString() } else { $null }
                    usable = $_.Usable; hasVenv = $_.HasVenv; hasPip = $_.HasPip; reason = $_.Reason }
            })
        }
        $payload | ConvertTo-Json -Depth 5
        if ($Best) { exit 0 } else { exit 1 }
    }

    Write-Host ""
    if ($FromCache) {
        Write-Host "[find-python] cached interpreter re-validated (use -Refresh to re-scan)" -ForegroundColor Cyan
    } else {
        Write-Host "[find-python] examined $(@($Ordered).Count) candidate(s)" -ForegroundColor Cyan
    }

    $show = if ($All) { $Ordered } else { @($Ordered | Where-Object { $_.Usable }) }
    foreach ($r in $show) {
        $mark = if ($r.Usable) { "OK  " } else { "skip" }
        $col  = if ($r.Usable) { "Green" } else { "DarkGray" }
        $ver  = if ($r.Version) { $r.Version.ToString() } else { "-" }
        Write-Host ("  {0} {1,-8} {2,-52} [{3}] {4}" -f $mark, $ver, $r.Path, $r.Source, $r.Reason) -ForegroundColor $col
    }
    $rejected = @($Ordered | Where-Object { -not $_.Usable })
    if (-not $All -and $rejected.Count -gt 0) {
        Write-Host "  ($($rejected.Count) rejected - re-run with -All to see them)" -ForegroundColor DarkGray
    }

    Write-Host ""
    if ($Best) {
        Write-Host "[find-python] use: $($Best.Path)  (Python $($Best.Version))" -ForegroundColor Green
        exit 0
    }

    Write-Host "[find-python] No usable Python 3 found (need $MinVersion or newer)." -ForegroundColor Red
    Write-Host "  Try:  -Deep            scan more of the disk" -ForegroundColor Yellow
    Write-Host "        -All             show why each candidate was rejected" -ForegroundColor Yellow
    Write-Host "        -Hint <path>     point straight at a python.exe you know about" -ForegroundColor Yellow
    Write-Host "  Install Python 3.10+ from python.org or your Company Portal," -ForegroundColor Yellow
    Write-Host "  ticking 'Add python.exe to PATH', then re-run this script." -ForegroundColor Yellow
    exit 1
}

# ------------------------------------------------------------- cache fast path
# Runs BEFORE the sweep - that is the whole point of it - but the cached path
# still goes through Test-Candidate, so a python.exe that was uninstalled,
# downgraded or replaced by a stub since we wrote it falls through to the full
# search instead of being handed back.
$cached = Read-PythonCache
if ($cached) {
    $cr = Test-Candidate ([pscustomobject]@{ Path = $cached.path; Source = "cache" })
    if ($cr.Usable) {
        Write-Findings -Best $cr -Ordered @($cr) -FromCache
    } else {
        Write-Verbose "cached '$($cached.path)' rejected: $($cr.Reason) - re-scanning"
    }
}

# ---------------------------------------------------------------- candidates
# 0. Hints win. In order: -Hint, $env:SPLA_PYTHON, a pinned python-path.txt
# next to this script, then conda envs parked in a OneDrive project tree (the
# usual shape on a managed work laptop, e.g.
#   %OneDriveCommercial%\1a_Projects\1!_Code\Python\.conda\python.exe
# note the spaces, '-' and '!' in those paths - everything here uses
# -LiteralPath so odd characters can't be mangled into wildcards).
foreach ($h in $Hint) { Add-Candidate $h "hint (-Hint)" }
if ($env:SPLA_PYTHON) { Add-Candidate $env:SPLA_PYTHON "hint (SPLA_PYTHON)" }

$pinFile = Join-Path $PSScriptRoot "python-path.txt"
try {
    if (Test-Path -LiteralPath $pinFile) {
        foreach ($line in (Get-Content -LiteralPath $pinFile)) {
            $t = $line.Trim()
            if ($t -and -not $t.StartsWith("#")) { Add-Candidate $t "hint (python-path.txt)" }
        }
    }
} catch { Write-Verbose "pin file unreadable: $_" }

# OneDrive roots: the env vars are authoritative, but fall back to enumerating
# the profile because a work tenant names the folder "OneDrive - <Agency>".
$oneDrives = New-Object System.Collections.ArrayList
foreach ($v in @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
    if ($v -and (Test-Path -LiteralPath $v)) { $null = $oneDrives.Add($v) }
}
try {
    if ($env:USERPROFILE -and (Test-Path -LiteralPath $env:USERPROFILE -PathType Container)) {
        foreach ($d in (Get-ChildItem -LiteralPath $env:USERPROFILE -Filter "OneDrive*" -Directory -Force -ErrorAction SilentlyContinue)) {
            if ($oneDrives -notcontains $d.FullName) { $null = $oneDrives.Add($d.FullName) }
        }
    }
} catch { Write-Verbose "OneDrive root probe failed: $_" }

foreach ($od in $oneDrives) {
    # The exact known layout first (cheap, no recursion, always attempted).
    foreach ($rel in @(
        "1a_Projects\1!_Code\Python\.conda\python.exe",
        "1a_Projects\1!_Code\Python\.conda\Scripts\python.exe"
    )) { Add-Candidate (Join-Path $od $rel) "OneDrive project conda" }

    if ($NoOneDriveScan) { continue }

    # Then any .conda / envs / .venv environment parked in the project tree -
    # bounded, so a 200k-file OneDrive costs seconds, not minutes.
    Invoke-BoundedWalk -Root $od -MaxDepth 5 -MaxDirs 1500 -BudgetMs 4000 -OnDir {
        param($d)
        if (@('.conda', '.venv', 'envs', 'venv') -notcontains $d.Name) { return }
        foreach ($rel in @("python.exe", "Scripts\python.exe")) {
            $p = Join-Path $d.FullName $rel
            if (Test-Path -LiteralPath $p -PathType Leaf) { Add-Candidate $p "OneDrive conda/venv" }
        }
        foreach ($sub in (Get-ChildItem -LiteralPath $d.FullName -Directory -Force -ErrorAction SilentlyContinue)) {
            if ($sub.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
            foreach ($rel in @("python.exe", "Scripts\python.exe")) {
                $p = Join-Path $sub.FullName $rel
                if (Test-Path -LiteralPath $p -PathType Leaf) { Add-Candidate $p "OneDrive conda env" }
            }
        }
    }
}

# Conda environments in their conventional homes.
foreach ($envRoot in @(
    "$env:USERPROFILE\.conda\envs",
    "$env:LOCALAPPDATA\conda\conda\envs",
    "$env:ProgramData\Anaconda3\envs",
    "$env:USERPROFILE\Anaconda3\envs",
    "$env:USERPROFILE\miniconda3\envs",
    "$env:LOCALAPPDATA\miniconda3\envs"
)) {
    Add-ChildCandidates $envRoot "*" "python.exe" "conda env"
}

# 1. The py launcher knows about every registered install ("py -0p").
try {
    $pyExe = (Get-Command py.exe -ErrorAction SilentlyContinue)
    if ($pyExe) {
        $run = Invoke-Native $pyExe.Source @("-0p")
        foreach ($line in $run.Lines) {
            # lines look like " -V:3.12 *        C:\Program Files\Python312\python.exe"
            $m = [regex]::Match($line, '([A-Za-z]:\\[^\r\n]*?python\.exe)')
            if ($m.Success) { Add-Candidate $m.Groups[1].Value "py -0p" }
        }
        # Older launchers have no -0p; ask the default interpreter directly.
        if ($run.Lines.Count -eq 0) {
            $one = Invoke-Native $pyExe.Source @("-3", "-c", "import sys;print(sys.executable)")
            if ($one.Code -eq 0 -and $one.Lines.Count -gt 0) {
                Add-Candidate ($one.Lines | Select-Object -Last 1) "py -3"
            }
        }
    }
} catch { Write-Verbose "py launcher probe failed: $_" }

# 2. PEP 514 registry entries (may be stale, blank, or point at non-Python).
foreach ($hive in @("HKCU:\Software\Python", "HKLM:\Software\Python", "HKLM:\Software\WOW6432Node\Python")) {
    try {
        if (-not (Test-Path -LiteralPath $hive)) { continue }
        foreach ($company in Get-ChildItem -LiteralPath $hive -ErrorAction SilentlyContinue) {
            foreach ($tag in Get-ChildItem -LiteralPath $company.PSPath -ErrorAction SilentlyContinue) {
                # Key names are user-controlled and routinely contain brackets;
                # -LiteralPath keeps them out of the wildcard parser.
                $ipKey = Join-Path $tag.PSPath "InstallPath"
                if (-not (Test-Path -LiteralPath $ipKey)) { continue }
                $ip = Get-ItemProperty -LiteralPath $ipKey -ErrorAction SilentlyContinue
                if (-not $ip) { continue }
                if ($ip.ExecutablePath) { Add-Candidate $ip.ExecutablePath "registry $($tag.PSChildName)" }
                $dir = $ip.'(default)'
                if ($dir) { Add-Candidate (Join-Path $dir "python.exe") "registry $($tag.PSChildName)" }
            }
        }
    } catch { Write-Verbose "registry probe failed for ${hive}: $_" }
}

# 3. PATH (stubs get filtered during validation).
foreach ($n in @("python.exe", "python3.exe")) {
    try {
        foreach ($c in (Get-Command $n -All -ErrorAction SilentlyContinue)) {
            if ($c.Source) { Add-Candidate $c.Source "PATH" }
        }
    } catch { Write-Verbose "PATH probe failed for ${n}: $_" }
}

# 4. Well-known install locations, including the real (non-stub) Store package.
Add-ChildCandidates "$env:LOCALAPPDATA\Programs\Python" "Python3*" "python.exe" "well-known"
Add-ChildCandidates "$env:ProgramFiles"                 "Python3*" "python.exe" "well-known"
Add-ChildCandidates "${env:ProgramFiles(x86)}"          "Python3*" "python.exe" "well-known"
Add-ChildCandidates "$env:SystemDrive\"                 "Python3*" "python.exe" "well-known"
Add-ChildCandidates "$env:SystemDrive\Tools"            "Python*"  "python.exe" "well-known"

# The genuine Store package sits two wildcard levels deep, so walk it in two.
try {
    $pkgRoot = "$env:LOCALAPPDATA\Packages"
    if (Test-Path -LiteralPath $pkgRoot -PathType Container) {
        foreach ($pkg in (Get-ChildItem -LiteralPath $pkgRoot -Filter "PythonSoftwareFoundation.Python.3*" -Directory -Force -ErrorAction SilentlyContinue)) {
            Add-ChildCandidates (Join-Path $pkg.FullName "LocalCache\local-programs\Python") "Python3*" "python.exe" "Store package"
        }
    }
} catch { Write-Verbose "Store package probe failed: $_" }

foreach ($fixed in @(
    "$env:ProgramData\Anaconda3\python.exe",
    "$env:USERPROFILE\Anaconda3\python.exe",
    "$env:USERPROFILE\miniconda3\python.exe",
    "$env:LOCALAPPDATA\miniconda3\python.exe",
    "$env:USERPROFILE\scoop\apps\python\current\python.exe",
    "$env:ProgramData\chocolatey\bin\python.exe"
)) { Add-Candidate $fixed "well-known" }

# 5. Whatever VS Code was told to use (the work PC runs Copilot in VS Code).
foreach ($settings in @(
    "$env:APPDATA\Code\User\settings.json",
    (Join-Path $PSScriptRoot ".vscode\settings.json"),
    (Join-Path (Get-Location).Path ".vscode\settings.json")
)) {
    try {
        if (Test-Path -LiteralPath $settings -PathType Leaf) {
            $raw = Get-Content -LiteralPath $settings -Raw -ErrorAction SilentlyContinue
            $m = [regex]::Match("$raw", '"python\.defaultInterpreterPath"\s*:\s*"([^"]+)"')
            if ($m.Success) { Add-Candidate ($m.Groups[1].Value -replace '\\\\', '\') "VS Code setting" }
        }
    } catch { Write-Verbose "VS Code settings probe failed for ${settings}: $_" }
}

# 6. Optional bounded scan for machines that hide Python somewhere custom.
# Same walker as the OneDrive pass, so -Deep cannot loop through a junction or
# run for minutes either - it just gets a bigger budget.
if ($Deep) {
    foreach ($root in @("$env:SystemDrive\", "$env:USERPROFILE", "$env:ProgramFiles", "$env:ProgramData")) {
        Invoke-BoundedWalk -Root $root -MaxDepth 4 -MaxDirs 4000 -BudgetMs 15000 -OnDir {
            param($d)
            $p = Join-Path $d.FullName "python.exe"
            if (Test-Path -LiteralPath $p -PathType Leaf) { Add-Candidate $p "deep scan" }
        }
    }
}

# ------------------------------------------------------------------ validate
# De-dupe by path before the (relatively expensive) execution probes.
$seen   = @{}
$unique = @()
foreach ($c in $candidates) {
    $k = $c.Path.ToLowerInvariant()
    if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $unique += $c }
}

$results = @()
foreach ($c in $unique) { $results += (Test-Candidate $c) }

# Collapse duplicates that resolved to the same sys.executable.
$byPath = @{}
foreach ($r in $results) {
    $k = $r.Path.ToLowerInvariant()
    if (-not $byPath.ContainsKey($k)) { $byPath[$k] = $r }
    elseif ($r.Usable -and -not $byPath[$k].Usable) { $byPath[$k] = $r }
}
$results = @($byPath.Values)

$ordered = @($results | Sort-Object -Property @{ Expression = { Get-Rank $_ } } -Descending)
$best    = $ordered | Where-Object { $_.Usable } | Select-Object -First 1

Write-Findings -Best $best -Ordered $ordered
