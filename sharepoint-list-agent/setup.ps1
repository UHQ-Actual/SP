<#
setup.ps1 - Make the SharePoint List Agent work on a Windows PC (no WSL needed).

Usage (from any PowerShell prompt, in this folder):
    powershell -ExecutionPolicy Bypass -File .\setup.ps1
    powershell -ExecutionPolicy Bypass -File .\setup.ps1 -Test
    powershell -ExecutionPolicy Bypass -File .\setup.ps1 -SkipVenv
    powershell -ExecutionPolicy Bypass -File .\setup.ps1 -Root "C:\Tools\SP\sharepoint-list-agent"

What it does:
    1. Finds a real Python 3.10+ ('py -3', then 'python', then 'python3').
    2. Creates server\.venv and installs the mcp SDK into it.
    3. Runs server\test_agent.py (the data-logic check). -Test also runs demo.py.
    4. Ensures config\exports\ exists.
    5. Prints the exact .vscode\mcp.json block for GitHub Copilot in VS Code,
       with the absolute interpreter and server paths already filled in.

Nothing here touches SharePoint or the network except pip.
#>
param(
    [string]$Root = $PSScriptRoot,
    [switch]$SkipVenv,
    [switch]$Test
)

$ErrorActionPreference = "Stop"

# Windows PowerShell 5.1 defaults to TLS 1.0, which PyPI and GitHub reject.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Keep Python's output readable: make it emit UTF-8 and make the console read
# UTF-8 back, otherwise non-ASCII characters in test output arrive as garbage.
$env:PYTHONIOENCODING = "utf-8"
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

function Say($msg)  { Write-Host "[setup] $msg" }
function Warn($msg) { Write-Host "[setup] WARNING: $msg" -ForegroundColor Yellow }

# $PSScriptRoot is empty when the script is dot-sourced rather than run as a file.
if (-not $Root -or $Root -eq "") { $Root = (Get-Location).Path }

# Resolve-Path throws under $ErrorActionPreference='Stop', so a typo'd -Root
# would abort with a bare .NET "Cannot find path" before ever reaching the
# wrong-folder guidance below. Catch it and say something useful instead.
function Show-RootHelp($why) {
    Say $why
    Say "Point -Root at the sharepoint-list-agent folder, for example:"
    Say "  powershell -ExecutionPolicy Bypass -File .\setup.ps1 -Root C:\SP\sharepoint-list-agent"
    Say "Or just cd into that folder and run .\setup.ps1 with no -Root at all."
}

if (-not (Test-Path -LiteralPath $Root)) {
    Show-RootHelp "The folder '$Root' does not exist."
    # -ErrorAction Continue so the script reaches 'exit 1' and returns a real exit code.
    Write-Error "[setup] -Root not found: '$Root'" -ErrorAction Continue
    exit 1
}

try {
    $Root = (Resolve-Path -LiteralPath $Root).Path
} catch {
    Show-RootHelp "Could not resolve '$Root' ($($_.Exception.Message))."
    Write-Error "[setup] -Root could not be resolved: '$Root'" -ErrorAction Continue
    exit 1
}

$ServerDir    = Join-Path $Root "server"
$ConfigDir    = Join-Path $Root "config"
$ExportDir    = Join-Path $ConfigDir "exports"
$ServerScript = Join-Path $ServerDir "sharepoint_list_mcp.py"
$TestScript   = Join-Path $ServerDir "test_agent.py"
$DemoScript   = Join-Path $ServerDir "demo.py"
$ReqFile      = Join-Path $ServerDir "requirements.txt"
$VenvDir      = Join-Path $ServerDir ".venv"
$VenvPy       = Join-Path $VenvDir "Scripts\python.exe"

Say "Repo root: $Root"

if (-not (Test-Path -LiteralPath $ServerScript)) {
    Say "Could not find server\sharepoint_list_mcp.py under '$Root'."
    Say "Run this script from inside the sharepoint-list-agent folder, for example:"
    Say "  cd C:\SP\sharepoint-list-agent"
    Say "  powershell -ExecutionPolicy Bypass -File .\setup.ps1"
    Say "Or point it at the folder explicitly with -Root <path>."
    # -ErrorAction Continue so the script reaches 'exit 1' and returns a real exit code.
    Write-Error "[setup] Wrong folder - server\sharepoint_list_mcp.py not found under '$Root'." -ErrorAction Continue
    exit 1
}

# ---------------------------------------------------------------- find Python
# Returns a hashtable @{ Exe; Pre; Display; Version } or $null.
# 'py -3' is tried first because the launcher is the reliable Windows entry
# point; 'python3' is tried last because on Windows it is usually the Microsoft
# Store App Execution Alias stub, which prints no version and exits non-zero.
function Test-PythonCandidate($exe, $pre, $display) {
    if (-not (Get-Command $exe -ErrorAction SilentlyContinue)) { return $null }

    $out = $null
    $code = 1
    try {
        # Native stderr must not become a terminating error while probing.
        $old = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        $out = (& $exe @pre --version 2>&1 | Out-String)
        $code = $LASTEXITCODE
        $ErrorActionPreference = $old
    } catch {
        $ErrorActionPreference = "Stop"
        Say "  '$display' could not be launched - skipping."
        return $null
    }

    if ($code -ne 0 -or -not $out -or $out.Trim() -eq "") {
        Say "  '$display' produced no usable version output (exit $code) - skipping."
        return $null
    }
    $m = [regex]::Match($out, 'Python\s+(\d+)\.(\d+)')
    if (-not $m.Success) {
        Say "  '$display' printed '$($out.Trim())' - not a Python version, skipping."
        return $null
    }

    $major = [int]$m.Groups[1].Value
    $minor = [int]$m.Groups[2].Value
    if ($major -lt 3 -or ($major -eq 3 -and $minor -lt 10)) {
        Say "  '$display' is Python $major.$minor - too old (need 3.10+), skipping."
        return $null
    }

    return @{ Exe = $exe; Pre = $pre; Display = $display; Version = "$major.$minor" }
}

Say "Looking for Python 3.10 or newer..."

# Preferred: Find-Python.ps1 searches the py launcher, the PEP 514 registry,
# PATH, every well-known install dir and VS Code's configured interpreter, then
# VALIDATES each by running it. That matters on locked-down machines where PATH
# yields a 0-byte Store alias stub and the registry can point at a folder with
# no python.exe in it. The simple probes below are only a fallback.
$py = $null
$finder = Join-Path $Root "Find-Python.ps1"
if (Test-Path -LiteralPath $finder) {
    try {
        # Native stderr from the interpreters Find-Python probes must not become
        # a terminating error here, and its stdout may carry more than one line.
        $old = $ErrorActionPreference
        $ErrorActionPreference = "Continue"
        $found = & $finder -Quiet -MinVersion "3.10"
        $finderCode = $LASTEXITCODE
        $exe = $null
        if ($finderCode -eq 0 -and $found) {
            # Last non-empty line only - a sitecustomize.py print or launcher
            # chatter must not end up glued onto the interpreter path.
            $exe = @($found | Where-Object { $_ -and "$_".Trim() -ne "" } | Select-Object -Last 1)[0]
            if ($exe) { $exe = "$exe".Trim() }
        }
        if ($exe -and (Test-Path -LiteralPath $exe -PathType Leaf)) {
            $verOut = (& $exe -c "import sys;print('%d.%d' % sys.version_info[:2])" 2>$null)
            $ver = @($verOut | Where-Object { $_ -and "$_".Trim() -ne "" } | Select-Object -Last 1)[0]
            $py = @{ Exe = $exe; Pre = @(); Display = $exe; Version = "$("$ver".Trim())" }
        } elseif ($exe) {
            Say "  Find-Python.ps1 returned '$exe', which is not a file - trying the simple probes."
        } else {
            Say "  Find-Python.ps1 found nothing usable - trying the simple probes."
        }
        $ErrorActionPreference = $old
    } catch {
        $ErrorActionPreference = "Stop"
        Say "  Find-Python.ps1 failed ($($_.Exception.Message)) - trying the simple probes."
    }
}

if (-not $py) { $py = Test-PythonCandidate "py"      @("-3") "py -3" }
if (-not $py) { $py = Test-PythonCandidate "python"  @() "python" }
if (-not $py) { $py = Test-PythonCandidate "python3" @() "python3" }

if (-not $py) {
    Write-Host ""
    Say "No usable Python 3.10+ was found on this machine."
    Say ""
    Say "Fix it, then re-run this script:"
    Say "  1. Install Python 3.10 or newer - from https://www.python.org/downloads/windows/"
    Say "     or from the Company Portal / Software Center if downloads are blocked."
    Say "  2. During install, tick 'Add python.exe to PATH' (and keep the py launcher)."
    Say "  3. Close and reopen PowerShell so the new PATH is picked up."
    Say "  4. powershell -ExecutionPolicy Bypass -File .\setup.ps1"
    Say ""
    Say "Note: a bare 'python3' on Windows is usually the Microsoft Store alias stub,"
    Say "not a real interpreter. Turn it off under Settings > Apps > Advanced app"
    Say "settings > App execution aliases if it keeps opening the Store."
    Write-Host ""
    # -ErrorAction Continue so the script reaches 'exit 1' and returns a real exit code.
    Write-Error "[setup] No Python 3.10+ found - install Python 3.10 or newer and re-run .\setup.ps1" -ErrorAction Continue
    exit 1
}

Say "Using $($py.Display)  (Python $($py.Version))"

# Absolute path of the interpreter, straight from Python itself. This is what
# the MCP config must point at - 'py' and 'python' are shims, not interpreters.
#
# Guarded like every other native call: interpreter stderr must not raise a
# NativeCommandError under 'Stop'. Out-String is deliberately NOT used - a
# corporate sitecustomize.py print, a PYTHONWARNINGS line or launcher chatter
# would be glued onto the path, and this value is both executed below and
# written verbatim into the mcp.json block the user pastes. Last non-empty
# line only, then proved to be a real file before it is trusted.
$pre = $py.Pre
$ErrorActionPreference = "Continue"
$pyPathOut = (& $py.Exe @pre -c "import sys; print(sys.executable)" 2>$null)
$pyPathCode = $LASTEXITCODE
$ErrorActionPreference = "Stop"

$BasePython = @($pyPathOut | Where-Object { $_ -and "$_".Trim() -ne "" } | Select-Object -Last 1)[0]
if ($BasePython) { $BasePython = "$BasePython".Trim() }

if (-not $BasePython) {
    Say "'$($py.Display)' ran but printed no interpreter path (exit $pyPathCode)."
    Say "That usually means a wrapper or alias is intercepting it rather than a real Python."
    Write-Error "[setup] Could not resolve the absolute path of the Python interpreter." -ErrorAction Continue
    exit 1
}
if (-not (Test-Path -LiteralPath $BasePython -PathType Leaf)) {
    Say "'$($py.Display)' reported its interpreter as '$BasePython', but no such file exists."
    Say "Re-run with -Root pointing at this folder after fixing the Python install,"
    Say "or set SPLA_PYTHON to the full path of a working python.exe."
    Write-Error "[setup] Reported interpreter path does not exist: '$BasePython'" -ErrorAction Continue
    exit 1
}
Say "Interpreter: $BasePython"

# ------------------------------------------------------------- exports folder
if (-not (Test-Path -LiteralPath $ExportDir)) {
    New-Item -ItemType Directory -Path $ExportDir -Force | Out-Null
    Say "Created $ExportDir"
} else {
    Say "Export folder present: $ExportDir"
}

# ---------------------------------------------------------------------- venv
# $McpInterpreter is what ends up in mcp.json: the venv python when we have one,
# otherwise the base interpreter (with a --user install of mcp behind it).
$McpInterpreter = $BasePython
$RunPython      = $BasePython   # what we run test/demo with
$UsedUserSite   = $false

if ($SkipVenv) {
    Say "-SkipVenv given - not creating a virtual environment."
    if (Test-Path -LiteralPath $VenvPy) {
        Say "Existing venv found; using $VenvPy"
        $McpInterpreter = $VenvPy
        $RunPython      = $VenvPy
    }
} else {
    # A .venv built on Linux/WSL has bin\python, not Scripts\python.exe. Rebuild it.
    if ((Test-Path -LiteralPath $VenvDir) -and -not (Test-Path -LiteralPath $VenvPy)) {
        Warn "server\.venv exists but has no Scripts\python.exe (built on another OS). Recreating it."
        # A running python.exe or an on-access AV scan holding a handle makes
        # this throw - routine on a managed PC. Don't let it kill the script
        # with a raw .NET exception; say what to do about it.
        try {
            Remove-Item -LiteralPath $VenvDir -Recurse -Force
        } catch {
            Warn "Could not delete $VenvDir ($($_.Exception.Message))."
            Warn "Close anything using it (a running python.exe, a VS Code terminal, an"
            Warn "open Explorer window), or delete the folder by hand, then re-run this script."
            Write-Error "[setup] Stale server\.venv could not be removed - delete it manually and re-run." -ErrorAction Continue
            exit 1
        }
    }

    if (Test-Path -LiteralPath $VenvPy) {
        Say "Virtual environment already present: $VenvDir"
    } else {
        Say "Creating virtual environment in $VenvDir"
        $ErrorActionPreference = "Continue"
        & $py.Exe @pre -m venv $VenvDir 2>&1 | ForEach-Object { Write-Host "[setup]   $_" }
        $venvCode = $LASTEXITCODE
        $ErrorActionPreference = "Stop"
        if ($venvCode -ne 0) { Warn "venv creation exited with code $venvCode." }
    }

    if (Test-Path -LiteralPath $VenvPy) {
        $McpInterpreter = $VenvPy
        $RunPython      = $VenvPy
        Say "Venv interpreter: $VenvPy"
    } else {
        Warn "Could not create a virtual environment (ensurepip may be missing, or policy blocks it)."
        Warn "Falling back to a per-user install against the base interpreter."
        Warn "The MCP config will therefore point at $BasePython, not a venv."
        $UsedUserSite = $true
    }
}

# ------------------------------------------------------------------ install
if ($SkipVenv -and -not (Test-Path -LiteralPath $VenvPy)) {
    Say "-SkipVenv given - skipping the mcp SDK install. Ensure 'mcp' is importable by $McpInterpreter."
} else {
    if ($UsedUserSite) {
        Say "Installing the mcp SDK with: $BasePython -m pip install --user -r server\requirements.txt"
        $installArgs = @("-m", "pip", "install", "--disable-pip-version-check", "--user", "-r", $ReqFile)
    } else {
        Say "Installing the mcp SDK with: $McpInterpreter -m pip install -r server\requirements.txt"
        $installArgs = @("-m", "pip", "install", "--disable-pip-version-check", "-r", $ReqFile)
    }

    $ErrorActionPreference = "Continue"
    & $McpInterpreter @installArgs 2>&1 | ForEach-Object { Write-Host "[setup]   $_" }
    $pipCode = $LASTEXITCODE
    $ErrorActionPreference = "Stop"

    if ($pipCode -ne 0) {
        Warn "pip failed (exit $pipCode)."
        Warn ""
        Warn "Most likely cause on a federal PC is a TLS-inspecting proxy: pip sees a"
        Warn "certificate signed by the appliance and refuses. pip ships its own TLS"
        Warn "stack and its own CA bundle, so the Tls12 line at the top of this script"
        Warn "does NOT help it - you have to tell pip to trust the hosts:"
        Warn "  `"$McpInterpreter`" -m pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r `"$ReqFile`""
        Warn ""
        Warn "If the network needs an explicit proxy as well:"
        Warn "  `"$McpInterpreter`" -m pip install --proxy http://user:pass@proxyhost:port --trusted-host pypi.org --trusted-host files.pythonhosted.org -r `"$ReqFile`""
        Warn ""
        Warn "If pip is blocked from writing machine-wide, add --user to either command."
        Warn "See the 'pip cannot reach PyPI' row in README.md for the same table."
    } else {
        Say "mcp SDK installed."
    }
}

# ------------------------------------------------------- verify the mcp SDK
# THE gate. test_agent.py is stdlib-only by design (README says so), so it
# passes happily on a machine where the mcp SDK never installed - it cannot
# tell us whether the server will start. Importing 'mcp' with the exact
# interpreter that goes into mcp.json is the only check that means anything.
#
# The failure this catches: a stripped/policy-managed Python ships without
# ensurepip, so 'py -3 -m venv' still produces Scripts\python.exe but with no
# pip. The install above warns and scrolls past, and without this block the
# user gets a green "Setup complete." plus a copy-pasteable config pointing at
# an interpreter that cannot import mcp - then an opaque stdio error in VS
# Code, far from the real cause. Same story behind a TLS-intercepting proxy.
#
# $mcpOk: $true verified, $false verified-broken, $null deliberately unchecked.
$mcpOk = $null

if ($SkipVenv -and -not (Test-Path -LiteralPath $VenvPy)) {
    Warn "-SkipVenv given - the mcp SDK was NOT verified for this interpreter."
    Warn "Check it yourself before starting the server:"
    Warn "  `"$McpInterpreter`" -c `"import mcp`""
} else {
    Say "Verifying the mcp SDK imports with: $McpInterpreter"
    $ErrorActionPreference = "Continue"
    # "$_" per record, NOT Out-String: with 2>&1 each stderr line arrives as an
    # ErrorRecord, and the default formatter buries Python's traceback under
    # PowerShell's own "NativeCommandError / CategoryInfo / At line:" noise.
    # ToString() on the record is just the text Python actually printed.
    $mcpOut = @(& $McpInterpreter -c "import mcp,sys;print(mcp.__file__ or 'ok')" 2>&1 |
                ForEach-Object { "$_" } | Where-Object { $_.Trim() -ne "" })
    $mcpCode = $LASTEXITCODE
    $ErrorActionPreference = "Stop"
    $mcpOk = ($mcpCode -eq 0)

    if ($mcpOk) {
        $where = @($mcpOut | Select-Object -Last 1)[0]
        Say "mcp SDK OK -> $("$where".Trim())"
    } else {
        Warn "The mcp SDK is NOT importable by $McpInterpreter (exit $mcpCode)."
        foreach ($line in $mcpOut) { Warn "  $line" }
        Warn ""
        Warn "The MCP server cannot start until this import works. To fix it:"
        Warn "  1. Retry the install, adding --trusted-host if a proxy inspects TLS:"
        Warn "     `"$McpInterpreter`" -m pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r `"$ReqFile`""
        Warn "  2. If pip itself is missing (no ensurepip in this Python), install a"
        Warn "     full Python 3.10+ from python.org or the Company Portal and re-run."
        Warn "  3. If you already have mcp in a different environment, point setup at it:"
        Warn "     `$env:SPLA_PYTHON = 'C:\path\to\that\python.exe'  ; .\setup.ps1"
        Warn "Then re-run: powershell -ExecutionPolicy Bypass -File .\setup.ps1"
    }
}

# --------------------------------------------------------------------- test
Say "Running the data-logic check: server\test_agent.py"
$ErrorActionPreference = "Continue"
& $RunPython $TestScript 2>&1 | ForEach-Object { Write-Host "[setup]   $_" }
$testCode = $LASTEXITCODE
$ErrorActionPreference = "Stop"

if ($testCode -eq 0) {
    Say "test_agent.py PASSED."
} else {
    Warn "test_agent.py FAILED (exit $testCode). Read the output above before using the tool."
}

if ($Test) {
    Say "Running the full MCP round-trip: server\demo.py"
    $ErrorActionPreference = "Continue"
    & $RunPython $DemoScript 2>&1 | ForEach-Object { Write-Host "[setup]   $_" }
    $demoCode = $LASTEXITCODE
    $ErrorActionPreference = "Stop"

    if ($demoCode -eq 0) {
        Say "demo.py PASSED - the MCP server answers over stdio."
    } else {
        Warn "demo.py FAILED (exit $demoCode). This usually means the mcp SDK is not installed for $RunPython."
    }
}

# ------------------------------------------------------------- the mcp block
# JSON needs every backslash doubled.
$jsonPy  = $McpInterpreter.Replace('\', '\\')
$jsonSrv = $ServerScript.Replace('\', '\\')
$jsonExp = $ExportDir.Replace('\', '\\')
# The VS Code workspace is normally the folder ABOVE this one (C:\SP), since
# sharepoint-list-agent is a subfolder of the SP repo. Fall back to $Root if it
# somehow sits at a drive root.
$WorkspaceDir = Split-Path -Parent $Root
if (-not $WorkspaceDir) { $WorkspaceDir = $Root }
$VsCodeMcp = Join-Path $WorkspaceDir ".vscode\mcp.json"

$block = @"
{
  "servers": {
    "sharepoint-list": {
      "type": "stdio",
      "command": "$jsonPy",
      "args": ["$jsonSrv"],
      "env": {
        "SPLA_EXPORT_DIR": "$jsonExp"
      }
    }
  }
}
"@

Write-Host ""
Say "=============================================================="
Say " GitHub Copilot in VS Code - paste this into your workspace at:"
Say "   $VsCodeMcp"
Say " (create the .vscode folder if it is not there; if mcp.json already"
Say "  exists, merge the 'sharepoint-list' entry into its 'servers' object)"
Say "=============================================================="
Write-Host ""
Write-Host $block
Write-Host ""

# Never hand over a config block that looks fine when we know it is not.
if ($false -eq $mcpOk) {
    Warn "=============================================================="
    Warn " WARNING: the mcp SDK is NOT importable by the interpreter"
    Warn " named in the block above. Pasting this config will not work"
    Warn " yet - VS Code will fail to start the server with an opaque"
    Warn " stdio error. Fix the import first (guidance above), then"
    Warn " re-run .\setup.ps1 and use the block it prints."
    Warn "=============================================================="
    Write-Host ""
} elseif ($null -eq $mcpOk) {
    Warn "Note: -SkipVenv was given, so the mcp SDK was NOT verified for the"
    Warn "interpreter above. Confirm it before starting the server:"
    Warn "  `"$McpInterpreter`" -c `"import mcp`""
    Write-Host ""
}

Say "Then in VS Code: Ctrl+Shift+P > 'MCP: List Servers' > sharepoint-list > Start,"
Say "and switch Copilot Chat to Agent mode so the four tools are offered."
Say "A copy of this shape also lives at config\mcp.windows.json."
Write-Host ""
Say "--------------------------------------------------------------"
Say " Getting data in (browser side - one paste, no install):"
Say "   Open your SharePoint list in Edge, press F12 > Console, paste all of"
Say "   browser\read-list.js, then run:  await SPLA.export('My List')"
Say "   Save/move the downloaded JSON into: $ExportDir"
Say "   (or run  await SPLA.folder()  once and pick that folder, so exports"
Say "    land there directly)"
Say "--------------------------------------------------------------"
Write-Host ""

$failures = @()
if ($testCode -ne 0)   { $failures += "test_agent.py failed (exit $testCode)" }
if ($false -eq $mcpOk) { $failures += "the mcp SDK is not importable by $McpInterpreter" }

if ($failures.Count -eq 0) {
    if ($null -eq $mcpOk) {
        Say "Setup finished (-SkipVenv: the mcp SDK was not verified)."
    } else {
        Say "Setup complete."
    }
} else {
    Say "Setup finished with failures - see the warnings above:"
    foreach ($f in $failures) { Say "  - $f" }
    exit 1
}
