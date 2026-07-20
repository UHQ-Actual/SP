# Workflow

The full loop, from a List in SharePoint to the agent cross-referencing it
against your mail.

**The Windows work PC (PowerShell + GitHub Copilot) is the primary path** and is
what steps 1–5 below describe. A Linux/WSL variant of each command follows at
the end. On Windows the interpreter is `python` or `py -3` — **never `python3`**,
which is normally an App Execution Alias stub that opens the Microsoft Store —
and the venv interpreter is `server\.venv\Scripts\python.exe`. The venv lives
inside `server\`, **not** at the project root.

## 1. Set up the machine (once)

```powershell
powershell -ExecutionPolicy Bypass -File .\update-sp.ps1 -Dest "C:\SP"
cd C:\SP\sharepoint-list-agent
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

`update-sp.ps1` clones on the first run and pulls after that. `setup.ps1` finds
an interpreter, creates the venv at `server\.venv`, installs the `mcp` SDK, and
prints the `.vscode\mcp.json` block for Copilot with the real resolved paths. It
fails loudly with the actual remedy if Python is missing or pip is blocked.

Launch every script as `powershell -ExecutionPolicy Bypass -File .\…`. A bare
`.\setup.ps1` is refused wherever the execution policy is `Restricted` (the
Windows client default), and the `-File` form clears that case.

**It does not clear every case.** `-ExecutionPolicy` sets only the **Process**
scope, and the `MachinePolicy` and `UserPolicy` scopes — the ones Group Policy
writes on a managed federal PC — take precedence over it. If the launch is still
refused with the same "running scripts is disabled" message, check which scope
is responsible:

```powershell
Get-ExecutionPolicy -List
```

If `MachinePolicy` or `UserPolicy` is anything other than `Undefined`, no
command-line flag will override it. Either request an exception from IT, or skip
the script and run what it would have run:

```powershell
cd C:\SP\sharepoint-list-agent
py -3 -m venv server\.venv
server\.venv\Scripts\python.exe -m pip install -r server\requirements.txt
server\.venv\Scripts\python.exe -c "import mcp; print('mcp ok')"
```

Separately: if `update-sp.ps1` fell back to downloading a ZIP (no `git` on the
box), the extracted `.ps1` files carry a Mark-of-the-Web stream and are refused
under `RemoteSigned` until cleared:

```powershell
Get-ChildItem C:\SP -Recurse -Filter *.ps1 | Unblock-File
```

## 2. Export a List (browser, once per refresh)

1. Open your SharePoint site in Edge, signed in as normal.
2. `F12` → **Console**.
3. Paste the entire contents of `C:\SP\sharepoint-list-agent\browser\read-list.js`
   and press Enter. The banner lists what you can call:
   ```
   await SPLA.folder()          pick config\exports\ once — exports write straight there
   SPLA.lists()                 list non-hidden lists/libraries
   await SPLA.read('My List')   read -> {web,list,columns,items}
   await SPLA.export('My List') read + write/download the JSON
   ```
4. Point the reader at the export folder, then pull:
   ```js
   SPLA.lists()                        // prints a table of non-hidden lists
   await SPLA.folder()                 // pick C:\SP\sharepoint-list-agent\config\exports
   await SPLA.export('Case Tracker')   // writes sp_Case_Tracker_YYYYMMDD.json into it
   ```
   - `SPLA.folder()` uses the File System Access API (**Edge/Chrome only**). Once
     granted, every subsequent `SPLA.export` in that tab writes straight into
     that folder — no download, no move step. The grant lasts **one page-load**
     (a browser rule; re-pick after a refresh), and any write failure falls back
     to a normal download and clears the folder. Browsers without the API just
     download as before.
   - `SPLA.read('Case Tracker')` returns the object without writing anything, if
     you just want to inspect it in the console first.
   - People and lookup columns come back as readable strings
     (`Jane Doe <jane.doe@dol.gov>`). Large lists page automatically and back
     off on throttling (429 / 503).

> Cross-site note: you can export any list **on the same tenant host** you're
> viewing (same-origin cookie). A different tenant needs its own paste on a tab
> signed into that tenant.

## 3. Stage the export

If you used `SPLA.folder()`, this step is already done. Otherwise move the
downloaded file from `%USERPROFILE%\Downloads` into
`C:\SP\sharepoint-list-agent\config\exports`:

```powershell
Move-Item $env:USERPROFILE\Downloads\sp_Case_Tracker_*.json `
          C:\SP\sharepoint-list-agent\config\exports\
```

Re-exporting later just overwrites — the server always reads whatever is
currently on disk.

```
config\exports\
  sample-list.json          # ships with the repo (synthetic)
  sp_Case_Tracker_20260719.json
```

## 4. Register the server with Copilot

Copilot reads `.vscode\mcp.json` from the folder you open in VS Code — with the
project at `C:\SP\sharepoint-list-agent`, that file is `C:\SP\.vscode\mcp.json`.

**Use the block `setup.ps1` printed.** It has the paths already resolved for
this machine, including the interpreter it actually settled on (a venv python,
or the base interpreter if the venv could not be built). `config/mcp.windows.json`
is the fallback for when you no longer have that output — it hardcodes `C:\SP`,
so if the checkout lives anywhere else (`update-sp.ps1 -Dest`, `setup.ps1
-Root`) every path in it must be edited to match, or the server just fails to
start with no useful message.

### If `.vscode\mcp.json` already exists

Do **not** copy over it. `Copy-Item` replaces the destination silently, and any
other MCP servers registered there are lost with no prompt and no backup. Back
it up first:

```powershell
Copy-Item C:\SP\.vscode\mcp.json C:\SP\.vscode\mcp.json.bak
```

Then open it and merge this single entry into the existing `"servers"` object
(comma-separated from the entries already present):

```json
"sharepoint-list": {
  "type": "stdio",
  "command": "C:\\SP\\sharepoint-list-agent\\server\\.venv\\Scripts\\python.exe",
  "args": ["C:\\SP\\sharepoint-list-agent\\server\\sharepoint_list_mcp.py"],
  "env": {
    "SPLA_EXPORT_DIR": "C:\\SP\\sharepoint-list-agent\\config\\exports"
  }
}
```

Replace `C:\\SP` with the real checkout path. Backslashes are doubled because
it is JSON. `command` must name a real interpreter executable — never `python3`
(the Microsoft Store alias stub) and never `py` (a launcher shim). If
`setup.ps1` warned that it could not create a venv and installed with
`pip install --user` instead, use the base interpreter it reported in place of
the `server\.venv` path.

### If there is no `.vscode\mcp.json` yet

Copying the whole file is safe in that case only:

```powershell
New-Item -ItemType Directory -Force -Path C:\SP\.vscode | Out-Null
if (Test-Path C:\SP\.vscode\mcp.json) {
    Write-Host "mcp.json exists - back it up and merge the entry by hand (above)"
} else {
    Copy-Item C:\SP\sharepoint-list-agent\config\mcp.windows.json C:\SP\.vscode\mcp.json
}
```

Either way, open `mcp.json` afterwards and confirm `command` and `args` name
files that actually exist, then start the server from the MCP view or reload the
window.

- **VS Code / GitHub Copilot** → `.vscode\mcp.json` (uses the `servers` key).
- **Claude Desktop** → `claude_desktop_config.json` (rename the key to
  `mcpServers`; inner shape is identical).

Set `SPLA_EXPORT_DIR` in the `env` block if your exports live elsewhere;
otherwise the server defaults to `../config/exports` next to the script.

`config/mcp.example.json` is the generic example and calls `python3` — it is the
Linux/WSL shape, not this one.

## 5. Use it from the agent

Typical asks, and the tool each drives:

| You ask the agent… | It calls |
|---|---|
| "What SharePoint data do you have?" | `sharepoint_list_status` |
| "What columns are in the Case Tracker list?" | `sharepoint_list_import` |
| "Show open cases for Bright Cleaners." | `sharepoint_list_items` (filter) |
| "This email subject is *'MW overtime — Acme'*. Which list rows relate?" | `sharepoint_list_related` |

Combined with the agent's Outlook access, the pattern is: read the mail subject
→ `sharepoint_list_related` to find the matching List rows → pull specifics with
`sharepoint_list_items`.

## Verify (no SharePoint needed)

From `C:\SP\sharepoint-list-agent`:

```powershell
server\.venv\Scripts\python.exe server\test_agent.py     # store + policy on the sample
server\.venv\Scripts\python.exe -m py_compile server\sharepoint_list_mcp.py server\store.py server\policy.py
server\.venv\Scripts\python.exe server\demo.py           # real MCP stdio round-trip
```

Both scripts resolve their own location, so the project root is the right place
to run them from. `test_agent.py` is stdlib-only and runs under a plain `python`
too; `demo.py` needs the `mcp` SDK, so it must be the venv interpreter.

## Linux/WSL variant

The dev-machine equivalents of the same steps — `python3`, POSIX venv paths:

```bash
cd server
python3 test_agent.py                    # verify — stdlib only, no venv needed
python3 -m py_compile sharepoint_list_mcp.py store.py policy.py

python3 -m venv .venv                    # step 1 — the venv lives in server/
.venv/bin/pip install -r requirements.txt   # only the `mcp` SDK
.venv/bin/python demo.py                 # real MCP stdio round-trip
```

The venv sits at `server/.venv`, mirroring `server\.venv` on Windows — the same
place `setup.ps1` puts it.

Register with `config/mcp.example.json` (it calls `python3`) instead of
`config/mcp.windows.json`, at `.vscode/mcp.json`. Steps 2 and 5 — the browser
export and the agent usage — are identical on both machines; only the staging
path differs (`config/exports/` rather than `config\exports\`).

## Redaction

Rows returned by `sharepoint_list_items` and `sharepoint_list_related` are
redacted by default: SSNs (including spaced/dotted forms), EINs, grouped
credit-card numbers, US phone numbers, and 7+-digit id runs are masked to
their last 4; dates, dollar amounts, and decimals are preserved. Emails and
personal names are deliberately **not** masked — cross-referencing people
between Outlook and List rows is the tool's purpose. International phone
formats are not covered; redaction is best-effort pattern matching, not DLP.

To see raw values for a trusted local task, pass `redact: false` on that call.
To change what gets masked, edit `_RULES` in `server/policy.py` — it's the
whole policy surface, applied to every string in every returned row.
