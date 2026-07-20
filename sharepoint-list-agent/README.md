# SharePoint List Agent

Give a coding agent (GitHub Copilot, Claude, etc.) read access to your
SharePoint **List** data in a locked-down federal (GCC / GCC High) tenant —
**without** Microsoft Graph, an app registration, an admin-consent prompt, or a
new login.

## Why this shape

The normal programmatic path — Graph with an access token — is very likely
gated in a DOL-style tenant: app-registration rules, admin consent for
`Sites.Read.All`, and Conditional Access all sit in the way. So this project
avoids tokens entirely and splits the work across the trust boundary:

- **The login-requiring step stays in the browser**, where your org SSO already
  works. A pasted script calls SharePoint's own REST API from inside the tab
  you're signed into, riding the session cookie. No token, no consent.
- **The agent only ever touches a local file** — a plain JSON export that has
  already passed through a redaction filter. It makes zero network calls.

```
  SharePoint (browser tab, your SSO)                Your machine
  ┌───────────────────────────┐    download    ┌────────────────────────────┐
  │ browser/read-list.js       │  ───JSON───▶   │ config/exports/*.json      │
  │  REST via session cookie   │                │            │               │
  └───────────────────────────┘                │            ▼               │
                                                │ server/  (MCP, stdio)      │
                                                │  redact → 4 read-only tools│
                                                │            │               │
                                                │            ▼               │
                                                │   coding agent (Copilot)   │
                                                └────────────────────────────┘
```

## Layout

| Path | What it is |
|---|---|
| `browser/read-list.js` | Paste into DevTools on your SharePoint site; exports a List to JSON. |
| `server/sharepoint_list_mcp.py` | The MCP stdio server — four read-only tools. |
| `server/store.py` | Pure data logic (load / filter / project / rank). No deps. |
| `server/policy.py` | Redaction pass (SSN / EIN / card / phone / id-run). |
| `server/test_agent.py` | Runnable check for store + policy (`python test_agent.py`). |
| `server/demo.py` | MCP stdio client that spawns the server and exercises all four tools. |
| `server/graph_export.py` | Device-code Graph exporter — same JSON shape, for tenants where Graph is allowed. |
| `server/requirements.txt` | Just the `mcp` SDK. |
| `config/exports/` | Drop exported JSON here. `sample-list.json` ships for a smoke test. |
| `config/mcp.example.json` | Generic MCP registration example (uses `python3`, i.e. Linux/WSL). |
| `config/mcp.windows.json` | Ready-made **Windows** registration — `server\.venv\Scripts\python.exe`, Windows paths hardcoded to `C:\SP`. |
| `setup.ps1` | **Windows setup.** Finds an interpreter, builds `server\.venv`, installs `mcp`, prints the `.vscode\mcp.json` block with resolved paths. |
| `docs/SPEC.md` | Full replication spec (hand to an agent to rebuild this). |
| `docs/WORKFLOW.md` | Step-by-step operating instructions. |

## Quickstart

1. **Export a List.** On your SharePoint site, open DevTools (F12) → Console,
   paste all of `browser/read-list.js`, then:
   ```js
   SPLA.lists()                  // see your lists
   await SPLA.folder()           // pick config\exports\ once — Edge/Chrome
   await SPLA.export('My List')  // writes sp_My_List_YYYYMMDD.json straight in
   ```
   `SPLA.folder()` opens a folder picker (File System Access API). Point it at
   `C:\SP\sharepoint-list-agent\config\exports` and **every** later
   `SPLA.export` in that tab writes directly there — no download, no moving
   files. The grant lasts one page-load (a browser rule — re-pick after a
   refresh), and any write failure falls back to a normal download.
2. **Stage it.** Nothing to do if you used `SPLA.folder()`. Otherwise move the
   downloaded JSON into `config/exports/`.
3. **Install + register.** Install the `mcp` SDK, then register the server with
   your agent — on Windows, `setup.ps1` does both (next section).
4. **Ask the agent.** It now has `sharepoint_list_status`,
   `sharepoint_list_import`, `sharepoint_list_items`, `sharepoint_list_related`.

## Windows work PC (PowerShell + GitHub Copilot)

**This is the primary path** — no WSL, no bash, Windows PowerShell 5.1 is fine.

Two rules for this machine: the interpreter is **`python`** or the launcher
**`py -3`**, never `python3` (on Windows that name is usually an App Execution
Alias stub that just opens the Microsoft Store); and the venv interpreter is
**`server\.venv\Scripts\python.exe`**, not `.venv/bin/python`. Note the venv
lives inside `server\`, **not** at the project root — `setup.ps1` builds it
there and every command below assumes it.

### 1. Get the repo

```powershell
powershell -ExecutionPolicy Bypass -File .\update-sp.ps1 -Dest "C:\SP"
```

First run clones, later runs pull. The script lives at the repo root
(`C:\SP\update-sp.ps1`), so on a brand-new machine save that one file somewhere
first and run it from there; afterwards run it in place. Everything below
assumes the project ends up at `C:\SP\sharepoint-list-agent`.

### 2. Run setup

```powershell
cd C:\SP\sharepoint-list-agent
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

It probes for a usable interpreter, creates the venv at
`C:\SP\sharepoint-list-agent\server\.venv`, installs the `mcp` SDK into it, and
prints the `.vscode\mcp.json` block to paste. If Python is absent or pip is
blocked it stops and tells you what to fix rather than half-finishing.

Sanity check by hand:

```powershell
py -3 --version
C:\SP\sharepoint-list-agent\server\.venv\Scripts\python.exe --version
```

### 3. Export a List from Edge

On the SharePoint site, `F12` → **Console**, paste all of
`C:\SP\sharepoint-list-agent\browser\read-list.js`, then:

```js
SPLA.lists()
await SPLA.folder()                  // pick C:\SP\sharepoint-list-agent\config\exports
await SPLA.export('Case Tracker')
```

The JSON lands in `config\exports\` directly. Re-pick the folder after any page
refresh — the browser only grants it for the current page-load.

### 4. Register with GitHub Copilot

Copilot reads `.vscode\mcp.json` from the folder you open in VS Code — with the
project at `C:\SP\sharepoint-list-agent`, that is `C:\SP\.vscode\mcp.json`.

**Prefer the block `setup.ps1` just printed.** It carries the real resolved
paths for your machine, including the interpreter it actually ended up with.
`config/mcp.windows.json` hardcodes `C:\SP`; if you relocated the checkout with
`update-sp.ps1 -Dest` or `setup.ps1 -Root`, every path in that file is wrong and
Copilot's only symptom is a server that fails to start.

**If `.vscode\mcp.json` already exists, do not copy over it.** `Copy-Item`
replaces the destination silently — any other MCP servers you registered are
gone, with no prompt and no backup. Back it up, then merge:

```powershell
Copy-Item C:\SP\.vscode\mcp.json C:\SP\.vscode\mcp.json.bak
```

Open `mcp.json` and add this one entry inside the existing `"servers"` object,
comma-separated from what is already there:

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

Swap `C:\\SP` for wherever the repo actually lives. Backslashes are doubled
because it is JSON. Two rules for `command`: it must be a real interpreter
executable — never `python3` (the Store alias stub) and never `py` (a launcher
shim, not an interpreter). And if `setup.ps1` reported that it could not build a
venv and fell back to `pip install --user mcp`, point `command` at the base
interpreter it named instead of the `server\.venv` path.

**Only if you have no `.vscode\mcp.json` yet** is copying the whole file safe:

```powershell
New-Item -ItemType Directory -Force -Path C:\SP\.vscode | Out-Null
if (Test-Path C:\SP\.vscode\mcp.json) {
    Write-Host "mcp.json exists - back it up and merge the entry by hand (above)"
} else {
    Copy-Item C:\SP\sharepoint-list-agent\config\mcp.windows.json C:\SP\.vscode\mcp.json
}
```

Open `C:\SP` in VS Code, confirm every path inside `mcp.json` points at a file
that exists, then start the server from the MCP view (or reload the window). Use
`config/mcp.example.json` **only** on Linux/WSL — it calls `python3`.

Claude Desktop uses the same inner shape under `mcpServers` instead of
`servers`, in `claude_desktop_config.json`.

### 5. Verify

```powershell
cd C:\SP\sharepoint-list-agent
server\.venv\Scripts\python.exe server\test_agent.py    # store + policy, no mcp needed
server\.venv\Scripts\python.exe server\demo.py          # real MCP stdio round-trip
server\.venv\Scripts\python.exe server\demo.py "My List"
```

Both scripts resolve their own paths, so the project root is the right place to
run them from. `test_agent.py` is stdlib-only and works with a plain `python`
too; `demo.py` needs the `mcp` SDK, hence the venv.

### When the machine fights back

| Symptom | Fix |
|---|---|
| `python3` opens the Microsoft Store | Use `py -3` (or `python`). Nothing here should call `python3`. |
| "running scripts is disabled on this system" | Launch as `powershell -ExecutionPolicy Bypass -File .\script.ps1`. If that *still* fails, Group Policy is setting the policy — see the next row. |
| …and `-ExecutionPolicy Bypass` changes nothing | `-ExecutionPolicy` only sets the **Process** scope, and `MachinePolicy` / `UserPolicy` outrank it. Run `Get-ExecutionPolicy -List`: if either policy scope is anything but `Undefined`, no command-line flag can override it. Ask IT for an exception, or run the steps by hand (below). |
| Scripts refused right after the ZIP-download fallback | Files extracted from a downloaded `.zip` carry a Mark-of-the-Web stream and are blocked under `RemoteSigned`. Clear it: `Get-ChildItem C:\SP -Recurse -Filter *.ps1 \| Unblock-File`. |
| `Invoke-WebRequest` fails on TLS / SSL | PowerShell 5.1 defaults to TLS 1.0. Run `[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12` first. `update-sp.ps1` already does this internally. |
| `pip` can't reach PyPI | pip has its own TLS stack — the line above does **not** help it. Try `python -m pip install --user -r server\requirements.txt`, add `--proxy http://host:port`, or for an intercepting proxy `--trusted-host pypi.org --trusted-host files.pythonhosted.org`. |
| No Python on the machine at all | The MCP server can't run. The exports are still plain JSON — open `C:\SP\sharepoint-list-agent\config\exports` in VS Code and Copilot can read the files directly, just without the redaction and filtering tools. |

### If you can't run `setup.ps1` at all

Group-policy-enforced execution policy blocks the script, not Python. Do the
three things it would have done yourself:

```powershell
cd C:\SP\sharepoint-list-agent
py -3 -m venv server\.venv
server\.venv\Scripts\python.exe -m pip install -r server\requirements.txt
server\.venv\Scripts\python.exe -c "import mcp; print('mcp ok')"
```

Then register the server with the JSON entry in step 4, using the paths above.

## Dev machine (Linux/WSL)

Same project, `python3` and POSIX venv paths. Verify the data logic without
SharePoint or the `mcp` install:

```bash
cd server && python3 test_agent.py
```

See the **server** working end-to-end (real MCP stdio round-trip) against
whatever is in `config/exports/` — ships working on the sample:

```bash
cd server
python3 -m venv .venv || uv venv .venv     # uv if python-venv is unavailable
.venv/bin/pip install mcp || uv pip install --python .venv/bin/python mcp
.venv/bin/python demo.py                    # auto-picks the first export
.venv/bin/python demo.py "My List"          # or target one
```

Register with `config/mcp.example.json` here — copy it to `.vscode/mcp.json`
(VS Code / Copilot use the `servers` key; Claude Desktop uses `mcpServers` in
`claude_desktop_config.json`, same inner shape). Replace the absolute paths, or
keep `${workspaceFolder}` if your host expands it. It calls a bare `python3`, so
the `mcp` SDK must be importable from it — point `command` at
`.../server/.venv/bin/python` instead if you installed into the venv above.
`config/mcp.windows.json` is the Windows counterpart. Both files are
schema-clean (only `servers`), so they can be dropped in without VS Code
flagging unknown keys; the same merge caution applies if `mcp.json` already
exists.

## The tools

| Tool | Purpose |
|---|---|
| `sharepoint_list_status` | Which exports are on disk (row/column counts, export dir, active redaction rules). |
| `sharepoint_list_import` | One List's columns (display + internal name + type) and row count. |
| `sharepoint_list_items` | Rows, with column pick, case-insensitive text filter, and paging. **Redacted by default.** |
| `sharepoint_list_related` | Rank rows against a topic (e.g. an Outlook subject line) so the right List rows surface. **Redacted by default.** |

## Safety

- **Read-only.** No tool writes to SharePoint; the server makes no network calls.
- **Redaction at the boundary.** SSNs (including spaced/dotted forms), EINs,
  grouped credit-card numbers, US phone numbers, and 7+-digit id runs are
  masked to their last 4 in `policy.py` — the last point on the machine before
  data reaches the model. Dates, dollar amounts, and decimals are preserved.
  Widen it there; it's the whole policy surface.
- **What is deliberately kept.** Emails and personal names are **not** masked —
  cross-referencing people between Outlook and List rows is the tool's purpose.
  International phone formats are not covered; redaction is best-effort pattern
  matching, not DLP.
- **Least data.** You choose which List to export and which columns the agent
  sees. Real exports are git-ignored (`config/exports/*.json`) so tenant data is
  never committed — only the synthetic sample is tracked.
- **No new trust.** You only ever read what your own SharePoint permissions
  already allow, through the session your browser already holds.

## Where Graph is allowed

`server/graph_export.py` is the Graph-side exporter: a stdlib-only device-code
sign-in (the Microsoft Graph Command Line Tools client id) that writes the
**same** canonical JSON into `config/exports/`, so the server and tools stay
unchanged. Flags: `--site`, `--list`, `--all`, `--scan`, `--request-code`,
`--poll`. The token is cached at `~/.cache/spla_token.json` (mode 0600) and
refreshed silently.

It targets **commercial-cloud endpoints only** — a GCC High tenant would need
`login.microsoftonline.us` / `graph.microsoft.us` via the `AUTH` / `GRAPH`
constants at the top of the file. The browser reader remains the path for
locked-down tenants.

Before trying this at work, confirm three things: (1) you can register an app,
(2) `Sites.Read.All` is grantable under user consent, (3) Conditional Access
allows the flow.

---
*Note: this is the one multi-file, pip-dependent member of the SP repo. The
other artifacts are single-file and zero-dependency by design; this tool needs a
local process the agent can talk to, so it is isolated in its own folder with
its own `.gitignore`.*
