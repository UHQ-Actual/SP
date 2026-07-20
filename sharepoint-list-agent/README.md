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
| `server/policy.py` | Redaction pass (SSN / phone / case-id). |
| `server/test_agent.py` | Runnable check for store + policy (`python test_agent.py`). |
| `server/requirements.txt` | Just the `mcp` SDK. |
| `config/exports/` | Drop exported JSON here. `sample-list.json` ships for a smoke test. |
| `config/mcp.example.json` | Example MCP registration for your agent. |
| `docs/SPEC.md` | Full replication spec (hand to an agent to rebuild this). |
| `docs/WORKFLOW.md` | Step-by-step operating instructions. |

## Quickstart

1. **Export a List.** On your SharePoint site, open DevTools (F12) → Console,
   paste all of `browser/read-list.js`, then:
   ```js
   SPLA.lists()                 // see your lists
   await SPLA.export('My List') // download sp_My_List_YYYYMMDD.json
   ```
2. **Stage it.** Move the downloaded JSON into `config/exports/`.
3. **Install + register.** `pip install -r server/requirements.txt`, then add
   `config/mcp.example.json` to your agent's MCP config (fix the paths).
4. **Ask the agent.** It now has `sharepoint_list_status`,
   `sharepoint_list_import`, `sharepoint_list_items`, `sharepoint_list_related`.

Verify the data logic without SharePoint or the `mcp` install:
```bash
cd server && python test_agent.py
```

See the **server** working end-to-end (real MCP stdio round-trip) against
whatever is in `config/exports/` — ships working on the sample:
```bash
cd server
python -m venv .venv || uv venv .venv      # uv if python-venv is unavailable
.venv/bin/pip install mcp || uv pip install --python .venv/bin/python mcp
.venv/bin/python demo.py                    # auto-picks the first export
.venv/bin/python demo.py "My List"          # or target one
```

## The tools

| Tool | Purpose |
|---|---|
| `sharepoint_list_status` | Which exports are on disk (row/column counts, export dir, active redaction rules). |
| `sharepoint_list_import` | One List's columns (display + internal name + type) and row count. |
| `sharepoint_list_items` | Rows, with column pick, case-insensitive text filter, and paging. **Redacted by default.** |
| `sharepoint_list_related` | Rank rows against a topic (e.g. an Outlook subject line) so the right List rows surface. **Redacted by default.** |

## Safety

- **Read-only.** No tool writes to SharePoint; the server makes no network calls.
- **Redaction at the boundary.** SSN-like numbers, phones, and 7–10 digit
  case/ID runs are masked in `policy.py` — the last point on the machine before
  data reaches the model. Widen it there; it's the whole policy surface.
- **Least data.** You choose which List to export and which columns the agent
  sees. Real exports are git-ignored (`config/exports/*.json`) so tenant data is
  never committed — only the synthetic sample is tracked.
- **No new trust.** You only ever read what your own SharePoint permissions
  already allow, through the session your browser already holds.

## If IT ever clears Graph

The upgrade is additive: an MSAL device-code or interactive sign-in for
`Sites.Read.All` that writes the **same** JSON shape into `config/exports/`. The
server and tools stay unchanged. Before building it, confirm three things —
(1) you can register an app, (2) `Sites.Read.All` is grantable under user
consent, (3) Conditional Access allows the flow — plus which cloud you're on
(commercial vs GCC High), since the endpoints differ. See `docs/SPEC.md`.

---
*Note: this is the one multi-file, pip-dependent member of the SP repo. The
other artifacts are single-file and zero-dependency by design; this tool needs a
local process the agent can talk to, so it is isolated in its own folder with
its own `.gitignore`.*
