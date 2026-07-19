# Workflow

The full loop, from a List in SharePoint to the agent cross-referencing it
against your mail.

## 1. Export a List (browser, one-time per refresh)

1. Open your SharePoint site in Edge/Chrome, signed in as normal.
2. `F12` → **Console**.
3. Paste the entire contents of `browser/read-list.js` and press Enter.
4. Discover and pull:
   ```js
   SPLA.lists()                        // prints a table of non-hidden lists
   await SPLA.export('Case Tracker')   // downloads sp_Case_Tracker_YYYYMMDD.json
   ```
   - `SPLA.read('Case Tracker')` returns the object without downloading, if you
     just want to inspect it in the console first.
   - People and lookup columns come back as readable strings
     (`Jane Doe <jane.doe@dol.gov>`). Large lists page automatically and back
     off on throttling (429 / 503).

> Cross-site note: you can export any list **on the same tenant host** you're
> viewing (same-origin cookie). A different tenant needs its own paste on a tab
> signed into that tenant.

## 2. Stage the export

Move the downloaded file into `config/exports/`. Re-exporting later just
overwrites — the server always reads whatever is currently on disk.

```
config/exports/
  sample-list.json          # ships with the repo (synthetic)
  sp_Case_Tracker_20260719.json
```

## 3. Install and register the server

```bash
pip install -r server/requirements.txt          # only the `mcp` SDK
```

Add the server to your agent. Copy `config/mcp.example.json` into place and fix
the paths:

- **VS Code / GitHub Copilot** → `.vscode/mcp.json` (uses the `servers` key).
- **Claude Desktop** → `claude_desktop_config.json` (rename the key to
  `mcpServers`; inner shape is identical).

Set `SPLA_EXPORT_DIR` in the `env` block if your exports live elsewhere;
otherwise the server defaults to `../config/exports` next to the script.

## 4. Use it from the agent

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

## 5. Verify (no SharePoint needed)

```bash
cd server && python test_agent.py     # exercises store + policy on the sample
python -m py_compile sharepoint_list_mcp.py store.py policy.py
```

## Redaction

Rows returned by `sharepoint_list_items` and `sharepoint_list_related` are
redacted by default. To see raw values for a trusted local task, pass
`redact: false` on that call. To change what gets masked, edit `_RULES` in
`server/policy.py` — it's the whole policy surface, applied to every string in
every returned row.
