# Replication spec

Hand this to a coding agent to rebuild the project from scratch. It captures the
contract each half must honor; the exact code is secondary.

## Constraint that dictates the design

Microsoft Graph in a locked-down federal (GCC / GCC High) tenant is very likely
blocked by app-registration rules, admin-consent requirements, and Conditional
Access. So do **not** rely on a Graph token. Read Lists through the browser
session the user is already signed into (same-origin cookie auth). The
sensitive, login-requiring step stays in the browser; the agent only ever sees a
local file that has passed a redaction filter.

## Canonical export schema

Both halves agree on this shape. Extra keys are ignored by the server.

```json
{
  "web": "https://tenant.sharepoint.com/sites/foo",
  "list": "My List",
  "listId": "guid-or-null",
  "exportedAt": "ISO-8601",
  "columns": [ { "name": "Title", "internal": "Title", "type": "Text" } ],
  "items":   [ { "Id": 1, "Title": "…", "Manager": "Jane Doe <jane@…>" } ]
}
```

`items` are keyed by column **display name**; people/lookup values are flattened
to readable strings.

## Half 1 — Browser reader (JavaScript, paste into DevTools)

- Site URL from `window._spPageContextInfo.webAbsoluteUrl` (fall back to
  `location.origin` with a warning).
- Enumerate: `GET {web}/_api/web/lists?$filter=Hidden eq false` →
  Title / ItemCount / BaseTemplate.
- Schema: `GET {web}/_api/web/lists/getByTitle('Name')/fields` → keep non-hidden,
  non-`_`-prefixed fields; skip `Computed/Attachments/ContentType/Guid/...`;
  keep custom fields plus the useful system ones
  (`Title, Created, Modified, Author, Editor`). Record `InternalName`,
  `TypeAsString`.
- Items: build `$select` from scalar internal names, `$expand` User and Lookup
  columns (`Field/Title`, plus `Field/EMail` for people). On query error, **fall
  back** to a plain fetch of scalars plus the raw `<Field>Id` lookup ids.
- Paging: follow `odata.nextLink` (also handle `@odata.nextLink` / `d.__next`).
- Headers: `Accept: application/json;odata=nometadata`,
  `credentials: same-origin`.
- Throttling: on **429 / 503**, honor `Retry-After` (seconds) and retry, with a
  cap.
- Flatten people → `Title <EMail>`, lookups → `Title`, multi-value → joined,
  Boolean → Yes/No, URL → `Description (Url)`.
- Download JSON in the schema above.

## Half 2 — MCP server (Python, `mcp` SDK / FastMCP)

- Load JSON exports from a configured folder (`$SPLA_EXPORT_DIR`, else
  `../config/exports`). Tolerate bad/foreign files by tagging them, not raising.
- Redaction pass applied to every returned row, in the server, right before
  return: SSNs (incl. spaced/dotted forms), EINs, grouped credit cards, US
  phone numbers, and 7+-digit id runs, masked to their last 4; dates, dollar
  amounts, and decimals preserved. Emails and personal names are deliberately
  retained — cross-referencing people is the tool's purpose.
- Four read-only tools:
  - `sharepoint_list_status` — exports on disk, with row/column counts.
  - `sharepoint_list_import(list)` — one List's columns + row count.
  - `sharepoint_list_items(list, columns?, filter?, limit?, offset?, redact?)`
    — rows, with column pick, text filter, paging; redacted by default.
  - `sharepoint_list_related(list, topic, top_k?, redact?)` — rank rows by share
    of topic terms matched, best first; redacted by default.
- Run as a **stdio** server, registered in the agent's MCP config.
- Keep the pure data logic dependency-free (separate module) so it is testable
  with plain `python3` (`python` on Windows); only `sharepoint_list_mcp.py` and
  `demo.py` import `mcp`.

## Layout

```
browser/   read-list.js
server/    sharepoint_list_mcp.py  store.py  policy.py  requirements.txt
           test_agent.py  demo.py  graph_export.py
config/    exports/  mcp.example.json
docs/      SPEC.md  WORKFLOW.md
```

- `demo.py` — MCP stdio client: spawns the server and exercises all four tools
  end-to-end against whatever is in the export folder.
- `graph_export.py` — device-code Graph exporter that writes the same canonical
  JSON into `config/exports/` (see the Graph exporter section below).

## Workflow

Run the browser script to export a List → drop the JSON in the server's export
folder → `pip install -r requirements.txt` → register the server with the agent
→ the agent reads and cross-references (e.g. Outlook subject line →
`sharepoint_list_related` → `sharepoint_list_items`).

## Graph exporter (where Graph is allowed)

`server/graph_export.py` ships this path: a stdlib-only device-code sign-in
(no MSAL; uses the Microsoft Graph Command Line Tools client id) that writes
the **same JSON shape** into the export folder, so the server and tools stay
unchanged. Flags `--site` / `--list` / `--all` / `--scan` / `--request-code` /
`--poll`; the token is cached at `~/.cache/spla_token.json` (0600) with silent
refresh. It hardcodes **commercial-cloud** endpoints — GCC High would need
`login.microsoftonline.us` / `graph.microsoft.us` via the `AUTH` / `GRAPH`
constants at the top of the file. The browser reader remains the path for
locked-down tenants. Before trying it at work, confirm: (1) you can register an
app, (2) `Sites.Read.All` is grantable under user consent, (3) Conditional
Access allows the flow.
