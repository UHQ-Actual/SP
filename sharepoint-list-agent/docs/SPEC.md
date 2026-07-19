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
  return: SSN (`123-45-6789`), phone, and 7–10 digit case/id runs.
- Four read-only tools:
  - `sharepoint_list_status` — exports on disk, with row/column counts.
  - `sharepoint_list_import(list)` — one List's columns + row count.
  - `sharepoint_list_items(list, columns?, filter?, limit?, offset?, redact?)`
    — rows, with column pick, text filter, paging; redacted by default.
  - `sharepoint_list_related(list, topic, top_k?, redact?)` — rank rows by share
    of topic terms matched, best first; redacted by default.
- Run as a **stdio** server, registered in the agent's MCP config.
- Keep the pure data logic dependency-free (separate module) so it is testable
  with plain `python`; only the entry point imports `mcp`.

## Layout

```
browser/   read-list.js
server/    sharepoint_list_mcp.py  store.py  policy.py  requirements.txt  test_agent.py
config/    exports/  mcp.example.json
docs/      SPEC.md  WORKFLOW.md
```

## Workflow

Run the browser script to export a List → drop the JSON in the server's export
folder → `pip install -r requirements.txt` → register the server with the agent
→ the agent reads and cross-references (e.g. Outlook subject line →
`sharepoint_list_related` → `sharepoint_list_items`).

## Upgrade path (only if IT clears Graph)

Add an MSAL device-code or interactive sign-in for `Sites.Read.All` that writes
the **same JSON shape** into the export folder. The server and tools stay
unchanged. Before building it, confirm: (1) you can register an app,
(2) `Sites.Read.All` is grantable under user consent, (3) Conditional Access
allows the flow — and which cloud you're on (commercial vs GCC High), since the
endpoints differ (`graph.microsoft.com` vs `graph.microsoft.us`, and the
matching AAD authority host).
