#!/usr/bin/env python3
"""Graph exporter — the Graph upgrade path described in docs/SPEC.md.

Reads SharePoint Lists via Microsoft Graph with a delegated device-code
sign-in and writes the SAME canonical JSON shape the browser reader produces
into config/exports/, so the MCP server needs no changes. Use it where Graph
is NOT blocked (e.g. a personal dev tenant) — the browser reader remains the
path for locked-down tenants.

AUTH and GRAPH below are the commercial-cloud endpoints. A GCC High tenant
needs login.microsoftonline.us and graph.microsoft.us instead — change the
two constants.

stdlib only — no MSAL, no requests.

  python3 graph_export.py                        # sign in, inventory the root site's lists
  python3 graph_export.py --all                  # export every kept list on the site
  python3 graph_export.py --list "My List"       # one list (searches other sites too)
  python3 graph_export.py --site https://x.sharepoint.com/sites/dev --all
  python3 graph_export.py --scan --all           # sweep every reachable site

Two-step mode (so an agent can show you the code, then poll in background):

  python3 graph_export.py --request-code --device-file /path/dc.json
  python3 graph_export.py --poll /path/dc.json --all
"""
from __future__ import annotations

import argparse
import base64
import json
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import urlencode, urlparse
from urllib.request import Request, urlopen

# "Microsoft Graph Command Line Tools" — well-known first-party public client
# that permits device-code sign-in; consent is still yours to grant.
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
AUTH = "https://login.microsoftonline.com/organizations/oauth2/v2.0"
GRAPH = "https://graph.microsoft.com/v1.0"
SCOPE = "https://graph.microsoft.com/Sites.Read.All offline_access openid profile"

KEEP_RO = {"ID", "Created", "Modified", "Author", "Editor"}
SKIP_NAMES = {
    "ContentType", "Attachments", "DocIcon", "Edit", "LinkTitle",
    "LinkTitleNoMenu", "ItemChildCount", "FolderChildCount", "AppAuthor",
    "AppEditor", "ComplianceAssetId",
}
LIB_TEMPLATES = {
    "documentLibrary", "webPageLibrary", "formLibrary", "pictureLibrary",
    "styleLibrary", "mySiteDocumentLibrary", "xmlForm",
}


def log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


# ---------------------------------------------------------------- HTTP layer
def form_post(url: str, fields: dict) -> dict:
    req = Request(url, data=urlencode(fields).encode(), headers={
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
    })
    try:
        with urlopen(req, timeout=30) as r:
            return json.load(r)
    except HTTPError as e:
        body = e.read().decode("utf-8", "replace")
        try:
            return json.loads(body)  # AAD returns structured errors as 400s
        except ValueError:
            raise RuntimeError("%s -> HTTP %s: %s" % (url, e.code, body[:300]))


def graph_get(token: str, url: str) -> dict:
    for attempt in range(6):
        req = Request(url, headers={
            "Authorization": "Bearer " + token,
            "Accept": "application/json",
        })
        try:
            with urlopen(req, timeout=60) as r:
                return json.load(r)
        except HTTPError as e:
            if e.code in (429, 503) and attempt < 5:
                wait = int(e.headers.get("Retry-After") or "5")
                log("[graph] throttled %s; waiting %ss" % (e.code, wait))
                time.sleep(wait)
                continue
            body = e.read().decode("utf-8", "replace")
            raise RuntimeError("GET %s -> HTTP %s: %s" % (url, e.code, body[:400]))
    raise RuntimeError("unreachable")


def graph_all(token: str, url: str) -> list:
    rows: list = []
    while url:
        j = graph_get(token, url)
        rows.extend(j.get("value", []))
        url = j.get("@odata.nextLink")
    return rows


# ------------------------------------------------------------- device flow
def request_code(device_file: Path) -> dict:
    resp = form_post(AUTH + "/devicecode", {"client_id": CLIENT_ID, "scope": SCOPE})
    if "device_code" not in resp:
        raise RuntimeError("device code request failed: %s" % json.dumps(resp)[:300])
    resp["_requested_at"] = time.time()
    device_file.parent.mkdir(parents=True, exist_ok=True)
    device_file.write_text(json.dumps(resp), encoding="utf-8")
    try:
        device_file.chmod(0o600)
    except OSError:
        pass
    public = {
        "verification_uri": resp.get("verification_uri", "https://microsoft.com/devicelogin"),
        "user_code": resp["user_code"],
        "expires_in_min": int(resp.get("expires_in", 900)) // 60,
    }
    print(json.dumps(public, indent=2))
    return resp


def poll_token(device_file: Path) -> dict:
    dc = json.loads(device_file.read_text(encoding="utf-8"))
    interval = int(dc.get("interval", 5))
    deadline = dc.get("_requested_at", time.time()) + int(dc.get("expires_in", 900))
    log("[auth] waiting for sign-in at %s (code %s)..." %
        (dc.get("verification_uri"), dc.get("user_code")))
    while time.time() < deadline:
        resp = form_post(AUTH + "/token", {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": CLIENT_ID,
            "device_code": dc["device_code"],
        })
        err = resp.get("error")
        if not err:
            log("[auth] signed in.")
            return resp
        if err == "authorization_pending":
            time.sleep(interval)
            continue
        if err == "slow_down":
            interval += 5
            time.sleep(interval)
            continue
        raise SystemExit("[auth] %s: %s" % (err, resp.get("error_description", "")[:300]))
    raise SystemExit("[auth] device code expired before sign-in; re-run --request-code")


def save_token(path: Path, tok: dict) -> None:
    tok = dict(tok)
    tok["_obtained_at"] = time.time()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(tok), encoding="utf-8")
    try:
        path.chmod(0o600)
    except OSError:
        pass


def load_cached_token(path: Path) -> dict | None:
    """Reuse a cached access token, refreshing it when expired. Returns a token
    response dict, or None when a fresh device-code sign-in is required."""
    if not path.is_file():
        return None
    try:
        t = json.loads(path.read_text(encoding="utf-8"))
    except ValueError:
        return None
    age = time.time() - t.get("_obtained_at", 0)
    if t.get("access_token") and age < int(t.get("expires_in", 3600)) - 300:
        log("[auth] using cached access token")
        return t
    if t.get("refresh_token"):
        log("[auth] refreshing cached token...")
        resp = form_post(AUTH + "/token", {
            "grant_type": "refresh_token",
            "client_id": CLIENT_ID,
            "refresh_token": t["refresh_token"],
            "scope": SCOPE,
        })
        if resp.get("access_token"):
            resp.setdefault("refresh_token", t["refresh_token"])
            resp.setdefault("id_token", t.get("id_token"))
            log("[auth] refreshed.")
            return resp
        log("[auth] refresh failed: %s" % resp.get("error", "?"))
    return None


def upn_from_token(tok: dict) -> str | None:
    """preferred_username out of the id_token payload — no extra Graph call."""
    idt = tok.get("id_token")
    if not idt:
        return None
    try:
        payload = idt.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload))
        return claims.get("preferred_username") or claims.get("upn")
    except Exception:
        return None


# ------------------------------------------------------------- site + lists
def resolve_site(token: str, site: str | None) -> dict:
    if not site:
        return graph_get(token, GRAPH + "/sites/root")
    if site.startswith("http"):
        p = urlparse(site)
        host, path = p.netloc, p.path.strip("/")
    else:
        host, _, path = site.partition("/")
        path = path.strip("/")
    # A pasted list-view or page URL must resolve to its site: truncate at the
    # first in-site marker, then drop a trailing *.aspx segment if one remains.
    low = "/" + path.lower() + "/"
    cut = len(low)
    for marker in ("/_layouts/", "/lists/", "/sitepages/", "/forms/"):
        i = low.find(marker)
        if 0 <= i < cut:
            cut = i
    path = path[:max(cut - 1, 0)].strip("/")
    segs = [s for s in path.split("/") if s]
    if segs and segs[-1].lower().endswith(".aspx"):
        segs.pop()
    path = "/".join(segs)
    url = GRAPH + "/sites/" + host + ((":/" + path) if path else "")
    return graph_get(token, url)


def site_lists(token: str, site_id: str) -> list:
    return graph_all(token, GRAPH + "/sites/%s/lists?$select=id,displayName,webUrl,list,system&$expand=columns" % site_id)


def wanted_lists(lists: list, include_libs: bool = False):
    keep, skipped = [], []
    for l in lists:
        fac = l.get("list") or {}
        if "system" in l or fac.get("hidden"):
            skipped.append((l.get("displayName", "?"), "system/hidden"))
        elif not include_libs and fac.get("template") in LIB_TEMPLATES:
            skipped.append((l.get("displayName", "?"), "library"))
        else:
            keep.append(l)
    return keep, skipped


def scan_sites(token: str, cap: int = 40) -> list:
    sites = graph_all(token, GRAPH + "/sites?search=*&$top=25")
    if len(sites) > cap:
        log("[scan] %d sites found; scanning first %d" % (len(sites), cap))
        sites = sites[:cap]
    return sites


# ------------------------------------------------------------- flattening
def column_type(col: dict) -> str:
    if col.get("personOrGroup") is not None:
        return "UserMulti" if col["personOrGroup"].get("allowMultipleSelection") else "User"
    if col.get("lookup") is not None:
        return "LookupMulti" if col["lookup"].get("allowMultipleValues") else "Lookup"
    if col.get("choice") is not None:
        # v1.0 choiceColumn has no allowMultipleSelections; multi-select is
        # signaled by displayAs == "checkBoxes".
        return "MultiChoice" if col["choice"].get("displayAs") == "checkBoxes" else "Choice"
    if col.get("dateTime") is not None:
        return "DateTime"
    if col.get("currency") is not None:
        return "Currency"
    if col.get("number") is not None:
        return "Number"
    if col.get("boolean") is not None:
        return "Boolean"
    if col.get("hyperlinkOrPicture") is not None:
        return "URL"
    if col.get("calculated") is not None:
        return "Calculated"
    if col.get("text") is not None:
        return "Note" if col["text"].get("allowMultipleLines") else "Text"
    return "Unknown"


def build_columns(cols: list) -> list:
    out = []
    for c in cols:
        n = c.get("name", "")
        if c.get("hidden") or n.startswith("_") or n in SKIP_NAMES:
            continue
        if c.get("readOnly") and n not in KEEP_RO:
            continue
        out.append({"name": c.get("displayName") or n, "internal": n, "type": column_type(c)})
    if not any(c["internal"] in ("ID", "Id") for c in out):
        out.insert(0, {"name": "ID", "internal": "Id", "type": "Counter"})
    return out


def flatten(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "Yes" if v else "No"
    if isinstance(v, list):
        return "; ".join(str(flatten(x)) for x in v if x is not None)
    if isinstance(v, dict):
        if "LookupValue" in v:
            email = v.get("Email") or ""
            base = str(v.get("LookupValue") or "")
            return "%s <%s>" % (base, email) if email and email != base else base
        if "Url" in v:
            d = v.get("Description")
            return "%s (%s)" % (d, v["Url"]) if d and d != v["Url"] else str(v["Url"])
        return json.dumps(v, ensure_ascii=False)
    return v


def build_rows(items: list, columns: list) -> list:
    rows = []
    for it in items:
        f = it.get("fields", {}) or {}
        try:
            rid = int(it.get("id"))
        except (TypeError, ValueError):
            rid = it.get("id")
        row = {"Id": rid}
        for c in columns:
            n = c["internal"]
            if n in ("Id", "ID"):
                continue
            if n == "Author":
                u = (it.get("createdBy") or {}).get("user", {})
                v = _person(u) or _lookup_fallback(f, n)
            elif n == "Editor":
                u = (it.get("lastModifiedBy") or {}).get("user", {})
                v = _person(u) or _lookup_fallback(f, n)
            elif n in f:
                v = flatten(f[n])
            else:
                v = _lookup_fallback(f, n)
            row[c["name"]] = v
        rows.append(row)
    return rows


def _person(u: dict) -> str:
    """identitySet user -> 'Name <email>', matching flatten()'s LookupValue shape."""
    name = u.get("displayName") or ""
    email = u.get("email") or ""
    return "%s <%s>" % (name, email) if name and email else name


def _lookup_fallback(fields: dict, name: str):
    lv = fields.get(name + "LookupId")
    if lv is None:
        return ""
    if isinstance(lv, list):
        return "; ".join("id:%s" % x for x in lv)
    return "id:%s" % lv


# ------------------------------------------------------------- export
_written: set = set()  # paths written this run, so scan collisions uniquify

def export_list(token: str, site: dict, l: dict, outdir: Path, tag_site: bool) -> dict:
    title = l.get("displayName", "?")
    log("[export] %s ..." % title)
    cols = l.get("columns") or graph_all(
        token, GRAPH + "/sites/%s/lists/%s/columns" % (site["id"], l["id"]))
    columns = build_columns(cols)
    # fields($select=...) makes Graph return lookup/person VALUES (LookupValue/
    # Email dicts) instead of bare <name>LookupId — see the fieldValueSet doc.
    internals = [c["internal"] for c in columns]
    lookupish = [c for c in columns if c["type"] in ("User", "UserMulti", "Lookup", "LookupMulti")]
    if len(lookupish) > 12:
        log("[export] %s: %d lookup/person columns — Graph resolves at most 12 per query; the rest fall back to id:N" % (title, len(lookupish)))
    items_url = GRAPH + "/sites/%s/lists/%s/items?$expand=fields($select=%s)&$top=200" % (
        site["id"], l["id"], ",".join(internals))
    try:
        items = graph_all(token, items_url)
    except RuntimeError as e:
        if "HTTP 400" not in str(e):
            raise
        log("[export] %s: fields($select=...) rejected; retrying with bare $expand=fields (lookups fall back to id:N)" % title)
        items = graph_all(
            token, GRAPH + "/sites/%s/lists/%s/items?$expand=fields&$top=200" % (site["id"], l["id"]))
    rows = build_rows(items, columns)

    export = {
        "web": site.get("webUrl"),
        "list": title,
        "listId": l.get("id"),
        "exportedAt": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source": "graph",
        "columns": columns,
        "items": rows,
    }
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", title)
    day = datetime.now(timezone.utc).strftime("%Y%m%d")
    parts = ["sp"]
    if tag_site:
        # site path is unique across the tenant; site.name is not
        slug = urlparse(site.get("webUrl", "")).path.strip("/").replace("/", "_") or site.get("name") or "site"
        parts.append(re.sub(r"[^A-Za-z0-9._-]+", "_", slug))
    parts += [safe, day]
    base = "_".join(parts)
    path = outdir / (base + ".json")
    n = 2
    while path in _written:
        path = outdir / ("%s_%d.json" % (base, n))
        n += 1
    _written.add(path)
    outdir.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(export, indent=2, ensure_ascii=False), encoding="utf-8")
    log("[export] %s -> %s (%d rows, %d cols)" % (title, path.name, len(rows), len(columns)))
    return {"list": title, "site": site.get("webUrl"), "file": path.name,
            "rows": len(rows), "columns": len(columns)}


# ------------------------------------------------------------- main
def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--site", help="site URL or host[/path]; default: tenant root site")
    ap.add_argument("--list", dest="list_title", help="export just this list (by display name)")
    ap.add_argument("--all", action="store_true", help="export every non-hidden, non-library list on the site")
    ap.add_argument("--scan", action="store_true", help="also sweep every reachable site collection")
    ap.add_argument("--include-libraries", action="store_true", help="treat document libraries as exportable")
    ap.add_argument("--out", type=Path, default=Path(__file__).resolve().parent.parent / "config" / "exports")
    ap.add_argument("--request-code", action="store_true", help="step 1: print sign-in code and save device state")
    ap.add_argument("--poll", type=Path, metavar="DEVICE_FILE", help="step 2: wait for sign-in, then export")
    ap.add_argument("--device-file", type=Path, default=Path.home() / ".cache" / "spla_devicecode.json")
    ap.add_argument("--token-cache", type=Path, default=Path.home() / ".cache" / "spla_token.json",
                    help="cached token (0600) so later runs skip the sign-in entirely")
    args = ap.parse_args()

    if args.request_code:
        request_code(args.device_file)
        return

    tok = load_cached_token(args.token_cache)
    if not tok:
        if args.poll:
            tok = poll_token(args.poll)
        else:  # interactive one-shot
            dc = request_code(args.device_file)
            log("\n  Go to %s and enter code %s\n" % (dc.get("verification_uri"), dc["user_code"]))
            tok = poll_token(args.device_file)
    save_token(args.token_cache, tok)

    token = tok["access_token"]
    summary = {"sites": [], "exports": [], "skipped": []}
    inventory = not args.list_title and not args.all  # neither flag: list, don't export

    def handle_site(site: dict, tag_site: bool) -> bool:
        lists = site_lists(token, site["id"])
        keep, skipped = wanted_lists(lists, args.include_libraries)
        summary["sites"].append(site.get("webUrl"))
        for name, why in skipped:
            summary["skipped"].append({"list": name, "site": site.get("webUrl"), "why": why})
        if inventory:
            log("[inventory] %s" % site.get("webUrl"))
            for l in keep:
                log("  %s (%s)" % (l.get("displayName", "?"), (l.get("list") or {}).get("template", "?")))
            return bool(keep)
        hit = False
        for l in keep:
            if args.list_title and l.get("displayName", "").lower() != args.list_title.lower():
                continue
            summary["exports"].append(export_list(token, site, l, args.out, tag_site))
            hit = True
        return hit

    root = resolve_site(token, args.site)
    found = handle_site(root, tag_site=False)

    need_scan = args.scan or (not found and not args.site)
    if need_scan:
        log("[scan] sweeping site collections...")
        scanned = {root.get("id")}
        for s in scan_sites(token):
            if s.get("id") in scanned:
                continue
            scanned.add(s.get("id"))
            try:
                if handle_site(s, tag_site=True) and args.list_title:
                    break
            except RuntimeError as e:
                summary["skipped"].append({"site": s.get("webUrl"), "why": str(e)[:160]})

        # Microsoft Lists "My lists" live in the user's personal site, which
        # sites?search=* often omits — probe it explicitly via the UPN.
        upn = upn_from_token(tok)
        if upn and "@" in upn:
            # SharePoint keeps hyphens in the personal-site slug; only @ and . map to _
            slug = upn.replace("@", "_").replace(".", "_").lower()
            myhost = (urlparse(root.get("webUrl", "")).netloc
                      .replace(".sharepoint.com", "-my.sharepoint.com")
                      .replace(".sharepoint.us", "-my.sharepoint.us"))
            try:
                ps = graph_get(token, GRAPH + "/sites/%s:/personal/%s" % (myhost, slug))
                if ps.get("id") not in scanned:
                    log("[scan] personal site (My lists): %s" % ps.get("webUrl"))
                    handle_site(ps, tag_site=True)
            except RuntimeError as e:
                summary["skipped"].append({"site": "personal:" + slug, "why": str(e)[:160]})

    print("SPLA_EXPORT_SUMMARY " + json.dumps(summary, ensure_ascii=False))
    if inventory:
        log('[done] inventory only — pass --all to export every list above, or --list "Name" for one')
        sys.exit(0)
    if not summary["exports"]:
        log("[done] nothing exported — check the list name/site, or pass --scan / --include-libraries")
        sys.exit(2)
    log("[done] %d export(s) written to %s" % (len(summary["exports"]), args.out))


if __name__ == "__main__":
    main()
