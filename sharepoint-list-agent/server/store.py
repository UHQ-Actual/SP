"""Pure data logic for the SharePoint List Agent.

No MCP, no network — just reads the downloaded JSON exports and slices them.
Kept dependency-free so it can be unit-tested with plain `python` (see
test_agent.py). The MCP entry point (sharepoint_list_mcp.py) is a thin wrapper
over these functions plus the redaction pass.

Export shape produced by the browser reader:
  { "web": str, "list": str, "listId": str|None, "exportedAt": str,
    "columns": [ {"name","internal","type"} ],
    "items":   [ {<column name>: <value>, ...} ] }
"""
from __future__ import annotations
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

_TOKEN_RE = re.compile(r"[a-z0-9]+")
# Common email/subject-line noise that should not drive relevance scoring.
_STOP = {
    "the", "a", "an", "of", "to", "and", "for", "in", "on", "re", "fw", "fwd",
    "with", "is", "at", "by", "or", "your", "our", "this", "that", "from",
}


def tokens(s: str) -> List[str]:
    return [t for t in _TOKEN_RE.findall(str(s).lower()) if t not in _STOP and len(t) > 1]


def load_exports(export_dir) -> List[Dict[str, Any]]:
    """Read every *.json in export_dir. Bad/foreign files come back tagged with
    an `_error` key rather than raising, so one bad file never blinds the rest.
    utf-8-sig tolerates the BOM Windows editors like to prepend; every entry
    (including errored ones) carries `_mtime` for newest-export tie-breaking."""
    out: List[Dict[str, Any]] = []
    p = Path(export_dir)
    if not p.is_dir():
        return out
    for fp in sorted(p.glob("*.json")):
        mtime = fp.stat().st_mtime
        try:
            data = json.loads(fp.read_text(encoding="utf-8-sig"))
        except Exception as e:  # noqa: BLE001 - report, don't crash the scan
            out.append({"_file": fp.name, "_error": "parse error: %s" % e, "_mtime": mtime})
            continue
        if not isinstance(data, dict) or "items" not in data:
            out.append({"_file": fp.name, "_error": "not a list export (missing 'items')", "_mtime": mtime})
            continue
        data["_file"] = fp.name
        data["_mtime"] = mtime
        out.append(data)
    return out


def valid(exports: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [e for e in exports if "_error" not in e]


def newest_key(e: Dict[str, Any]):
    """Sort key for picking the newest of several exports of the same list:
    exportedAt first, file mtime as the fallback."""
    return (e.get("exportedAt") or "", e.get("_mtime") or 0)


def find(exports: List[Dict[str, Any]], name: str) -> Optional[Dict[str, Any]]:
    """Match by exact list title, then exact filename, then partial title.
    When a tier matches several exports (dated re-exports of the same list),
    the newest wins — by exportedAt, then file mtime."""
    good = valid(exports)
    nl = (name or "").strip().lower()
    hits = [e for e in good if str(e.get("list", "")).lower() == nl]
    if hits:
        return max(hits, key=newest_key)
    for e in good:
        f = str(e["_file"]).lower()
        if f == nl or f == nl + ".json":
            return e
    hits = [e for e in good if nl and nl in str(e.get("list", "")).lower()]
    if hits:
        return max(hits, key=newest_key)
    return None


def available(exports: List[Dict[str, Any]]) -> List[str]:
    return [str(e.get("list", e.get("_file"))) for e in valid(exports)]


def resolve_columns(export: Dict[str, Any], requested):
    """Map requested column tokens (display OR internal name) to the actual
    item keys. The counter column is special-cased: 'id'/'ID'/'Id' (any case)
    resolve to 'Id', the key the exporters actually write, not the 'ID'
    display name. Returns (keys, unresolved) — unresolved keeps the original
    spelling of requested names that matched nothing."""
    keymap: Dict[str, str] = {}
    for c in export.get("columns", []):
        nm = c.get("name")
        keymap[str(c.get("name", "")).lower()] = nm
        keymap[str(c.get("internal", "")).lower()] = nm
    keymap["id"] = "Id"  # counter column's item key is 'Id', whatever the display name says
    out: List[str] = []
    unresolved: List[str] = []
    for r in requested or []:
        k = keymap.get(str(r).lower())
        if k:
            if k not in out:
                out.append(k)
        else:
            unresolved.append(r)
    return out, unresolved


def filter_items(items: List[dict], needle: str) -> List[dict]:
    if not needle:
        return list(items)
    nl = needle.lower()
    # Per-field match: a needle can't span two columns, and None never matches.
    return [r for r in items if any(nl in str(v).lower() for v in r.values() if v is not None)]


def project(items: List[dict], keys: List[str]) -> List[dict]:
    if not keys:
        return list(items)
    return [{k: r.get(k) for k in keys if k in r} for r in items]


def rank_related(items: List[dict], topic: str, top_k: int = 10) -> List[dict]:
    """Rank rows by share of topic terms they contain. Returns
    [{score, matched, row}], best first, zero-score rows dropped."""
    terms = set(tokens(topic))
    if not terms:
        return []
    scored: List[dict] = []
    for r in items:
        rset = set(tokens(" ".join(str(v) for v in r.values() if v is not None)))
        matched = sorted(terms & rset)
        if not matched:
            continue
        scored.append({"score": round(len(matched) / len(terms), 3), "matched": matched, "row": r})
    scored.sort(key=lambda x: (-x["score"], -len(x["matched"])))
    return scored[: max(0, top_k)]
