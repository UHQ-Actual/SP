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
    an `_error` key rather than raising, so one bad file never blinds the rest."""
    out: List[Dict[str, Any]] = []
    p = Path(export_dir)
    if not p.is_dir():
        return out
    for fp in sorted(p.glob("*.json")):
        try:
            data = json.loads(fp.read_text(encoding="utf-8"))
        except Exception as e:  # noqa: BLE001 - report, don't crash the scan
            out.append({"_file": fp.name, "_error": "parse error: %s" % e})
            continue
        if not isinstance(data, dict) or "items" not in data:
            out.append({"_file": fp.name, "_error": "not a list export (missing 'items')"})
            continue
        data["_file"] = fp.name
        out.append(data)
    return out


def valid(exports: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [e for e in exports if "_error" not in e]


def find(exports: List[Dict[str, Any]], name: str) -> Optional[Dict[str, Any]]:
    """Match by exact list title, then exact filename, then partial title."""
    good = valid(exports)
    nl = (name or "").strip().lower()
    for e in good:
        if str(e.get("list", "")).lower() == nl:
            return e
    for e in good:
        f = str(e["_file"]).lower()
        if f == nl or f == nl + ".json":
            return e
    for e in good:
        if nl and nl in str(e.get("list", "")).lower():
            return e
    return None


def available(exports: List[Dict[str, Any]]) -> List[str]:
    return [str(e.get("list", e.get("_file"))) for e in valid(exports)]


def resolve_columns(export: Dict[str, Any], requested) -> List[str]:
    """Map requested column tokens (display OR internal name) to the item keys."""
    keymap: Dict[str, str] = {}
    for c in export.get("columns", []):
        nm = c.get("name")
        keymap[str(c.get("name", "")).lower()] = nm
        keymap[str(c.get("internal", "")).lower()] = nm
    out: List[str] = []
    for r in requested or []:
        k = keymap.get(str(r).lower())
        if k and k not in out:
            out.append(k)
    return out


def filter_items(items: List[dict], needle: str) -> List[dict]:
    if not needle:
        return list(items)
    nl = needle.lower()
    return [r for r in items if nl in " ".join(str(v) for v in r.values()).lower()]


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
        rset = set(tokens(" ".join(str(v) for v in r.values())))
        matched = sorted(terms & rset)
        if not matched:
            continue
        scored.append({"score": round(len(matched) / len(terms), 3), "matched": matched, "row": r})
    scored.sort(key=lambda x: (-x["score"], -len(x["matched"])))
    return scored[: max(1, top_k)]
