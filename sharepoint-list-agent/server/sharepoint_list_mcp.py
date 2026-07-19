#!/usr/bin/env python3
"""SharePoint List Agent — read-only MCP server.

Makes NO network calls of its own. It reads the JSON exports the browser reader
downloaded (see ../browser/read-list.js), applies a redaction pass, and exposes
four read-only tools to the coding agent over stdio:

  sharepoint_list_status   which exports are on disk
  sharepoint_list_import   one List's columns + row count
  sharepoint_list_items    rows, with column pick / text filter / paging
  sharepoint_list_related  rank rows against a topic (e.g. an email subject)

Export folder resolution: $SPLA_EXPORT_DIR, else ../config/exports next to this
file. Every row that leaves the process passes through policy.redact_item first.

Run:      python sharepoint_list_mcp.py
Install:  pip install -r requirements.txt   (just the `mcp` SDK)
"""
from __future__ import annotations
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import FastMCP

import store
import policy

EXPORT_DIR = Path(
    os.environ.get("SPLA_EXPORT_DIR")
    or (Path(__file__).resolve().parent.parent / "config" / "exports")
).resolve()

mcp = FastMCP("sharepoint-list")


def _load():
    return store.load_exports(EXPORT_DIR)


@mcp.tool()
def sharepoint_list_status() -> Dict[str, Any]:
    """List which SharePoint List exports are available locally, with row and
    column counts. Call this first to see what the agent can read. No arguments."""
    exports = _load()
    rows: List[Dict[str, Any]] = []
    for e in exports:
        if "_error" in e:
            rows.append({"file": e["_file"], "error": e["_error"]})
            continue
        rows.append({
            "list": e.get("list"),
            "web": e.get("web"),
            "file": e["_file"],
            "columns": len(e.get("columns", [])),
            "items": len(e.get("items", [])),
            "exportedAt": e.get("exportedAt"),
        })
    return {
        "export_dir": str(EXPORT_DIR),
        "count": len(store.valid(exports)),
        "exports": rows,
        "redaction": policy.POLICY_SUMMARY,
    }


@mcp.tool()
def sharepoint_list_import(list: str) -> Dict[str, Any]:
    """Summarize one List's schema: its columns (display name, internal name,
    type) and its row count.

    Args:
        list: the List title or export filename (partial title also matches).
    """
    exports = _load()
    e = store.find(exports, list)
    if not e:
        return {"error": "No export found for %r" % list, "available": store.available(exports)}
    return {
        "list": e.get("list"),
        "web": e.get("web"),
        "file": e["_file"],
        "item_count": len(e.get("items", [])),
        "columns": e.get("columns", []),
    }


@mcp.tool()
def sharepoint_list_items(
    list: str,
    columns: Optional[List[str]] = None,
    filter: str = "",
    limit: int = 50,
    offset: int = 0,
    redact: bool = True,
) -> Dict[str, Any]:
    """Return rows from a List export. Redacted by default.

    Args:
        list: List title or export filename.
        columns: optional subset of columns to return (display or internal names).
        filter: case-insensitive substring matched across each row's values.
        limit: max rows to return (0 = no limit). Default 50.
        offset: rows to skip, for paging. Default 0.
        redact: mask SSNs / phones / case-ids before returning. Default True.
    """
    exports = _load()
    e = store.find(exports, list)
    if not e:
        return {"error": "No export found for %r" % list, "available": store.available(exports)}

    matched = store.filter_items(e.get("items", []), filter)
    total = len(matched)
    limit = max(0, int(limit))
    offset = max(0, int(offset))
    page = matched[offset: offset + limit] if limit else matched[offset:]

    keys = store.resolve_columns(e, columns) if columns else []
    if keys:
        page = store.project(page, ["Id"] + keys)
    if redact:
        page = [policy.redact_item(r) for r in page]

    return {
        "list": e.get("list"),
        "total_matched": total,
        "offset": offset,
        "returned": len(page),
        "redacted": bool(redact),
        "items": page,
    }


@mcp.tool()
def sharepoint_list_related(
    list: str,
    topic: str,
    top_k: int = 10,
    redact: bool = True,
) -> Dict[str, Any]:
    """Rank a List's rows by textual relevance to a topic — e.g. paste an Outlook
    subject line to find the List rows about the same matter.

    Args:
        list: List title or export filename.
        topic: free text (subject line, keywords) to rank rows against.
        top_k: how many top rows to return. Default 10.
        redact: mask SSNs / phones / case-ids before returning. Default True.
    """
    exports = _load()
    e = store.find(exports, list)
    if not e:
        return {"error": "No export found for %r" % list, "available": store.available(exports)}

    ranked = store.rank_related(e.get("items", []), topic, top_k)
    if not ranked and not store.tokens(topic):
        return {"error": "topic produced no searchable terms", "list": e.get("list")}

    results = []
    for hit in ranked:
        row = policy.redact_item(hit["row"]) if redact else hit["row"]
        results.append({"score": hit["score"], "matched": hit["matched"], "row": row})

    return {
        "list": e.get("list"),
        "topic": topic,
        "returned": len(results),
        "redacted": bool(redact),
        "results": results,
    }


if __name__ == "__main__":
    mcp.run()
