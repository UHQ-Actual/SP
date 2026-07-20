#!/usr/bin/env python3
"""Live demo — spawn the MCP server over stdio and call all four tools.

    python demo.py                       # auto-picks the first export on disk
    python demo.py "My List"             # target a specific list
    python demo.py "My List" "topic ..." # and a specific relevance topic

Talks the real MCP protocol to sharepoint_list_mcp.py against whatever is in
config/exports/ (ships working against the synthetic sample-list.json, and works
unchanged on a real export you drop in). Proves the agent-facing half end-to-end
without needing SharePoint. Requires the `mcp` SDK: pip install -r requirements.txt
"""
import asyncio
import json
import sys
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

HERE = Path(__file__).resolve().parent


def _data(result):
    """Unwrap a CallToolResult into a plain Python object."""
    d = getattr(result, "structuredContent", None)
    if isinstance(d, dict):
        return d.get("result", d)  # FastMCP wraps dict returns under "result"
    if d is not None:
        return d
    for c in result.content:
        t = getattr(c, "text", None)
        if t:
            try:
                return json.loads(t)
            except Exception:
                return t
    return None


def show(title, result):
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)
    print(json.dumps(_data(result), indent=2, ensure_ascii=False))


def _topic_from(rows):
    """Build a plausible relevance topic from the first row's non-redacted text."""
    for r in rows or []:
        vals = [str(v) for v in r.values()
                if isinstance(v, str) and v and not v.startswith("[") and "*" not in v]
        toks = " ".join(vals).split()
        if len(toks) >= 3:
            return " ".join(toks[:6])
    return ""


async def main():
    list_override = sys.argv[1] if len(sys.argv) > 1 else None
    topic_override = sys.argv[2] if len(sys.argv) > 2 else None

    params = StdioServerParameters(command=sys.executable, args=[str(HERE / "sharepoint_list_mcp.py")])
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = await session.list_tools()
            print("Registered tools:", ", ".join(t.name for t in tools.tools))

            status = await session.call_tool("sharepoint_list_status", {})
            show("sharepoint_list_status()", status)

            exports = (_data(status) or {}).get("exports", [])
            names = [e["list"] for e in exports if e.get("list")]
            if not names:
                print("\nNo exports on disk. Drop a JSON in config/exports/ and re-run.")
                return

            target = list_override or names[0]
            print("\n>>> Demoing list: %r" % target)

            show("sharepoint_list_import(list=%r)" % target,
                 await session.call_tool("sharepoint_list_import", {"list": target}))

            items = await session.call_tool("sharepoint_list_items", {"list": target, "limit": 5})
            show("sharepoint_list_items(list=%r, limit=5)  [redacted]" % target, items)

            topic = topic_override or _topic_from((_data(items) or {}).get("items"))
            if topic:
                show("sharepoint_list_related(list=%r, topic=%r)" % (target, topic),
                     await session.call_tool("sharepoint_list_related", {"list": target, "topic": topic}))
            else:
                print("\n(skipped related — no non-redacted text to build a topic from)")


if __name__ == "__main__":
    asyncio.run(main())
