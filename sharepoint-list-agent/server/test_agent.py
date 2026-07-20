#!/usr/bin/env python3
"""Runnable check for the data logic — no `mcp` install needed.

    python test_agent.py

Exercises store.py + policy.py against config/exports/sample-list.json, which is
the only "test" this project has (mirrors the repo's `node --check` convention).
The MCP entry point is verified separately with `python -m py_compile`.
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import store
import policy

EXPORTS = HERE.parent / "config" / "exports"
_checks = 0


def ok(label, cond):
    global _checks
    _checks += 1
    if not cond:
        raise AssertionError("FAIL: " + label)
    print("PASS", label)


def main():
    exports = store.load_exports(EXPORTS)
    ok("sample export loads", len(store.valid(exports)) >= 1)

    e = store.find(exports, "Case Tracker")
    ok("find by title", e is not None)
    ok("find by partial title", store.find(exports, "case") is not None)
    ok("row count", len(e["items"]) == 2)

    # column resolution: display name and internal name both map to the item key
    ok("resolve display name", store.resolve_columns(e, ["Status"]) == ["Status"])
    ok("resolve internal name", store.resolve_columns(e, ["ContactPhone"]) == ["Contact Phone"])

    # filter
    f = store.filter_items(e["items"], "overtime")
    ok("filter matches one row", len(f) == 1 and f[0]["Title"] == "MW-2026-0011")
    ok("filter miss returns none", store.filter_items(e["items"], "zzzznope") == [])

    # projection
    p = store.project(e["items"], ["Title", "Status"])
    ok("projection keeps only asked cols", set(p[0].keys()) == {"Title", "Status"})

    # relevance ranking against a subject-line-like topic
    ranked = store.rank_related(e["items"], "Acme minimum wage complaint")
    ok("related ranks a hit first", ranked and ranked[0]["row"]["Employer"] == "Acme Diner LLC")
    ok("related drops zero-score rows", all(h["score"] > 0 for h in ranked))

    # redaction — the security-critical pass
    red = policy.redact_item(e["items"][0])
    blob = str(red)
    ok("ssn masked", "123-45-6789" not in blob)
    ok("phone masked", "202-555-0143" not in blob and "(202) 555-0143" not in blob)
    ok("parenthesized phone fully masked (no stray paren)",
       "([phone]" not in blob and "Call back at [phone]." in red["Notes"])
    ok("case-id masked", "4471203" not in blob)
    ok("non-sensitive text preserved", red["Employer"] == "Acme Diner LLC")
    ok("investigator name preserved", "Jane Doe" in blob)

    print("\nALL PASS — %d exports, %d checks" % (len(store.valid(exports)), _checks))


if __name__ == "__main__":
    main()
