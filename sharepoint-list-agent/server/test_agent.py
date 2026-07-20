#!/usr/bin/env python3
"""Runnable check for the data logic — no `mcp` install needed.

    python3 test_agent.py

Exercises store.py + policy.py against config/exports/sample-list.json plus
synthetic fixtures, and locks in the behavior mandated by the 32-finding
review: redaction breadth (SSN variants, EIN, cards, letter-glued ids),
false-positive guards (dates, dollars, decimals), numeric-value redaction,
newest-export-wins, per-field filtering, and the redact-before-filter
composition that closes the search-oracle leak. The MCP entry point is
verified separately with `python3 -m py_compile` and demo.py.
"""
import json
import sys
import tempfile
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


def r(s):
    return policy.redact_text(s)


def main():
    exports = store.load_exports(EXPORTS)
    ok("sample export loads", len(store.valid(exports)) >= 1)

    e = store.find(exports, "Case Tracker")
    ok("find by title", e is not None)
    ok("find by partial title", store.find(exports, "case") is not None)
    ok("find by filename", store.find(exports, "sample-list.json") is not None)
    ok("row count", len(e["items"]) == 2)

    # ---- newest-export-wins (finding: stale dated files shadow fresh ones)
    with tempfile.TemporaryDirectory() as d:
        base = {"list": "Dup", "columns": [], "items": [{"Id": 1, "Note": "OLD"}]}
        old = dict(base, exportedAt="2026-07-18T00:00:00Z")
        new = dict(base, exportedAt="2026-07-19T00:00:00Z",
                   items=[{"Id": 1, "Note": "NEW"}])
        Path(d, "sp_dup_20260718.json").write_text(json.dumps(old))
        Path(d, "sp_dup_20260719.json").write_text(json.dumps(new))
        got = store.find(store.load_exports(d), "Dup")
        ok("find prefers newest exportedAt", got["items"][0]["Note"] == "NEW")

        # BOM tolerance (finding: utf-8-sig)
        bom = dict(base, list="Bom")
        Path(d, "bom.json").write_bytes(b"\xef\xbb\xbf" + json.dumps(bom).encode())
        ok("BOM export still loads", store.find(store.load_exports(d), "Bom") is not None)

    # ---- column resolution: tuple return, ID mapping, unresolved reporting
    keys, unresolved = store.resolve_columns(e, ["Status"])
    ok("resolve display name", keys == ["Status"] and unresolved == [])
    keys, unresolved = store.resolve_columns(e, ["ContactPhone"])
    ok("resolve internal name", keys == ["Contact Phone"])
    keys, unresolved = store.resolve_columns(e, ["ID"])
    ok("ID maps to real item key 'Id'", keys == ["Id"])
    keys, unresolved = store.resolve_columns(e, ["Status", "Bogus"])
    ok("unresolved names reported", keys == ["Status"] and unresolved == ["Bogus"])
    keys, unresolved = store.resolve_columns(e, ["Bogus"])
    ok("nothing resolved -> empty keys, not full rows", keys == [] and unresolved == ["Bogus"])

    # ---- filtering: per-field, None-safe
    f = store.filter_items(e["items"], "overtime")
    ok("filter matches one row", len(f) == 1 and f[0]["Title"] == "MW-2026-0011")
    ok("filter miss returns none", store.filter_items(e["items"], "zzzznope") == [])
    ok("needle cannot span two fields", store.filter_items(e["items"], "open jane") == [])
    ok("None values are not matchable", store.filter_items([{"A": None}], "none") == [])

    # ---- projection
    p = store.project(e["items"], ["Title", "Status"])
    ok("projection keeps only asked cols", set(p[0].keys()) == {"Title", "Status"})

    # ---- relevance ranking
    ranked = store.rank_related(e["items"], "Acme minimum wage complaint")
    ok("related ranks a hit first", ranked and ranked[0]["row"]["Employer"] == "Acme Diner LLC")
    ok("related drops zero-score rows", all(h["score"] > 0 for h in ranked))
    ok("top_k=0 returns empty", store.rank_related(e["items"], "acme wage", top_k=0) == [])

    # ---- redaction: coverage the review demanded
    ok("ssn dashes", r("123-45-6789") == "***-**-****")
    ok("ssn spaces", r("123 45 6789") == "***-**-****")
    ok("ssn dots", r("123.45.6789") == "***-**-****")
    ok("ein masked", r("EIN 12-3456789") == "EIN **-*******")
    ok("card grouped", r("4111 1111 1111 1111").endswith("1111") and "4111" not in r("4111 1111 1111 1111"))
    ok("card ungrouped", r("4111111111111111") == "************1111")
    ok("phone kept-shape", r("202-555-0143") == "[phone]")
    ok("parenthesized phone fully masked (no stray paren)",
       r("Call back at (202) 555-0143.") == "Call back at [phone].")
    ok("7-digit case id keeps last 4", r("Ref 5583002") == "Ref ***3002")
    ok("letter-glued id masked", r("Case#4471203x") == "Case#***1203x")
    ok("11+ digit run masked", r("acct 300123456789") == "acct ********6789")

    # ---- false-positive guards
    ok("YYYYMMDD date preserved", r("20260719") == "20260719")
    ok("dollar amount preserved", r("$1234567") == "$1234567")
    ok("decimal component preserved", r("release 1.2.3456789") == "release 1.2.3456789")

    # ---- numeric-typed values (the critical bypass)
    red = policy.redact_item({"SSN": 123456789, "Case": 4471203, "Rate": 17.5, "Open": True})
    ok("numeric SSN masked", red["SSN"] == "*****6789")
    ok("numeric case id masked", red["Case"] == "***1203")
    ok("benign number keeps type", red["Rate"] == 17.5)
    ok("bool untouched", red["Open"] is True)

    # ---- deliberate retentions
    ok("emails/names preserved (by design)",
       r("Jane Doe <jane.doe@dol.gov>") == "Jane Doe <jane.doe@dol.gov>")

    # ---- sample rows end-to-end
    row = policy.redact_item(e["items"][0])
    blob = str(row)
    ok("sample ssn masked", "123-45-6789" not in blob)
    ok("sample phones masked", "202-555-0143" not in blob and "(202) 555-0143" not in blob)
    ok("sample case id masked", "4471203" not in blob)
    ok("employer preserved", row["Employer"] == "Acme Diner LLC")
    ok("investigator preserved", "Jane Doe" in blob)

    # ---- the oracle fix, as a composition law: search over redacted rows
    shielded = [policy.redact_item(x) for x in e["items"]]
    ok("redacted view hides raw ssn from filter",
       store.filter_items(shielded, "123-45-6789") == [])
    ok("redacted view hides raw case id from ranking",
       store.rank_related(shielded, "4471203 zzz") == [])

    print("\nALL PASS — %d exports, %d checks" % (len(store.valid(exports)), _checks))


if __name__ == "__main__":
    main()
