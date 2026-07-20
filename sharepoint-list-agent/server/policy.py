"""Redaction policy.

Runs in the server, right before rows would be returned to the model — the last
point on the machine where the data is still local. Masks likely-sensitive
values by pattern (never by column name, so it also catches sensitive strings
that happen to sit inside a free-text notes field).

Patterns, applied in order:
  ssn              123-45-6789            -> ***-**-****
  phone            (202) 555-0143         -> [phone]
  long-number      case / id runs 7-10    -> ****1234  (keeps last 4)

Add or tighten patterns here — this is the whole policy surface.
"""
from __future__ import annotations
import re
from typing import Any, List

_SSN = re.compile(r"\b\d{3}-\d{2}-\d{4}\b")
# Lookarounds (not \b) so an optional leading "(" is consumed and mid-number
# runs are not partially matched.
_PHONE = re.compile(r"(?<!\d)(?:\+?1[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}(?!\d)")
_LONGNUM = re.compile(r"\b\d{7,10}\b")


def _mask_tail(m: "re.Match") -> str:
    s = m.group(0)
    return "*" * (len(s) - 4) + s[-4:]


# (name, compiled pattern, replacement) — order matters: most specific first.
_RULES = [
    ("ssn", _SSN, "***-**-****"),
    ("phone", _PHONE, "[phone]"),
    ("long-number (case/id)", _LONGNUM, _mask_tail),
]

POLICY_SUMMARY: List[str] = [name for name, _p, _r in _RULES]


def redact_text(s: str) -> str:
    for _name, pat, repl in _RULES:
        s = pat.sub(repl, s)
    return s


def redact_value(v: Any) -> Any:
    if isinstance(v, str):
        return redact_text(v)
    if isinstance(v, list):
        return [redact_value(x) for x in v]
    if isinstance(v, dict):
        return {k: redact_value(x) for k, x in v.items()}
    return v


def redact_item(item: dict) -> dict:
    return {k: redact_value(v) for k, v in item.items()}
