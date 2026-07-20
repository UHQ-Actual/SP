"""Redaction policy.

Runs in the server, right before rows would be returned to the model — the last
point on the machine where the data is still local. Masks likely-sensitive
values by pattern (never by column name, so it also catches sensitive strings
that happen to sit inside a free-text notes field).

Patterns, applied in order:
  ssn          123-45-6789 / 123 45 6789 / 123.45.6789 -> ***-**-****
  ein          12-3456789                              -> **-*******
  card         4111 1111 1111 1111 (grouped)           -> ************1111
  phone        (202) 555-0143 (separator required)     -> [phone]
  long-number  7+ digit runs (ids, bare SSNs/cards)    -> ****1234  (keeps last 4)

Deliberate scope limits: emails and personal names are NOT redacted — the tool
exists to cross-reference people between Outlook and List rows, so masking them
would defeat its purpose. International / non-US phone formats are not covered.
Redaction is best-effort pattern matching, not DLP; do not rely on it as the
only control for genuinely classified data.

Add or tighten patterns here — this is the whole policy surface.
"""
from __future__ import annotations
import re
from typing import Any, List

# Lookarounds (not \b) throughout so runs glued to letters/punctuation are
# caught and mid-number runs are never partially matched.
_SSN = re.compile(r"(?<!\d)\d{3}[-.\s]\d{2}[-.\s]\d{4}(?!\d)")
_EIN = re.compile(r"(?<!\d)\d{2}-\d{7}(?!\d)")
# Grouped card numbers only (Visa/MC/Discover 4-4-4-x and Amex 4-6-5);
# ungrouped 13-19 digit runs fall through to the long-number rule.
_CARD = re.compile(
    r"(?<!\d)(?:\d{4}[ -]\d{4}[ -]\d{4}[ -]\d{1,4}|\d{4}[ -]\d{6}[ -]\d{5})(?!\d)"
)
# A separator or parens is required, so a bare 10-digit run falls to the
# long-number rule instead of being mislabeled as a phone.
_PHONE = re.compile(
    r"(?<!\d)(?:\+?1[-.\s]?)?(?:\(\d{3}\)\s?|\d{3}[-.\s])\d{3}[-.\s]?\d{4}(?!\d)"
)
_LONGNUM = re.compile(r"(?<!\d)\d{7,}(?!\d)")
_YYYYMMDD = re.compile(r"(19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])")


def _mask_tail(m: "re.Match") -> str:
    s = m.group(0)
    return "*" * (len(s) - 4) + s[-4:]


def _mask_card(m: "re.Match") -> str:
    digits = re.sub(r"\D", "", m.group(0))
    return "*" * (len(digits) - 4) + digits[-4:]


def _mask_longnum(m: "re.Match") -> str:
    s = m.group(0)
    if len(s) == 8 and _YYYYMMDD.fullmatch(s):
        return s  # YYYYMMDD date
    if m.start() > 0 and m.string[m.start() - 1] in "$€£.":
        return s  # currency amount / decimal tail / version component
    nxt = m.string[m.end():m.end() + 2]
    if len(nxt) == 2 and nxt[0] == "." and nxt[1].isdigit():
        return s  # head of a decimal number
    return _mask_tail(m)


# (name, compiled pattern, replacement) — order matters: most specific first.
_RULES = [
    ("ssn", _SSN, "***-**-****"),
    ("ein", _EIN, "**-*******"),
    ("card", _CARD, _mask_card),
    ("phone", _PHONE, "[phone]"),
    ("long-number (7+ digit ids)", _LONGNUM, _mask_longnum),
]

POLICY_SUMMARY: List[str] = [name for name, _p, _r in _RULES]


def redact_text(s: str) -> str:
    for _name, pat, repl in _RULES:
        s = pat.sub(repl, s)
    return s


def redact_value(v: Any) -> Any:
    if isinstance(v, bool):  # bool is an int subclass — check first, keep as-is
        return v
    if isinstance(v, (int, float)):
        # Numeric columns can hold SSNs / account numbers; inspect the digits.
        # Benign numbers keep their type; sensitive ones become masked strings.
        s = redact_text(str(v))
        return v if s == str(v) else s
    if isinstance(v, str):
        return redact_text(v)
    if isinstance(v, list):
        return [redact_value(x) for x in v]
    if isinstance(v, dict):
        return {k: redact_value(x) for k, x in v.items()}
    return v


def redact_item(item: dict) -> dict:
    return {k: redact_value(v) for k, v in item.items()}
