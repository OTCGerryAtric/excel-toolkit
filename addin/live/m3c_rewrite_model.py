"""Clear Sheet References (M3c): a Python model of TkFormulas.Tk_RewriteSheetPrefix, step for step.

The offline tests run the shared cases (m3c_cases.py) through it and check the VBA follows the same rules;
the live check runs the same cases through the VBA function itself and uses this model to say which
formulas in a real workbook should change.
"""

from __future__ import annotations

import re


def _is_name_char(ch: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9_.]", ch)) or ord(ch) > 127


def _quote_end(formula: str, start: int, quote: str) -> int:
    k = start + 1
    while k < len(formula):
        if formula[k] == quote:
            if k + 1 < len(formula) and formula[k + 1] == quote:
                k += 2
                continue
            return k
        k += 1
    return len(formula) - 1


def _word_end(formula: str, i: int) -> int:
    j = i
    while j + 1 < len(formula) and _is_name_char(formula[j + 1]):
        j += 1
    return j


def _token_start(formula: str, last: int) -> int:
    k = last
    while k >= 0 and (formula[k] == "$" or _is_name_char(formula[k])):
        k -= 1
    return k + 1


def _is_cell_reference(token: str) -> bool:
    m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})", token)
    if not m:
        return False
    column = 0
    for ch in m.group(1).upper():
        column = column * 26 + ord(ch) - 64
    return column <= 16384 and int(m.group(2)) <= 1048576


def _prefix_name_before(formula: str, bang: int) -> str:
    if bang < 1:
        return ""
    if formula[bang - 1] == "'":
        closing, k = bang - 1, bang - 2
        while k >= 0:
            if formula[k] == "'":
                if k > 0 and formula[k - 1] == "'":
                    k -= 2
                    continue
                break
            k -= 1
        return formula[k + 1:closing].replace("''", "'") if k >= 0 else ""
    k = bang - 1
    while k >= 0 and _is_name_char(formula[k]):
        k -= 1
    return "" if k >= 0 and formula[k] == "]" else formula[k + 1:bang]


def _range_half_before(formula: str, colon: int, sheet: str) -> bool:
    if colon < 1:
        return False
    if formula[colon - 1] == ")":
        return True
    k = _token_start(formula, colon - 1)
    if not _is_cell_reference(formula[k:colon]):
        return False
    if k > 0 and formula[k - 1] == "!":
        return _prefix_name_before(formula, k - 1).lower() == sheet.lower()
    return True


def _own_prefix(formula: str, i: int, name: str, sheet: str) -> bool:
    prev = formula[i - 1] if i > 0 else ""
    return name.lower() == sheet.lower() and prev not in ("]", "#") \
        and (prev != ":" or _range_half_before(formula, i - 1, sheet))


def _prefix_at(formula: str, p: int) -> bool:
    if p >= len(formula):
        return False
    if formula[p] == "'":
        j = _quote_end(formula, p, "'")
    elif _is_name_char(formula[p]):
        j = _word_end(formula, p)
    else:
        return False
    return j + 1 < len(formula) and formula[j + 1] == "!"


def _has_ambiguous_colon(formula: str) -> bool:
    i = 0
    while i < len(formula):
        ch = formula[i]
        if ch in "\"'":
            i = _quote_end(formula, i, ch) + 1
            continue
        if ch == ":" and _prefix_at(formula, i + 1):
            if i == 0:
                return True
            if formula[i - 1] != ")":
                k = _token_start(formula, i - 1)
                if k == i:
                    return True
                if not _is_cell_reference(formula[k:i]) and (k == 0 or formula[k - 1] != "!"):
                    return True
        i += 1
    return False


def _remove_prefixes(formula: str, sheet: str) -> str:
    out, i, n = [], 0, len(formula)
    while i < n:
        ch = formula[i]
        if ch == '"':
            j = _quote_end(formula, i, '"')
        elif ch == "'":
            j = _quote_end(formula, i, "'")
            if j + 1 < n and formula[j + 1] == "!" and _own_prefix(formula, i, formula[i + 1:j].replace("''", "'"), sheet):
                i = j + 2
                continue
        elif _is_name_char(ch):
            j = _word_end(formula, i)
            if j + 1 < n and formula[j + 1] == "!" and _own_prefix(formula, i, formula[i:j + 1], sheet):
                i = j + 2
                continue
        else:
            j = i
        out.append(formula[i:j + 1])
        i = j + 1
    return "".join(out)


def rewrite_sheet_prefix(formula: str, sheet: str) -> tuple[str, bool]:
    """(result, ambiguous): the same steps as TkFormulas.Tk_RewriteSheetPrefix, 0-based."""
    result = _remove_prefixes(formula, sheet)
    if result != formula and (_has_ambiguous_colon(formula) or _has_ambiguous_colon(result)):
        return formula, True
    return result, False


def remove_sheet_prefix(formula: str, sheet: str) -> str:
    return rewrite_sheet_prefix(formula, sheet)[0]
