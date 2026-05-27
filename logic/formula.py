# -*- coding: utf-8 -*-
"""Quantity formula parser.

Supports formulas like:
    "=30*2"        -> quantity=60,    notes="30*2(2pcs)"
    "=10+23+123"   -> quantity=156,   notes="10+23+123(3pcs)"
    "=10*5+10"     -> quantity=60,    notes="10*5+10(6pcs)"
    "=9*9+1000000" -> quantity=1000081, notes="9*9+1000000(10pcs)"

Rules:
  - Split by '+', each term is one "piece".
  - If a term contains '*', the number AFTER '*' is the count of pieces.
  - If a term is a plain number (no '*'), it counts as 1 piece.
  - Total pieces = sum of piece counts from all terms.
  - Quantity = arithmetic evaluation of the expression.
  - Space in formula is ignored.
"""

import re


def parse_formula(raw):
    """Parse a quantity formula string.

    Args:
        raw: A formula string, with or without leading '='.
             e.g. "30*2", "=10+23+123", "=10*5+10"

    Returns:
        dict with keys:
            quantity (float): calculated total
            notes (str): formatted notes like "30*2(2个)"
            piece_count (int): total number of pieces
            error (str or None): error message if parsing failed
    """
    if raw is None:
        return _empty_result()

    s = raw.strip()
    if not s:
        return _empty_result()

    # strip leading '='
    if s.startswith("="):
        s = s[1:].strip()

    if not s:
        return _empty_result()

    # remove all whitespace
    s = s.replace(" ", "").replace("\t", "")

    # validate: only digits, +, *, . allowed
    if not re.match(r'^[\d+\-*/.]+$', s):
        return {
            "quantity": 0,
            "notes": "",
            "piece_count": 0,
            "error": "Formula contains invalid characters"
        }

    # split by '+'
    terms = s.split("+")

    total_pieces = 0
    for term in terms:
        if not term:
            continue
        if "*" in term:
            # get the number after the last '*'
            parts = term.split("*")
            try:
                count_part = parts[-1]
                total_pieces += int(float(count_part))
            except (ValueError, IndexError):
                pass
        else:
            total_pieces += 1

    # evaluate expression safely
    try:
        # Only allow digits, +, *, -, /, .
        quantity = eval(s, {"__builtins__": {}}, {})
        quantity = float(quantity)
    except Exception:
        return {
            "quantity": 0,
            "notes": "",
            "piece_count": 0,
            "error": "Cannot evaluate expression"
        }

    # build notes
    if total_pieces > 0:
        notes = u"{}（{}个）".format(s, total_pieces)
    else:
        notes = ""

    return {
        "quantity": quantity,
        "notes": notes,
        "piece_count": total_pieces,
        "error": None
    }


def expand_formula(raw):
    """Expand a quantity formula into individual label items.

    Returns list of dicts with keys:
        quantity (float): the quantity value for this label
        count (int): how many labels to print with this quantity
    """
    if not raw:
        return []
    s = raw.strip()
    if s.startswith("="):
        s = s[1:].strip()
    if not s:
        return []
    s = s.replace(" ", "").replace("\t", "")
    if not re.match(r'^[\d+\-*/.]+$', s):
        return []
    items = []
    for term in s.split("+"):
        if not term:
            continue
        if "*" in term:
            parts = term.split("*")
            try:
                qty = float(eval(parts[0], {"__builtins__": {}}, {}))
                cnt = int(float(parts[-1]))
                if cnt > 0:
                    items.append({"quantity": qty, "count": cnt})
            except (ValueError, IndexError, SyntaxError):
                items.append({"quantity": 0, "count": 1})
        else:
            try:
                qty = float(eval(term, {"__builtins__": {}}, {}))
                items.append({"quantity": qty, "count": 1})
            except (ValueError, SyntaxError):
                pass
    return items


def _empty_result():
    return {
        "quantity": 0,
        "notes": "",
        "piece_count": 0,
        "error": None
    }
