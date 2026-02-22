"""Utility functions for LaTeX escaping and color conversion."""

from __future__ import annotations

import re

_LATEX_SPECIAL = {
    "&": r"\&",
    "%": r"\%",
    "$": r"\$",
    "#": r"\#",
    "_": r"\_",
    "{": r"\{",
    "}": r"\}",
    "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
    "\\": r"\textbackslash{}",
}

_LATEX_RE = re.compile("|".join(re.escape(k) for k in _LATEX_SPECIAL))


_UNICODE_TO_LATEX = {
    "±": "$\\pm$",
    "×": "$\\times$",
    "≤": "$\\leq$",
    "≥": "$\\geq$",
    "→": "$\\rightarrow$",
    "←": "$\\leftarrow$",
    "≈": "$\\approx$",
    "≠": "$\\neq$",
    "—": "\\textemdash",
}

_UNICODE_RE = re.compile("|".join(re.escape(k) for k in _UNICODE_TO_LATEX))


def latex_escape(text: str) -> str:
    """Escape special LaTeX characters in text.

    Unicode math symbols (±, ×, ≤, ≥, etc.) are auto-converted to LaTeX commands.
    """
    if not isinstance(text, str):
        text = str(text)
    # Extract Unicode math symbols before escaping
    parts = _UNICODE_RE.split(text)
    symbols = _UNICODE_RE.findall(text)
    result = []
    for i, part in enumerate(parts):
        result.append(_LATEX_RE.sub(lambda m: _LATEX_SPECIAL[m.group()], part))
        if i < len(symbols):
            result.append(_UNICODE_TO_LATEX[symbols[i]])
    return "".join(result)


def hex_to_latex_color(hex_color: str) -> str:
    """Convert hex color like '#FF0000' to LaTeX xcolor RGB spec."""
    h = hex_color.lstrip("#")
    if len(h) == 6:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return f"{r},{g},{b}"
    return "0,0,0"


def format_number(value, fmt: str, strip_leading_zero: bool = True) -> str:
    """Format a numeric value with a format spec like '.2f' or '.1%'.

    Leading zeros are stripped for values in (-1, 1) — e.g. 0.451 → .451.
    """
    try:
        v = float(value)
        s = format(v, fmt)
        if strip_leading_zero and -1 < v < 1:
            s = s.replace("0.", ".", 1) if s.startswith("0.") else s.replace("-0.", "-.", 1)
        return s
    except (ValueError, TypeError):
        return str(value)
