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
    "±": "$\\pm{}$",
    "×": "$\\times$",
    "≤": "$\\leq$",
    "≥": "$\\geq$",
    "→": "$\\rightarrow$",
    "←": "$\\leftarrow$",
    "≈": "$\\approx$",
    "≠": "$\\neq$",
    "—": "\\textemdash",
    "✓": "\\checkmark",
    "✗": "\\ding{55}",
    "↑": "$\\uparrow$",
    "↓": "$\\downarrow$",
    "★": "$\\bigstar$",
    "∼": "$\\sim$",
    # Arrows
    "⇒": "$\\Rightarrow$",
    "⇐": "$\\Leftarrow$",
    "⇑": "$\\Uparrow$",
    "⇓": "$\\Downarrow$",
    # Greek lowercase
    "α": "$\\alpha$",
    "β": "$\\beta$",
    "γ": "$\\gamma$",
    "δ": "$\\delta$",
    "ε": "$\\epsilon$",
    "ζ": "$\\zeta$",
    "η": "$\\eta$",
    "θ": "$\\theta$",
    "κ": "$\\kappa$",
    "λ": "$\\lambda$",
    "μ": "$\\mu$",
    "π": "$\\pi$",
    "ρ": "$\\rho$",
    "σ": "$\\sigma$",
    "τ": "$\\tau$",
    "φ": "$\\phi$",
    "ω": "$\\omega$",
    # Greek uppercase
    "Σ": "$\\Sigma$",
    "Ω": "$\\Omega{}$",
    # Math symbols
    "∞": "$\\infty$",
    "·": "$\\cdot$",
    "⋯": "$\\cdots$",
    "ℓ": "$\\ell$",
    # Special symbols
    "△": "$\\triangle$",
    "▼": "$\\blacktriangledown$",
    "†": "\\textdagger{}",
    "‡": "\\textdaggerdbl{}",
    "§": "\\S{}",
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


# Standard LaTeX/xcolor named colors → RGB hex
_LATEX_COLORS = {
    "red": "FF0000", "blue": "0000FF", "green": "008000", "black": "000000",
    "white": "FFFFFF", "gray": "808080", "grey": "808080", "cyan": "00FFFF",
    "magenta": "FF00FF", "yellow": "FFFF00", "orange": "FF8000",
    "purple": "800080", "brown": "804000", "violet": "8000FF",
    "pink": "FFC0CB", "lime": "00FF00", "olive": "808000", "teal": "008080",
    "darkgray": "404040", "lightgray": "C0C0C0",
    # xcolor dvipsnames
    "ForestGreen": "009B55", "NavyBlue": "006EB8", "RoyalBlue": "0071BC",
    "MidnightBlue": "003262", "SkyBlue": "46C5DD", "TealBlue": "00827F",
    "Cerulean": "00A2E3", "ProcessBlue": "00B0F0", "Aquamarine": "00B5BE",
    "BlueGreen": "00B3B8", "Turquoise": "00B4CD", "Emerald": "00A99D",
    "JungleGreen": "00A99A", "PineGreen": "008B72", "SeaGreen": "3FBC9D",
    "OliveGreen": "3C8031", "LimeGreen": "8DC73F", "YellowGreen": "98CC70",
    "GreenYellow": "F7F206", "SpringGreen": "C6DC67", "Green": "00A64F",
    "Yellow": "FFF200", "Goldenrod": "FFDF00", "Dandelion": "FDBC42",
    "Apricot": "FBB982", "Peach": "F7965A", "Melon": "F89E7B",
    "YellowOrange": "FAA21A", "Orange": "F7941D", "BurntOrange": "F7941D",
    "Bittersweet": "C84B0F", "RedOrange": "F26035", "OrangeRed": "F26035",
    "Red": "ED1B23", "BrickRed": "B6321C", "Salmon": "F69289",
    "WildStrawberry": "EE2967", "Rhodamine": "EF559F", "RubineRed": "ED017D",
    "CarnationPink": "F7A7C4", "Lavender": "EFC5E0", "Thistle": "D9B0D4",
    "Orchid": "AF72B0", "DarkOrchid": "A55EA5", "Fuchsia": "8C368C",
    "Mulberry": "A93C93", "RedViolet": "A1246B", "VioletRed": "EF1261",
    "Maroon": "AF3235", "Mahogany": "A52A2A", "Sepia": "671800",
    "Brown": "792500", "RawSienna": "974006", "Tan": "DB9065",
    "Plum": "92268F", "RoyalPurple": "613F99", "BlueViolet": "473992",
    "Violet": "58429B", "Periwinkle": "7977B8", "CadetBlue": "626D9F",
    "CornflowerBlue": "92A8D1", "Cyan": "00AEEF", "Magenta": "EC008C",
    "Purple": "99479B", "Gray": "949698",
    # CSS extended colors commonly used in LaTeX
    "darkgreen": "006400", "darkblue": "00008B", "darkred": "8B0000",
    "lightblue": "ADD8E6", "lightgreen": "90EE90", "steelblue": "4682B4",
    "royalblue": "4169E1", "forestgreen": "228B22", "navyblue": "000080",
    "crimson": "DC143C", "coral": "FF7F50", "gold": "FFD700",
    "indigo": "4B0082", "tomato": "FF6347", "hotpink": "FF69B4",
    "deepskyblue": "00BFFF", "dodgerblue": "1E90FF", "firebrick": "B22222",
    "darkorange": "FF8C00", "darkviolet": "9400D3", "slateblue": "6A5ACD",
    "slategray": "708090", "slategrey": "708090",
}


def _latex_color_to_hex(color: str) -> str | None:
    """Convert LaTeX color spec to '#RRGGBB' hex string.

    Supports: named colors, xcolor mixing (e.g. 'gray!20'),
    hex codes, [RGB]{r,g,b} format.
    """
    color = color.strip()
    if not color:
        return None
    # Already #RRGGBB
    if color.startswith("#") and len(color) == 7:
        return color.upper()
    # Bare 6-digit hex
    if len(color) == 6 and all(c in "0123456789abcdefABCDEF" for c in color):
        return f"#{color.upper()}"
    # xcolor mixing: "color!percent"
    if "!" in color:
        parts = color.split("!")
        base_name = parts[0].strip()
        base = _LATEX_COLORS.get(base_name) or _LATEX_COLORS.get(base_name.lower())
        if base and len(parts) >= 2:
            try:
                pct = float(parts[1]) / 100.0
            except ValueError:
                return None
            br, bg, bb = int(base[0:2], 16), int(base[2:4], 16), int(base[4:6], 16)
            r = int(br * pct + 255 * (1 - pct))
            g = int(bg * pct + 255 * (1 - pct))
            b = int(bb * pct + 255 * (1 - pct))
            return f"#{r:02X}{g:02X}{b:02X}"
        return None
    # Named color
    hex_val = _LATEX_COLORS.get(color) or _LATEX_COLORS.get(color.lower())
    if hex_val:
        return f"#{hex_val}"
    return None


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
