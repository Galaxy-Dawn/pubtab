"""LaTeX table parser — converts .tex to TableData."""

from __future__ import annotations

import re
from typing import List, Optional, Tuple

from .models import Cell, CellStyle, TableData


def read_tex(tex: str) -> TableData:
    """Parse a LaTeX table string into TableData.

    Handles: multicolumn, multirow, textbf, textit, underline, diagbox.
    """
    # Extract tabular body
    body = _extract_tabular_body(tex)
    if body is None:
        raise ValueError("No tabular environment found")

    # Split into rows (by \\), filter out rules
    raw_rows = _split_rows(body)

    # Parse each row into cells
    parsed_rows: List[List[Cell]] = []
    for row_str in raw_rows:
        cells = _parse_row(row_str)
        if cells is not None:
            parsed_rows.append(cells)

    if not parsed_rows:
        raise ValueError("No data rows found")

    num_cols = max(
        sum(c.colspan for c in row) for row in parsed_rows
    )

    # Expand rows to full width with placeholder cells
    expanded = []
    for row in parsed_rows:
        full_row = _expand_row(row, num_cols)
        expanded.append(full_row)

    # Handle negative multirow: move content from bottom to top of merged range
    for r, row in enumerate(expanded):
        for c, cell in enumerate(row):
            if cell.rowspan < 0:
                span = abs(cell.rowspan)
                target_r = r - span + 1
                if target_r >= 0:
                    expanded[target_r][c] = Cell(
                        value=cell.value, style=cell.style,
                        rowspan=span, colspan=cell.colspan,
                    )
                    expanded[r][c] = Cell(value="", style=CellStyle())

    # Detect header rows (rows before first data-like row)
    header_rows = _detect_header_rows(expanded, num_cols)

    return TableData(
        cells=expanded,
        num_rows=len(expanded),
        num_cols=num_cols,
        header_rows=header_rows,
    )


def _strip_comments(tex: str) -> str:
    """Remove LaTeX % comments (but not escaped \\%)."""
    return re.sub(r"(?<!\\)%[^\n]*", "", tex)


def _extract_tabular_body(tex: str) -> Optional[str]:
    """Extract content between \\begin{tabular} and \\end{tabular}."""
    tex = _strip_comments(tex)
    # Match column spec with nested braces (e.g. {@{} p{0.9cm} ...@{}})
    m = re.search(
        rf"\\begin\{{tabular\}}\{{({_NESTED})\}}(.*?)\\end\{{tabular\}}",
        tex, re.DOTALL,
    )
    return m.group(2).strip() if m else None


def _split_rows(body: str) -> List[str]:
    """Split tabular body into row strings, skipping rule commands."""
    # Split by \\ respecting brace nesting (don't split inside {})
    parts = _split_by_double_backslash(body)
    rows = []
    for part in parts:
        original = part.strip()
        # Remove rule commands within a row chunk (with optional arguments like [0.5pt])
        s = re.sub(r"\\(toprule|bottomrule|midrule)(?:\[[^\]]*\])?\s*", "", original)
        s = re.sub(r"\\cmidrule(\([^)]*\))?\{[^}]*\}\s*", "", s)
        s = s.strip()
        if s:
            rows.append(s)
        elif not original:
            # Preserve empty rows (needed for multirow placeholders)
            rows.append("")
        # else: rule-only row, skip
    # Trim leading/trailing empty rows
    while rows and rows[-1] == "":
        rows.pop()
    while rows and rows[0] == "":
        rows.pop(0)
    return rows


def _split_by_double_backslash(s: str) -> List[str]:
    """Split string by \\\\ respecting brace nesting."""
    parts = []
    depth = 0
    current = []
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == "{":
            depth += 1
            current.append(ch)
        elif ch == "}":
            depth -= 1
            current.append(ch)
        elif ch == "\\" and i + 1 < len(s) and s[i + 1] == "\\" and depth == 0:
            parts.append("".join(current))
            current = []
            i += 2
            continue
        else:
            current.append(ch)
        i += 1
    parts.append("".join(current))
    return parts


_NESTED = r"(?:[^{}]|\{(?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*\})*"
_MULTICOLUMN_RE = re.compile(
    rf"\\multicolumn\{{(\d+)\}}\{{({_NESTED})\}}\s*\{{({_NESTED})\}}"
)
_MULTIROW_RE = re.compile(
    rf"\\multirow(?:\[[^\]]*\])?\{{(-?[\d.]+)\}}(?:\[[^\]]*\])?(?:\{{[^}}]*\}}|\*)\s*\{{({_NESTED})\}}"
)


def _parse_row(row_str: str) -> Optional[List[Cell]]:
    """Parse a single row string into a list of Cells."""
    # Split by & respecting brace nesting
    parts = _split_by_ampersand(row_str)
    cells = []
    for part in parts:
        cell = _parse_cell(part.strip())
        cells.append(cell)
    return cells if cells else None


def _split_by_ampersand(s: str) -> List[str]:
    """Split string by & respecting LaTeX brace nesting and escaped \\&."""
    parts = []
    depth = 0
    current = []
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == "{":
            depth += 1
            current.append(ch)
        elif ch == "}":
            depth -= 1
            current.append(ch)
        elif ch == "\\" and i + 1 < len(s) and s[i + 1] == "&":
            current.append("\\&")
            i += 2
            continue
        elif ch == "&" and depth == 0:
            parts.append("".join(current))
            current = []
        else:
            current.append(ch)
        i += 1
    parts.append("".join(current))
    return parts


def _strip_outer_braces(text: str) -> str:
    """Strip redundant outer braces: '{\\underline{.451}}' → '\\underline{.451}'."""
    while True:
        t = text.strip()
        if len(t) >= 2 and t[0] == "{" and t[-1] == "}":
            # Verify the braces are matched (not part of inner content)
            depth = 0
            for i, ch in enumerate(t):
                if ch == "{":
                    depth += 1
                elif ch == "}":
                    depth -= 1
                if depth == 0 and i < len(t) - 1:
                    break  # Closing brace isn't the last char
            else:
                # The outer braces wrap the entire content
                text = t[1:-1].strip()
                continue
        break
    return text


def _parse_cell(text: str) -> Cell:
    """Parse a single cell text into a Cell object."""
    # Normalize whitespace (newlines between \multicolumn{}{} and {content})
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return Cell(value="", style=CellStyle())

    # Strip rule commands in a loop (multiple may precede content)
    for _ in range(10):
        prev = text
        text = re.sub(r"^\s*\[[\d.]+\w*\]\s*", "", text)  # [0.8pt] width spec
        text = re.sub(r"^\\(hdashline|hline|thickhline|Xhline(?:\{[^}]*\}|[\d.]*)|addlinespace(?:\[[^\]]*\])?)\s*", "", text)
        text = re.sub(r"^\\(cmidrule|cdashline|cline)(?:\([^)]*\))?(?:\[[^\]]*\])?\{[^}]*\}\s*", "", text)
        if text == prev:
            break

    colspan = 1
    rowspan = 1
    alignment = "center"
    bg_color = None
    text_color = None
    rotation = 0

    # Extract colors FIRST (before multirow/multicolumn which use .match())
    for _ in range(3):  # handle multiple color commands
        m = re.search(r"\\cellcolor(?:\[[^\]]*\])?\{([^}]*)\}", text)
        if m:
            bg_color = m.group(1)
            text = (text[:m.start()] + text[m.end():]).strip()
            continue
        m = re.search(r"\\rowcolor(?:\[[^\]]*\])?\{([^}]*)\}", text)
        if m:
            bg_color = m.group(1)
            text = (text[:m.start()] + text[m.end():]).strip()
            continue
        break

    # Strip leading font commands that block .match() for multirow/multicolumn
    text = re.sub(r"^\\(bf|it|rm|tt|sc|bfseries|itshape|rmfamily|ttfamily|scshape|boldmath|bm)\b\s*", "", text).strip()

    # Extract multicolumn
    m = _MULTICOLUMN_RE.match(text)
    if m:
        colspan = int(m.group(1))
        alignment = m.group(2).strip()
        text = m.group(3).strip()

    # Extract multirow
    m = _MULTIROW_RE.match(text)
    if m:
        rowspan = round(float(m.group(1)))  # negative = content at bottom
        text = m.group(2).strip()

    # Extract \textcolor{color}{content}
    m = re.search(r"\\textcolor(?:\[[^\]]*\])?\{([^}]*)\}\{([^}]*)\}", text)
    if m:
        text_color = m.group(1)
        text = text[:m.start()] + m.group(2) + text[m.end():]

    # Extract \colorbox{color}{content} (loop for multiple occurrences)
    for _ in range(10):
        m = re.search(r"\\colorbox(?:\[[^\]]*\])?\{([^}]*)\}\{([^}]*)\}", text)
        if m:
            text = text[:m.start()] + m.group(2) + text[m.end():]
            continue
        break

    # Extract \rotatebox{angle}{content} or \rotatebox[origin=c]{angle}{content}
    m = re.search(r"\\rotatebox(?:\[[^\]]*\])?\{(\d+)\}\{([^}]*)\}", text)
    if m:
        rotation = int(m.group(1))
        text = text[:m.start()] + m.group(2) + text[m.end():]

    # Strip redundant outer braces: {{\underline{.451}}} → \underline{.451}
    text = _strip_outer_braces(text)

    # Parse formatting and extract value
    bold, italic, underline, diagbox_parts, value = _parse_formatting(text)

    # Try to parse as number
    num_value = _try_parse_number(value)

    style = CellStyle(
        bold=bold,
        italic=italic,
        underline=underline,
        alignment=alignment,
        diagbox=diagbox_parts,
        bg_color=bg_color,
        color=text_color,
        rotation=rotation,
    )

    return Cell(
        value=num_value if num_value is not None else value,
        style=style,
        rowspan=rowspan,
        colspan=colspan,
    )


def _parse_formatting(text: str) -> Tuple[bool, bool, bool, Optional[List[str]], str]:
    """Extract bold/italic/underline/diagbox and return clean value.

    Returns: (bold, italic, underline, diagbox_parts, clean_text)
    """
    bold = False
    italic = False
    underline = False
    diagbox_parts = None

    # Iteratively unwrap formatting commands
    changed = True
    while changed:
        changed = False
        t = text.strip()

        m = re.fullmatch(r"\\textbf\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}", t)
        if m:
            bold = True
            text = m.group(1).strip()
            changed = True
            continue

        m = re.fullmatch(r"\\textit\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}", t)
        if m:
            italic = True
            text = m.group(1).strip()
            changed = True
            continue

        m = re.fullmatch(r"\\emph\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}", t)
        if m:
            italic = True
            text = m.group(1).strip()
            changed = True
            continue

        m = re.fullmatch(r"\\underline\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}", t)
        if m:
            underline = True
            text = m.group(1).strip()
            changed = True
            continue

    # Check for diagbox
    m = re.fullmatch(r"\\diagbox\{([^}]*)\}\{([^}]*)\}", text.strip())
    if m:
        diagbox_parts = [m.group(1), m.group(2)]
        text = ""

    # Clean up LaTeX artifacts
    text = _clean_latex(text)

    return bold, italic, underline, diagbox_parts, text


def _clean_latex(text: str) -> str:
    """Remove common LaTeX commands and convert back to plain text."""
    # Reverse writer escape sequences (latex_escape produces these)
    text = text.replace("\\textasciicircum{}", "^")
    text = text.replace("\\textasciicircle{}", "^")
    text = text.replace("\\textasciitilde{}", "~")
    text = text.replace("\\textbackslash{}", "\\")
    # \makecell{Things-\\EEG} → Things-\nEEG (preserve line breaks as newline)
    text = re.sub(r"\\makecell(?:\[[^\]]*\])?\{([^}]*)\}", lambda m: m.group(1).replace("\\\\", "\n"), text)
    # \specialcell[t]{content} → content (like makecell)
    text = re.sub(r"\\specialcell(?:\[[^\]]*\])?\{([^}]*)\}", lambda m: m.group(1).replace("\\\\", "\n"), text)
    # \shortstack{content} → content (like makecell)
    text = re.sub(r"\\shortstack(?:\[[^\]]*\])?\{([^}]*)\}", lambda m: m.group(1).replace("\\\\", "\n"), text)
    # \var{$\pm$.005} → $\pm$.005 (strip wrapper, keep content)
    text = re.sub(r"\\var\{([^}]*)\}", r"\1", text)
    # Strip math mode: $D_\text{stage 1}$ → D_stage 1
    text = re.sub(r"\$([^$]*)\$", lambda m: re.sub(r"\\text\{([^}]*)\}", r"\1", m.group(1)), text)
    # Math symbols
    text = text.replace("\\pm", "±")
    text = text.replace("$\\times$", "×")
    text = text.replace("\\texttimes", "×")
    text = text.replace("\\times", "×")
    text = text.replace("\\textemdash", "—")
    text = text.replace("\\uparrow", "↑")
    text = text.replace("\\downarrow", "↓")
    text = text.replace("\\Downarrow", "⇓")
    text = text.replace("\\Uparrow", "⇑")
    text = text.replace("\\rightarrow", "→")
    text = text.replace("\\leftarrow", "←")
    text = text.replace("\\Rightarrow", "⇒")
    text = text.replace("\\Leftarrow", "⇐")
    text = text.replace("\\alpha", "α")
    text = text.replace("\\beta", "β")
    text = text.replace("\\gamma", "γ")
    text = re.sub(r"\\mu\b", "μ", text)
    text = text.replace("\\sigma", "σ")
    text = text.replace("\\delta", "δ")
    text = text.replace("\\epsilon", "ε")
    text = text.replace("\\theta", "θ")
    text = text.replace("\\lambda", "λ")
    text = text.replace("\\pi", "π")
    text = text.replace("\\omega", "ω")
    text = text.replace("\\tau", "τ")
    text = text.replace("\\triangle", "△")
    text = text.replace("\\star", "★")
    text = text.replace("\\textdagger", "†")
    text = text.replace("\\dagger", "†")
    text = text.replace("\\ddag", "‡")
    text = re.sub(r"\\dag(?![a-zA-Z])", "†", text)
    text = re.sub(r"\\S(?![a-zA-Z])", "§", text)
    text = text.replace("\\Sigma", "Σ")
    text = text.replace("\\Omega", "Ω")
    text = text.replace("\\sim", "∼")
    text = text.replace("\\infty", "∞")
    text = text.replace("\\leq", "≤")
    text = text.replace("\\geq", "≥")
    text = text.replace("\\neq", "≠")
    text = text.replace("\\approx", "≈")
    text = text.replace("\\cdot", "·")
    text = text.replace("\\ell", "ℓ")
    # Superscript: ^{content} → content, ^X → X
    text = re.sub(r"\^\\text\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\^\\dagger", "†", text)
    text = re.sub(r"\^\\mathrm\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\^\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\^([a-zA-Z0-9*†]+)", r"\1", text)
    # Subscript: _{content} → content, _X → X
    text = re.sub(r"(?<!\\)_\{([^}]*)\}", r"\1", text)
    text = re.sub(r"(?<!\\)_([a-zA-Z0-9])(?![a-zA-Z0-9])", r"\1", text)
    # Special symbols
    text = text.replace("\\Checkmark", "✓")
    text = text.replace("\\checkmark", "✓")
    text = text.replace("\\cmark", "✓")
    text = text.replace("\\XSolidBrush", "✗")
    text = text.replace("\\XSolid", "✗")
    text = text.replace("\\xmark", "✗")
    text = text.replace("\\ding{51}", "✓")
    text = text.replace("\\ding{52}", "✓")
    text = text.replace("\\ding{55}", "✗")
    # \ding without braces (after brace stripping)
    text = re.sub(r"\\ding\s*5[12]", "✓", text)
    text = re.sub(r"\\ding\s*55", "✗", text)
    text = text.replace("\\bigstar", "★")
    text = text.replace("\\blacktriangledown", "▼")
    text = text.replace("\\compareyes", "✓")
    text = text.replace("\\comparepartially", "∼")
    text = text.replace("\\compareno", "✗")
    text = text.replace("\\tick", "✓")
    text = text.replace("\\cross", "✗")
    text = text.replace("\\varepsilon", "ε")
    text = text.replace("\\phi", "φ")
    text = text.replace("\\rho", "ρ")
    text = text.replace("\\kappa", "κ")
    text = text.replace("\\eta", "η")
    text = text.replace("\\zeta", "ζ")
    # Non-breaking space
    text = text.replace("~", " ")
    # Remove citations: \citep{...}, \cite{...}, \citet{...}, \citeyearpar{...}, etc.
    text = re.sub(r"~?\\cite[a-z]*\{[^}]*\}", "", text)
    # Remove \\ inside cells (line break in makecell) — must come before '\ ' conversion
    text = text.replace("\\\\", " ")
    # Remove \, and other spacing; convert '\ ' (forced space) to space
    text = text.replace("\\ ", " ")
    text = re.sub(r"\\[,;!]", "", text)
    text = re.sub(r"\\q?quad\s*", " ", text)
    # Remove font style commands: \bf, \rm, \it, \tt, \sc, \bfseries, \boldmath, etc.
    text = re.sub(r"\\(bf|rm|it|tt|sc|bfseries|rmfamily|itshape|ttfamily|scshape|boldmath|unboldmath|bm)\b\s*", "", text)
    # Remove \textsc{...}, \texttt{...}, \textrm{...}, \textit{...}, \textbf{...}
    text = re.sub(r"\\text(sc|tt|rm|it|bf|sf|sl)\{([^}]*)\}", r"\2", text)
    # Remove \text{content} (plain, no style suffix)
    text = re.sub(r"\\text\{([^}]*)\}", r"\1", text)
    # Remove \textsuperscript{...}
    text = re.sub(r"\\textsuperscript\{([^}]*)\}", r"\1", text)
    # Remove \mathbf{...}, \mathcal{...}, etc.
    text = re.sub(r"\\math(bf|cal|it|rm|tt|sf)\{([^}]*)\}", r"\2", text)
    text = re.sub(r"\\boldsymbol\{([^}]*)\}", r"\1", text)
    # Remove \mathrm{...}, \textrm{...}
    text = re.sub(r"\\mathrm\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\textrm\{([^}]*)\}", r"\1", text)
    # Remove size commands
    text = re.sub(r"\\(scriptsize|small|footnotesize|normalsize|tiny|large|Large|LARGE|huge|Huge)\b\s*", "", text)
    # Remove \hline, \hdashline, \cline{n-m}, \Xhline{width}, \addlinespace[width], \thickhline
    text = re.sub(r"\\hdashline\s*", "", text)
    text = re.sub(r"\\hline\s*", "", text)
    text = re.sub(r"\\thickhline\s*", "", text)
    text = re.sub(r"\\Xhline\{[^}]*\}\s*", "", text)
    text = re.sub(r"\\Xhline[\d.]+\w*\s*", "", text)
    text = re.sub(r"\\addlinespace(?:\[[^\]]*\])?\s*", "", text)
    text = re.sub(r"\\cline\{[^}]*\}\s*", "", text)
    # Remove \noalign{...} (vertical spacing between rows)
    text = re.sub(r"\\noalign\{[^}]*\}\s*", "", text)
    text = re.sub(r"\\cdashline\{[^}]*\}\s*", "", text)
    # Remove \cdashlinelr{n-m}
    text = re.sub(r"\\cdashlinelr\{[^}]*\}\s*", "", text)
    # Remove \specialrule{w}{above}{below}
    text = re.sub(r"\\specialrule\{[^}]*\}\{[^}]*\}\{[^}]*\}\s*", "", text)
    # Remove \iffalse...\fi conditional blocks
    text = re.sub(r"\\iffalse\b.*?\\fi\b\s*", "", text, flags=re.DOTALL)
    # Remove rule width spec: [0.5pt], [1pt], etc. at the start
    text = re.sub(r"^\s*\[[\d.]+\w*\]\s*", "", text)
    # Remove \color{name} and \color[HTML]{code} (standalone color commands)
    text = re.sub(r"\\color(?:\[[^\]]*\])?\{[^}]*\}\s*", "", text)
    # Remove \colorXxx (no-brace form, e.g. \colorblue, \colorred)
    text = re.sub(r"\\color[a-zA-Z]+\s*", "", text)
    # Remove \arrayrulecolor{...} or \arrayrulecolorXxx (no-brace form)
    text = re.sub(r"\\arrayrulecolor(?:\{[^}]*\}|[a-zA-Z]+)\s*", "", text)
    # Remove \grow, \brow (custom row color commands)
    text = re.sub(r"\\[gb]row(?![a-zA-Z])\s*", "", text)
    # Remove \vrule with optional width spec
    text = re.sub(r"\\vrule\s*(?:width\s*\\[a-zA-Z]+)?\s*", "", text)
    # Remove \multirowcell{N}{content} → content
    text = re.sub(r"\\multirowcell\{[^}]*\}\{([^}]*)\}", r"\1", text)
    # \hbarthree{a}{b}{c} → a/b/c (custom horizontal bar command)
    text = re.sub(r"\\hbarthree\{([^}]*)\}\{([^}]*)\}\{([^}]*)\}", r"\1/\2/\3", text)
    # Remove custom commands
    # Custom ranking/color prefix commands: \firstcolor0.74 → 0.74, \grel 55.1 → 55.1
    text = re.sub(r"\\(firstcolor|secondcolor|gold|silve|bronze|grel|gbf)(?![a-zA-Z])\s*", "", text)
    # Short prefix commands: \fs53.60 → 53.60, \nd96.5 → 96.5, \ok 90.0 → 90.0
    text = re.sub(r"\\(fs|nd|rd|ok|no)(?![a-zA-Z])\s*", "", text)
    # \up8.09 → ↑8.09, \down0.45 → ↓0.45
    text = re.sub(r"\\up(?![a-zA-Z])\s*", "↑", text)
    text = re.sub(r"\\down(?![a-zA-Z])\s*", "↓", text)
    # Logo commands → text labels (these are content, not decoration)
    text = re.sub(r"\\languagelogos?\b\s*", "[Lang]", text)
    text = re.sub(r"\\imagelogo\b\s*", "[Img]", text)
    text = re.sub(r"\\videologo\b\s*", "[Vid]", text)
    text = re.sub(r"\\(pixtralemoji|llamaemoji|Locating|Gray|gray|icono|doct|codet|chordalone|chordalthree|chordalfour|chordaltwo|lgin|hgreen|myred|hlgood|hlbad|freeze|update|best|second|third|UR|icoyes|icohalf|Lgray|first|cycle|zzb|zza|Delta|usym|locogpt|Qwenemoji|Googleemoji|glmemoji|Claudeemoji|Openaiemoji|mathct|blue|red|green|greyc|gres|grem|grexl|gret|SSR|SR|champmark|champ|Large|lv|name|codeio|codeiopp|model|Ours|thedit|thevae|negCS|championlogo|silverlogo|bronzelogo|refinv|refeq|lvert|rvert|circ)\b[a-z]*\s*", "", text)
    # Remove \raisebox{...}{content} or \raisebox[...]{...}{content}
    text = re.sub(r"\\raisebox\{[^}]*\}(?:\[[^\]]*\])*\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\raisebox[\d.]+\w*(?:\[[^\]]*\])*\s*", "", text)
    # Remove \parbox{width}{content} or \parbox[align]{width}{content}
    text = re.sub(r"\\parbox(?:\[[^\]]*\])?\{[^}]*\}\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\parbox[\d.]+\w*\s*", "", text)
    # Remove \begintabular[...]{...} (handle nested braces in column spec)
    text = re.sub(rf"\\begintabular(?:\[[^\]]*\])?\{{({_NESTED})\}}\s*", "", text)
    # Remove \hspace{...} or \hspace*
    text = re.sub(r"\\hspace\*?\{[^}]*\}\s*", "", text)
    text = re.sub(r"\\hspace[\d.]+\w*\s*", "", text)
    # Remove \includegraphics[...]{...}
    text = re.sub(r"\\includegraphics(?:\[[^\]]*\])?\{[^}]*\}", "", text)
    # Remove \footnotemark, \footnotetext
    text = re.sub(r"\\footnotemark(?:\[[^\]]*\])?", "", text)
    text = re.sub(r"\\footnotetext\{[^}]*\}", "", text)
    # Unescape special chars
    for esc, ch in [("\\&", "&"), ("\\%", "%"), ("\\$", "$"),
                     ("\\#", "#"), ("\\_", "_"), ("\\{", "{"), ("\\}", "}")]:
        text = text.replace(esc, ch)
    # Strip leftover formatting commands (e.g. from typos like .\underline{050})
    text = re.sub(r"\\(?:textbf|textit|underline|emph)\{([^}]*)\}", r"\1", text)
    # Remove stray braces
    text = text.replace("{", "").replace("}", "")
    # Fallback: brace-stripped \begintabular[c]@c@content → content
    text = re.sub(r"\\begintabular(?:\[[^\]]*\])?[@clrp|]+\s*", "", text)
    # Fallback: brace-stripped \multicolumn1c|content → content
    text = re.sub(r"\\multicolumn\d+[clrp|]+\s*", "", text)
    # Fallback: brace-stripped \multirow2*content or \multirow-2*[-2pt]content → content
    text = re.sub(r"\\multirow-?[\d.]+\*(?:\[[^\]]*\])?\s*", "", text)
    # Fallback: \textcolorblueX → X (brace-stripped \textcolor{blue}{X})
    text = re.sub(r"\\textcolor(?:blue|red|green|black|gray|grey|cyan|magenta|yellow|orange|purple|brown|violet|pink|white)\s*", "", text)
    # Fallback: \textit(...) without braces → strip \textit
    text = re.sub(r"\\text(?:it|bf|rm|sc|tt|sf|sl)(?![a-zA-Z])", "", text)
    # Remove \iffalse (standalone, content after it is commented out)
    text = re.sub(r"\\iffalse\b.*", "", text)
    # Remove \fi (standalone)
    text = re.sub(r"\\fi\b\s*", "", text)
    # Remove \ref cross-references: \refinv:xxx, \refeq:xxx
    text = re.sub(r"\\ref[a-z]*:[^\s,)]+", "", text)
    # Remove \mathcal without braces (brace-stripped): \mathcalP → P
    text = re.sub(r"\\math(?:bf|cal|it|rm|tt|sf)(?=[A-Z0-9])", "", text)
    # Remove trailing backslash
    text = re.sub(r"\\$", "", text)
    # Normalize spacing around ±: "0.626 ±0.018" → "0.626±0.018"
    text = re.sub(r"\s*±\s*", "±", text)
    # Collapse multiple spaces
    text = re.sub(r"  +", " ", text)
    return text.strip()


def _try_parse_number(text: str) -> Optional[float]:
    """Try to parse text as a number. Returns float or None."""
    t = text.strip()
    if not t:
        return None
    # Handle .451 style (no leading zero)
    try:
        return float(t)
    except ValueError:
        return None


def _expand_row(cells: List[Cell], num_cols: int) -> List[Cell]:
    """Expand a row to full width, adding placeholders for merged cells."""
    result = []
    for cell in cells:
        result.append(cell)
        # Add horizontal placeholders
        for _ in range(cell.colspan - 1):
            result.append(Cell(value="", style=CellStyle()))
    # Pad to num_cols
    while len(result) < num_cols:
        result.append(Cell(value="", style=CellStyle()))
    return result[:num_cols]


def _detect_header_rows(rows: List[List[Cell]], num_cols: int) -> int:
    """Detect number of header rows by looking for multirow in first row."""
    if not rows:
        return 1
    max_rowspan = max((c.rowspan for c in rows[0]), default=1)
    return max(max_rowspan, 1)
