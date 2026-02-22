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
        # Remove rule commands within a row chunk
        s = re.sub(r"\\(toprule|bottomrule|midrule)\s*", "", original)
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
    rf"\\multicolumn\{{(\d+)\}}\{{([^}}]*)\}}\s*\{{({_NESTED})\}}"
)
_MULTIROW_RE = re.compile(
    rf"\\multirow\{{(\d+)\}}\{{[^}}]*\}}\s*\{{({_NESTED})\}}"
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
    """Split string by & respecting LaTeX brace nesting."""
    parts = []
    depth = 0
    current = []
    for ch in s:
        if ch == "{":
            depth += 1
            current.append(ch)
        elif ch == "}":
            depth -= 1
            current.append(ch)
        elif ch == "&" and depth == 0:
            parts.append("".join(current))
            current = []
        else:
            current.append(ch)
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

    colspan = 1
    rowspan = 1
    alignment = "center"

    # Extract multicolumn
    m = _MULTICOLUMN_RE.match(text)
    if m:
        colspan = int(m.group(1))
        alignment = m.group(2).strip()
        text = m.group(3).strip()

    # Extract multirow
    m = _MULTIROW_RE.match(text)
    if m:
        rowspan = int(m.group(1))
        text = m.group(2).strip()

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
    # \makecell{Things-\\EEG} → Things-EEG (remove line breaks, strip command)
    text = re.sub(r"\\makecell\{([^}]*)\}", lambda m: m.group(1).replace("\\\\", ""), text)
    # \var{$\pm$.005} → $\pm$.005 (strip wrapper, keep content)
    text = re.sub(r"\\var\{([^}]*)\}", r"\1", text)
    # Strip math mode: $D_\text{stage 1}$ → D_stage 1
    text = re.sub(r"\$([^$]*)\$", lambda m: m.group(1).replace("\\text{", "").replace("}", ""), text)
    # $\pm$ → ± (handle ± outside math mode too)
    text = text.replace("\\pm", "±")
    text = text.replace("$\\times$", "×")
    text = text.replace("\\textemdash", "—")
    # Remove citations: \citep{...}, \cite{...}, \citet{...}, etc.
    text = re.sub(r"~?\\cite[pt]?\{[^}]*\}", "", text)
    # Remove \, and other spacing
    text = re.sub(r"\\[,;!]", "", text)
    # Unescape special chars
    for esc, ch in [("\\&", "&"), ("\\%", "%"), ("\\$", "$"),
                     ("\\#", "#"), ("\\_", "_")]:
        text = text.replace(esc, ch)
    # Strip leftover formatting commands (e.g. from typos like .\underline{050})
    text = re.sub(r"\\(?:textbf|textit|underline|emph)\{([^}]*)\}", r"\1", text)
    # Remove stray braces
    text = text.replace("{", "").replace("}", "")
    # Normalize spacing around ±: "0.626 ±0.018" → "0.626±0.018"
    text = re.sub(r"\s*±\s*", "±", text)
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
