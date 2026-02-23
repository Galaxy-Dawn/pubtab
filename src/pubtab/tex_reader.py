"""LaTeX table parser — converts .tex to TableData."""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

from .models import Cell, CellStyle, TableData
from .utils import _latex_color_to_hex

# Module-level custom color registry (populated by read_tex per call)
_custom_colors: Dict[str, str] = {}


def _normalize_color(raw: str, spec: str = "") -> str | None:
    """Normalize a LaTeX color to '#RRGGBB' hex.

    Args:
        raw: Color value (name, hex, 'gray!20', '229,229,229').
        spec: Optional color model spec like 'RGB' or 'HTML'.
    """
    spec_upper = spec.strip("[]").upper()
    spec_raw = spec.strip("[]")
    if spec_upper == "RGB":
        parts = raw.split(",")
        if len(parts) == 3:
            try:
                if spec_raw == "rgb":
                    # Float 0.0-1.0
                    r, g, b = (min(255, round(float(p) * 255)) for p in parts)
                else:
                    # Integer 0-255
                    r, g, b = int(parts[0]), int(parts[1]), int(parts[2])
                return f"#{r:02X}{g:02X}{b:02X}"
            except ValueError:
                pass
    if spec_upper == "GRAY":
        try:
            v = min(255, round(float(raw.strip()) * 255))
            return f"#{v:02X}{v:02X}{v:02X}"
        except ValueError:
            pass
    if spec == "HTML":
        raw = raw.strip()
        if len(raw) == 6:
            return f"#{raw.upper()}"
    return _custom_colors.get(raw.strip()) or _latex_color_to_hex(raw)


def _trim_trailing_empty_cols(rows: List[List[Cell]], num_cols: int) -> int:
    """Remove trailing columns that are empty in all rows.

    A column is "used" if a non-empty cell starts there, or a non-full-width
    colspan from an earlier column covers it.
    """
    while num_cols > 1:
        col = num_cols - 1
        col_used = False
        for row in rows:
            pos = 0
            for cell in row:
                end = pos + cell.colspan
                if pos <= col < end and cell.value not in ("", None):
                    # Full-width span doesn't count
                    if cell.colspan < num_cols:
                        col_used = True
                        break
                pos = end
                if pos > col:
                    break
            if col_used:
                break
        if col_used:
            break
        for row in rows:
            if len(row) > col:
                row.pop(col)
            for i, cell in enumerate(row):
                if cell.colspan >= num_cols:
                    row[i] = Cell(value=cell.value, style=cell.style,
                                  rowspan=cell.rowspan, colspan=num_cols - 1,
                                  rich_segments=cell.rich_segments)
        num_cols -= 1
    return num_cols


def _merge_visual_multirow(rows: List[List[Cell]], header_rows: int, hline_before: List[bool]) -> None:
    """Merge visual multirow patterns in first column.

    For each non-empty cell in column A (after header), expands it to cover
    adjacent empty cells with the same bg_color. Stops at \\hline boundaries.

    This handles LaTeX tables that use visual positioning (placing label in
    middle row) instead of explicit \\multirow commands.
    """
    if len(rows) <= header_rows + 1:
        return

    num_rows = len(rows)

    # Find non-empty cells in column A after header
    for r in range(header_rows, num_rows):
        cell = rows[r][0]
        if not cell.value or cell.rowspan > 1:
            continue

        bg_color = cell.style.bg_color
        if not bg_color:
            continue

        # Expand upward (stop at header or hline)
        top = r
        while top > header_rows:
            # Stop if there's an hline before the row above
            if hline_before[top]:
                break
            above = rows[top - 1][0]
            if above.value or above.style.bg_color != bg_color:
                break
            top -= 1

        # Expand downward (stop at hline)
        bottom = r
        while bottom < num_rows - 1:
            below = rows[bottom + 1][0]
            if below.value or below.style.bg_color != bg_color:
                break
            # Stop if there's an hline before the next row
            if bottom + 1 < len(hline_before) and hline_before[bottom + 1]:
                break
            bottom += 1

        # If we found rows to merge
        span = bottom - top + 1
        if span > 1:
            # Move label to top row
            rows[top][0] = Cell(
                value=cell.value,
                style=cell.style,
                rowspan=span,
                colspan=cell.colspan,
            )
            # Clear the original position if different from top
            if top != r:
                rows[r][0] = Cell(value="", style=CellStyle())
            # Clear other rows in the span
            for clear_r in range(top + 1, bottom + 1):
                rows[clear_r][0] = Cell(value="", style=CellStyle())


def read_tex(
    tex: str,
    newcommands: Optional[Dict[str, Tuple[int, str]]] = None,
    definecolors: Optional[Dict[str, str]] = None,
    rowcolors: Optional[Tuple[int, str, str]] = None,
) -> TableData:
    """Parse a LaTeX table string into TableData.

    Handles: multicolumn, multirow, textbf, textit, underline, diagbox.
    """
    global _custom_colors
    # Populate custom color registry for this parse
    _custom_colors = _parse_definecolors(_strip_comments(tex))
    if definecolors:
        _custom_colors.update(definecolors)
    # Expand \newcommand definitions
    cmds = _parse_newcommands(_strip_comments(tex))
    if newcommands:
        cmds.update(newcommands)
    if cmds:
        tex = _expand_newcommands(tex, cmds)
    # Extract tabular body
    body = _extract_tabular_body(tex)
    if body is None:
        raise ValueError("No tabular environment found")

    # Strip \iffalse...\fi blocks before row splitting (they can span multiple rows)
    body = re.sub(r"\\iffalse\b.*?\\fi\b\s*", "", body, flags=re.DOTALL)

    # Replace nested \begin{tabular}...\end{tabular} with \makecell{...}
    # so the \\ inside doesn't get treated as a row separator
    body = re.sub(
        rf"\\begin\{{tabular\}}(?:\[[^\]]*\])?\{{({_NESTED})\}}(.*?)\\end\{{tabular\}}",
        lambda m: r"\makecell{" + m.group(2).strip() + "}",
        body, flags=re.DOTALL,
    )

    # Split into rows (by \\), filter out rules, track hline positions
    raw_rows_with_hline = _split_rows_with_hline(body)
    raw_rows = [r for r, _ in raw_rows_with_hline]
    hline_before = [h for _, h in raw_rows_with_hline]

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
    parsed_cell_counts = []  # original cell count per row (before expansion)
    for row in parsed_rows:
        parsed_cell_counts.append(len(row))
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

    # Apply \rowcolors alternating background to rows without explicit bg_color
    if rowcolors:
        rc_start, rc_color1, rc_color2 = rowcolors
        hex1 = _normalize_color(rc_color1)
        hex2 = _normalize_color(rc_color2)
        for r, row in enumerate(expanded):
            row_num = r + 1  # 1-based
            if row_num < rc_start:
                continue
            bg = hex1 if (row_num - rc_start) % 2 == 0 else hex2
            if not bg:
                continue
            for c, cell in enumerate(row):
                if not cell.style.bg_color:
                    expanded[r][c] = Cell(
                        value=cell.value,
                        style=CellStyle(
                            bold=cell.style.bold, italic=cell.style.italic,
                            underline=cell.style.underline, color=cell.style.color,
                            bg_color=bg, alignment=cell.style.alignment,
                            fmt=cell.style.fmt, diagbox=cell.style.diagbox,
                            rotation=cell.style.rotation,
                        ),
                        rowspan=cell.rowspan, colspan=cell.colspan,
                    )

    # Detect header rows (rows before first data-like row)
    header_rows = _detect_header_rows(expanded, num_cols)

    # Auto-merge empty header cells with content cells below
    if header_rows > 1 and len(expanded) >= header_rows:
        # Track positions covered by multicolumn/multirow in row 0
        covered = set()
        for c, cell in enumerate(expanded[0]):
            if cell.colspan > 1:
                for dc in range(1, cell.colspan):
                    if c + dc < num_cols:
                        covered.add(c + dc)
        for c in range(num_cols):
            if c in covered:
                continue
            top = expanded[0][c]
            if top.value in ("", None) and top.colspan == 1 and top.rowspan <= 1:
                below = expanded[1][c]
                if below.value not in ("", None):
                    # Move content up with rowspan
                    expanded[0][c] = Cell(
                        value=below.value, style=below.style,
                        rowspan=header_rows, colspan=below.colspan,
                    )
                    expanded[1][c] = Cell(value="", style=CellStyle())
                elif below.value in ("", None) and below.colspan == 1:
                    # Both empty — merge for clean header
                    expanded[0][c] = Cell(
                        value="", style=top.style,
                        rowspan=header_rows, colspan=1,
                    )

    # Expand section header rows (string content in first cell, rest empty) to full width
    # Only if the row was parsed as a single cell (no & separators) — rows with multiple
    # parsed cells are data rows with empty trailing values, not section headers.
    for i, row in enumerate(expanded):
        c0 = row[0]
        if (parsed_cell_counts[i] == 1
                and isinstance(c0.value, str) and c0.value
                and c0.colspan < num_cols
                and all(not c.value and c.value != 0 for c in row[1:])):
            row[0] = Cell(value=c0.value, style=c0.style,
                          rowspan=c0.rowspan, colspan=num_cols)

    # Merge visual multirow patterns in first column (e.g., topology labels spanning multiple rows)
    # Detect groups of consecutive rows with same bg_color in column A, with one label in the group
    _merge_visual_multirow(expanded, header_rows, hline_before)

    # Trim trailing empty columns (e.g., colspec has 14 cols but data only uses 7)
    num_cols = _trim_trailing_empty_cols(expanded, num_cols)

    return TableData(
        cells=expanded,
        num_rows=len(expanded),
        num_cols=num_cols,
        header_rows=header_rows,
    )


def _strip_comments(tex: str) -> str:
    """Remove LaTeX % comments (but not escaped \\%)."""
    return re.sub(r"(?<!\\)%[^\n]*", "", tex)


def _parse_definecolors(tex: str) -> Dict[str, str]:
    """Parse \\definecolor definitions from tex, return {name: '#RRGGBB'}."""
    colors: Dict[str, str] = {}
    for m in re.finditer(r"\\definecolor\{([^}]+)\}\{([^}]+)\}\{([^}]+)\}", tex):
        name, model, spec = m.group(1).strip(), m.group(2).strip(), m.group(3).strip()
        hex_color = _normalize_color(spec, f"[{model}]")
        if hex_color:
            colors[name] = hex_color
    return colors


def parse_definecolors(tex: str) -> Dict[str, str]:
    """Extract \\definecolor definitions from a full .tex file (public API)."""
    return _parse_definecolors(_strip_comments(tex))


def parse_definecolors_before(tex: str, pos: int) -> Dict[str, str]:
    """Extract \\definecolor definitions from tex up to position pos."""
    return _parse_definecolors(_strip_comments(tex[:pos]))


def _parse_rowcolors(tex: str) -> Optional[Tuple[int, str, str]]:
    """Parse last \\rowcolors command, return (start_row, color1, color2) or None."""
    result = None
    for m in re.finditer(r"\\rowcolors\{(\d+)\}\{([^}]*)\}\{([^}]*)\}", tex):
        result = (int(m.group(1)), m.group(2).strip(), m.group(3).strip())
    return result


def parse_rowcolors_before(tex: str, pos: int) -> Optional[Tuple[int, str, str]]:
    """Parse last \\rowcolors before position pos."""
    return _parse_rowcolors(_strip_comments(tex[:pos]))


def _parse_newcommands(tex: str) -> Dict[str, Tuple[int, str]]:
    """Parse \\newcommand/\\renewcommand/\\providecommand definitions from tex."""
    cmds: Dict[str, Tuple[int, str]] = {}
    pattern = re.compile(
        r"\\(?:new|renew|provide)command\s*"
        r"(?:\{\\([a-zA-Z]+)\}|\\([a-zA-Z]+))"
        r"(?:\[(\d+)\])?"
        rf"\{{({_NESTED})\}}",
        re.DOTALL,
    )
    for m in pattern.finditer(tex):
        name = m.group(1) or m.group(2)
        nargs = int(m.group(3)) if m.group(3) else 0
        body = m.group(4)
        if r"\begin{" in body or r"\includegraphics" in body:
            continue
        cmds[name] = (nargs, body)
    return cmds


def _expand_newcommands(tex: str, cmds: Dict[str, Tuple[int, str]]) -> str:
    """Expand \\newcommand definitions in tex, up to 10 rounds."""
    if not cmds:
        return tex
    sorted_cmds = sorted(cmds.items(), key=lambda x: len(x[0]), reverse=True)
    for _ in range(10):
        changed = False
        for name, (nargs, body) in sorted_cmds:
            pat = r"\\" + re.escape(name) + r"(?![a-zA-Z])"
            if nargs == 0:
                new_tex = re.sub(pat, lambda m, b=body: b, tex)
            else:
                for _i in range(nargs):
                    pat += rf"\s*\{{({_NESTED})\}}"

                def _repl(m, b=body, n=nargs):
                    r = b
                    for i in range(n):
                        r = r.replace(f"#{i+1}", m.group(i + 1))
                    return r

                new_tex = re.sub(pat, _repl, tex, flags=re.DOTALL)
            if new_tex != tex:
                changed = True
                tex = new_tex
        if not changed:
            break
    return tex


def parse_newcommands(tex: str) -> Dict[str, Tuple[int, str]]:
    """Extract \\newcommand definitions from a full .tex file (public API)."""
    return _parse_newcommands(_strip_comments(tex))


def _extract_tabular_body(tex: str) -> Optional[str]:
    """Extract content between \\begin{tabular} and \\end{tabular}."""
    tex = _strip_comments(tex)
    # Find outermost \begin{tabular}{colspec}...\end{tabular}
    m = re.search(
        rf"\\begin\{{tabular\}}(?:\[[^\]]*\])?\{{({_NESTED})\}}",
        tex, re.DOTALL,
    )
    if not m:
        return None
    start = m.end()
    # Walk forward counting nested \begin{tabular}/\end{tabular}
    depth = 1
    pos = start
    begin_tag = "\\begin{tabular}"
    end_tag = "\\end{tabular}"
    while pos < len(tex) and depth > 0:
        bi = tex.find(begin_tag, pos)
        ei = tex.find(end_tag, pos)
        if ei == -1:
            break
        if bi != -1 and bi < ei:
            depth += 1
            pos = bi + len(begin_tag)
        else:
            depth -= 1
            if depth == 0:
                return tex[start:ei].strip()
            pos = ei + len(end_tag)
    return None


def _split_rows(body: str) -> List[str]:
    """Split tabular body into row strings, skipping rule commands."""
    return [r for r, _ in _split_rows_with_hline(body)]


def _split_rows_with_hline(body: str) -> List[Tuple[str, bool]]:
    """Split tabular body into (row_string, has_hline_before) tuples.

    The has_hline_before flag indicates if a \\hline appeared before this row,
    which can be used to detect group boundaries in visual multirow patterns.
    """
    # Split by \\ respecting brace nesting (don't split inside {})
    parts = _split_by_double_backslash(body)
    rows = []
    for part in parts:
        original = part.strip()
        # Check for \hline at the start (indicates group boundary)
        has_hline = bool(re.match(r"^\s*\\hline\b", original))
        # Remove rule commands within a row chunk (with optional arguments like [0.5pt])
        s = re.sub(r"^\s*\\hline\s*", "", original)  # Remove leading \hline first
        s = re.sub(r"\\(toprule|bottomrule|midrule)(?:\[[^\]]*\])?\s*", "", s)
        s = re.sub(r"\\cmidrule(\([^)]*\))?\{[^}]*\}\s*", "", s)
        s = s.strip()
        if s:
            rows.append((s, has_hline))
        elif not original or has_hline:
            # Preserve empty rows (needed for multirow placeholders)
            # Also preserve hline-only rows as empty with hline flag
            rows.append(("", has_hline))
        # else: rule-only row (other than hline), skip
    # Trim leading/trailing empty rows (but preserve their hline flags for next non-empty)
    while rows and rows[-1][0] == "" and not rows[-1][1]:
        rows.pop()
    while rows and rows[0][0] == "" and not rows[0][1]:
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
    # Extract \rowcolor before splitting — it applies to the entire row
    row_bg = None
    m = re.search(r"\\rowcolor(\[[^\]]*\])?\{([^}]*)\}", row_str)
    if m:
        row_bg = _normalize_color(m.group(2), m.group(1) or "")
        row_str = (row_str[:m.start()] + row_str[m.end():]).strip()

    # Extract custom row color commands: \grow{...} or \grow ... → gray, \brow{...} or \brow ... → lightblue
    # Handle both \cmd{content} and \cmd content (LaTeX first-token argument)
    m = re.search(r"\\grow(?:\{([^}]*)\}|(?!\{)\s*(\S+))", row_str)
    if m:
        row_bg = "#F0F0F0"  # gray94
        content = m.group(1) or m.group(2) or ""
        row_str = (row_str[:m.start()] + content + row_str[m.end():]).strip()
    m = re.search(r"\\brow(?:\{([^}]*)\}|(?!\{)\s*(\S+))", row_str)
    if m:
        row_bg = "#DDEBF7"  # lightblue RGB(221,235,247)
        content = m.group(1) or m.group(2) or ""
        row_str = (row_str[:m.start()] + content + row_str[m.end():]).strip()

    parts = _split_by_ampersand(row_str)
    cells = []
    for part in parts:
        cell = _parse_cell(part.strip())
        if row_bg and not cell.style.bg_color:
            cell = Cell(
                value=cell.value, style=CellStyle(
                    bold=cell.style.bold, italic=cell.style.italic,
                    underline=cell.style.underline, color=cell.style.color,
                    bg_color=row_bg, alignment=cell.style.alignment,
                    fmt=cell.style.fmt, diagbox=cell.style.diagbox,
                    rotation=cell.style.rotation,
                ),
                rowspan=cell.rowspan, colspan=cell.colspan,
            )
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
        m = re.search(r"\\cellcolor(\[[^\]]*\])?\{([^}]*)\}", text)
        if m:
            bg_color = _normalize_color(m.group(2), m.group(1) or "")
            text = (text[:m.start()] + text[m.end():]).strip()
            continue
        m = re.search(r"\\rowcolor(\[[^\]]*\])?\{([^}]*)\}", text)
        if m:
            bg_color = _normalize_color(m.group(2), m.group(1) or "")
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

    # Convert mid-text \color{name}{content} → \textcolor{name}{content}
    # so the rich_segments logic below can detect it
    text = re.sub(
        rf"\\color(\[[^\]]*\])?\{{({_NESTED})\}}\{{({_NESTED})\}}",
        lambda m: f"\\textcolor{m.group(1) or ''}{{{m.group(2)}}}{{{m.group(3)}}}",
        text,
    )

    # Detect multiple \textcolor → rich text segments
    _tc_matches = list(re.finditer(rf"\\textcolor(?:\[[^\]]*\])?\{{({_NESTED})\}}\{{({_NESTED})\}}", text))
    rich_segments = None
    if len(_tc_matches) >= 1:
        segs = []
        pos = 0
        for _m in _tc_matches:
            if _m.start() > pos:
                plain = _clean_latex(text[pos:_m.start()]).strip()
                if plain:
                    segs.append((plain, None, False, False, False))
            color = _normalize_color(_m.group(1))
            sb, si, su, _, content = _parse_formatting(_m.group(2))
            content = content.strip()
            if content:
                segs.append((content, color, sb, si, su))
            pos = _m.end()
        if pos < len(text):
            remaining = _clean_latex(text[pos:]).strip()
            if remaining:
                segs.append((remaining, None, False, False, False))
        if len(segs) > 1:
            rich_segments = tuple(segs)

    # Extract \textcolor{color}{content}
    for _ in range(10):
        m = re.search(rf"\\textcolor(\[[^\]]*\])?\{{({_NESTED})\}}\{{({_NESTED})\}}", text)
        if not m:
            break
        if not text_color:
            text_color = _normalize_color(m.group(2), m.group(1) or "")
        text = text[:m.start()] + m.group(3) + text[m.end():]

    # Extract standalone \color{name} (switch command, affects rest of group)
    # Only treat as cell-level color if at the start (no content before it)
    if not text_color:
        m = re.search(r"\\color(\[[^\]]*\])?\{([^}]*)\}", text)
        if m and not text[:m.start()].strip():
            text_color = _normalize_color(m.group(2), m.group(1) or "")
            text = (text[:m.start()] + text[m.end():]).strip()

    # Extract \fcolorbox{frame}{bg}{content} and \colorbox{color}{content}
    for _ in range(10):
        m = re.search(r"\\fcolorbox\{[^}]*\}\{([^}]*)\}\{([^}]*)\}", text)
        if m:
            if not bg_color:
                bg_color = _normalize_color(m.group(1))
            text = text[:m.start()] + m.group(2) + text[m.end():]
            continue
        m = re.search(r"\\colorbox(\[[^\]]*\])?\{([^}]*)\}\{([^}]*)\}", text)
        if m:
            if not bg_color:
                bg_color = _normalize_color(m.group(2), m.group(1) or "")
            text = text[:m.start()] + m.group(3) + text[m.end():]
            continue
        break

    # Extract \rotatebox{angle}{content} or \rotatebox[origin=c]{angle}{content}
    m = re.search(r"\\rotatebox(?:\[[^\]]*\])?\{(\d+)\}\{([^}]*)\}", text)
    if m:
        rotation = int(m.group(1))
        text = text[:m.start()] + m.group(2) + text[m.end():]

    # Strip redundant outer braces: {{\underline{.451}}} → \underline{.451}
    text = _strip_outer_braces(text)

    # Extract custom color formatting: \gbf{...} → green bold, \rbf{...} → red bold
    gbf_bold = False
    for _ in range(3):
        m = re.search(r"\\gbf\{([^}]*)\}", text)
        if m:
            text_color = "#228B22"  # darkgreen
            gbf_bold = True
            text = text[:m.start()] + m.group(1) + text[m.end():]
            continue
        m = re.search(r"\\rbf\{([^}]*)\}", text)
        if m:
            text_color = "#FF0000"  # red
            gbf_bold = True
            text = text[:m.start()] + m.group(1) + text[m.end():]
            continue
        m = re.search(rf"\\gray\{{({_NESTED})\}}", text)
        if m:
            if not text_color:
                text_color = "#808080"  # gray
            text = text[:m.start()] + m.group(1) + text[m.end():]
            continue
        break

    # Parse formatting and extract value
    bold, italic, underline, diagbox_parts, value = _parse_formatting(text)
    if gbf_bold:
        bold = True

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
        rich_segments=rich_segments,
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

        m = re.fullmatch(r"\\(?:underline|ul)\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}", t)
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
    # \begin{tabular}[c]{@{}c@{}}content\end{tabular} → content (nested tabular as makecell)
    text = re.sub(
        r"\\begin\{tabular\}(?:\[[^\]]*\])?\{[^}]*\}(.*?)\\end\{tabular\}",
        lambda m: m.group(1).replace("\\\\", "\n"),
        text, flags=re.DOTALL,
    )
    # \makebox[width][pos]{content} → content
    text = re.sub(r"\\makebox(?:\[[^\]]*\])*\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}", r"\1", text)
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
    text = text.replace("\\ddagger", "‡")
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
    text = text.replace("\\cdots", "⋯")
    text = re.sub(r"\\cdot(?![a-zA-Z])", "·", text)
    text = text.replace("\\ell", "ℓ")
    # Superscript: ^{content} → content, ^X → X
    text = re.sub(r"\^\\text\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\^\\dagger\b", "†", text)
    text = re.sub(r"\^\\ddagger\b", "‡", text)
    text = re.sub(r"\^\\ddag\b", "‡", text)
    text = re.sub(r"\^\\S\b", "§", text)
    text = re.sub(r"\^\\star\b", "★", text)
    text = re.sub(r"\^\\mathrm\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\^\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\^([a-zA-Z0-9*†‡§★]+)", r"\1", text)
    # Subscript: _{content} → _content, _X → _X (preserve underscore)
    text = re.sub(r"(?<!\\)_\{([^}]*)\}", r"_\1", text)
    text = re.sub(r"(?<!\\)_([a-zA-Z0-9])(?![a-zA-Z0-9])", r"_\1", text)
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
    text = text.replace("\\ding{56}", "✗")
    # \ding without braces (after brace stripping)
    text = re.sub(r"\\ding\s*5[12]", "✓", text)
    text = re.sub(r"\\ding\s*5[56]", "✗", text)
    text = text.replace("\\bigstar", "★")
    text = text.replace("\\blacktriangledown", "▼")
    text = text.replace("\\compareyes", "✓")
    text = text.replace("\\comparepartially", "∼")
    text = text.replace("\\compareno", "✗")
    text = text.replace("\\tick", "✓")
    text = text.replace("\\cross", "✗")
    # \usym{XXXX} → Unicode character
    text = text.replace("\\usym{2713}", "✓")
    text = text.replace("\\usym{2714}", "✓")
    text = text.replace("\\usym{2717}", "✗")
    text = re.sub(r"\\usym\{([0-9A-Fa-f]{4,5})\}", lambda m: chr(int(m.group(1), 16)), text)
    # \faIcon{...} → strip (FontAwesome icons)
    text = re.sub(r"\\faIcon\{[^}]*\}", "", text)
    # FontAwesome direct commands
    text = text.replace("\\faCheckCircle", "✓")
    text = text.replace("\\faTimesCircle", "✗")
    # \kern dimension → strip
    text = re.sub(r"\\kern\s*-?[\d.]+\w*\s*", "", text)
    # \rlap{content} → content
    text = re.sub(r"\\rlap\{([^}]*)\}", r"\1", text)
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
    text = re.sub(r"\\xspace\b\s*", "", text)
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
    text = re.sub(r"\\(pixtralemoji|llamaemoji|Locating|Gray|gray|icono|lgin|hgreen|myred|hlgood|hlbad|freeze|update|best|second|third|UR|icoyes|icohalf|Lgray|first|zzb|zza|Delta|usym|locogpt|Qwenemoji|Googleemoji|glmemoji|Claudeemoji|Openaiemoji|blue|red|green|greyc|gres|grem|grexl|gret|SSR|SR|champmark|champ|Large|lv|name|codeio|codeiopp|model|Ours|negCS|championlogo|silverlogo|bronzelogo|refinv|refeq|lvert|rvert|circ)\b[a-z]*\s*", "", text)
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
    # Generic fallback: convert remaining \command to plain text (preserve custom commands like \thedit)
    text = re.sub(r"\\([a-zA-Z]+)\s*", r"\1", text)
    # Remove trailing backslash
    text = re.sub(r"\\$", "", text)
    # Normalize spacing around ±: "0.626 ±0.018" → "0.626±0.018"
    text = re.sub(r"\s*±\s*", "±", text)
    # Collapse multiple spaces
    text = re.sub(r"  +", " ", text)
    return text.strip()


def _try_parse_number(text: str) -> Optional[float]:
    """Try to parse text as a number. Returns float or None.

    Note: Strings starting with '+' are NOT parsed as numbers to preserve the sign.
    """
    t = text.strip()
    if not t:
        return None
    # Preserve strings starting with '+' (e.g., '+1.12') as strings
    if t.startswith('+'):
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
    """Detect number of header rows.

    Checks: multirow spans, or multicolumn cells indicating a multi-row header.
    """
    if not rows:
        return 1
    max_rowspan = max((c.rowspan for c in rows[0]), default=1)
    if max_rowspan > 1:
        return max_rowspan
    # First row has partial multicolumn → likely a multi-row header
    if len(rows) > 1 and any(1 < c.colspan < num_cols for c in rows[0]):
        return 2
    return 1
