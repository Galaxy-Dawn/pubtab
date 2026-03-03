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
    # Handle color!percent mixing where base is a custom-defined color
    raw_stripped = raw.strip()
    if "!" in raw_stripped:
        _parts = raw_stripped.split("!")
        _base_name = _parts[0].strip()
        _base_hex = _custom_colors.get(_base_name)
        if _base_hex and len(_parts) >= 2:
            try:
                _pct = float(_parts[1]) / 100.0
                _h = _base_hex.lstrip("#")
                _br = int(_h[0:2], 16)
                _bg = int(_h[2:4], 16)
                _bb = int(_h[4:6], 16)
                _r = int(_br * _pct + 255 * (1 - _pct))
                _g = int(_bg * _pct + 255 * (1 - _pct))
                _b = int(_bb * _pct + 255 * (1 - _pct))
                return f"#{_r:02X}{_g:02X}{_b:02X}"
            except ValueError:
                pass
    return _custom_colors.get(raw_stripped) or _latex_color_to_hex(raw)


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


def read_tex_multi(
    tex: str,
    newcommands: Optional[Dict[str, Tuple[int, str]]] = None,
    definecolors: Optional[Dict[str, str]] = None,
) -> List[TableData]:
    """Parse all tabular environments from a LaTeX string.

    For multi-table files, returns one TableData per tabular environment.
    Custom colors and newcommands defined before each table are resolved.

    Returns:
        List of TableData, one per tabular environment found.
    """
    stripped = _strip_comments(tex)

    # Parse shared newcommands from the full file
    cmds = _parse_newcommands(stripped)
    if newcommands:
        cmds.update(newcommands)

    # Expand newcommands in the full tex once
    expanded_tex = _expand_newcommands(tex, cmds) if cmds else tex

    # Extract all complete tabular blocks with positions from expanded tex
    all_blocks = _extract_all_tabular_blocks(expanded_tex)
    if not all_blocks:
        raise ValueError("No tabular environment found")

    expanded_stripped = _strip_comments(expanded_tex)
    tables: List[TableData] = []
    for block, block_pos in all_blocks:
        # Resolve colors defined before this table
        block_colors = _parse_definecolors(expanded_stripped[:block_pos])
        if definecolors:
            block_colors.update(definecolors)
        # Resolve rowcolors before this table
        block_rowcolors = _parse_rowcolors(expanded_stripped[:block_pos])
        try:
            table = read_tex(
                block,
                newcommands={},  # already expanded
                definecolors=block_colors,
                rowcolors=block_rowcolors,
            )
            tables.append(table)
        except (ValueError, IndexError):
            continue  # skip unparseable tables

    if not tables:
        raise ValueError("No parseable tabular environments found")
    return tables


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
    """Parse \\newcommand/\\renewcommand/\\providecommand/\\def definitions from tex."""
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
    # Also parse \def\name{body} (no-argument form only)
    def_pattern = re.compile(rf"\\def\\([a-zA-Z]+)\s*\{{({_NESTED})\}}", re.DOTALL)
    for m in def_pattern.finditer(tex):
        name, body = m.group(1), m.group(2)
        if r"\begin{" in body or r"\includegraphics" in body:
            continue
        cmds.setdefault(name, (0, body))
    return cmds


def _expand_newcommands(tex: str, cmds: Dict[str, Tuple[int, str]]) -> str:
    """Expand \\newcommand definitions in tex, up to 10 rounds."""
    if not cmds:
        return tex
    sorted_cmds = sorted(cmds.items(), key=lambda x: len(x[0]), reverse=True)
    for _ in range(10):
        changed = False
        for name, (nargs, body) in sorted_cmds:
            pat = re.escape(name) + r"(?![a-zA-Z])"
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


def _strip_newcommand_defs(tex: str) -> str:
    """Remove \\newcommand/\\renewcommand/\\providecommand definitions.

    Prevents _extract_tabular_body from picking up \\begin{tabular} inside
    command definitions (e.g. \\specialcell).
    """
    return re.sub(
        r"\\(?:new|renew|provide)command\s*"
        r"(?:\{\\[a-zA-Z]+\}|\\[a-zA-Z]+)"
        r"(?:\[\d+\])?(?:\[[^\]]*\])?"
        rf"\{{({_NESTED})\}}",
        "",
        tex,
        flags=re.DOTALL,
    )


def _extract_tabular_body(tex: str) -> Optional[str]:
    """Extract content between \\begin{tabular} and \\end{tabular}.

    When a file contains multiple tabular environments (e.g. two side-by-side
    tables), returns the body of the *largest* one (by character length) so
    that the most-content-rich table is chosen rather than always the first.
    """
    tex = _strip_comments(tex)
    # Strip \newcommand definitions so we don't pick up \begin{tabular} inside
    # command bodies (e.g. \specialcell uses \begin{tabular} internally)
    tex = _strip_newcommand_defs(tex)

    begin_tag = "\\begin{tabular}"
    end_tag = "\\end{tabular}"
    bodies: List[str] = []

    search_from = 0
    while True:
        m = re.search(
            rf"\\begin\{{tabular\}}(?:\[[^\]]*\])?\{{({_NESTED})\}}",
            tex[search_from:], re.DOTALL,
        )
        if not m:
            break
        abs_start = search_from + m.end()
        # Walk forward counting nested \begin{tabular}/\end{tabular}
        depth = 1
        pos = abs_start
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
                    bodies.append(tex[abs_start:ei].strip())
                    search_from = ei + len(end_tag)
                    break
                pos = ei + len(end_tag)
        else:
            break

    if not bodies:
        return None
    # Return the largest body — for multi-table files this picks the main table
    return max(bodies, key=len)


def _extract_all_tabular_blocks(tex: str) -> List[Tuple[str, int]]:
    """Extract all complete tabular blocks with their start positions.

    Returns:
        List of (full_block, start_position) tuples. Each full_block includes
        the \\begin{tabular}{...}...\\end{tabular} tags so it can be passed
        directly to read_tex().
    """
    # Only strip comments, NOT newcommand defs, to preserve positions.
    # When called from read_tex_multi, the input is already expanded so
    # newcommand defs don't need to be stripped for finding tabular blocks.
    tex_clean = _strip_comments(tex)

    begin_tag = "\\begin{tabular}"
    end_tag = "\\end{tabular}"
    results: List[Tuple[str, int]] = []

    search_from = 0
    while True:
        m = re.search(
            rf"\\begin\{{tabular\}}(?:\[[^\]]*\])?\{{({_NESTED})\}}",
            tex_clean[search_from:], re.DOTALL,
        )
        if not m:
            break
        block_start = search_from + m.start()
        abs_start = search_from + m.end()
        depth = 1
        pos = abs_start
        while pos < len(tex_clean) and depth > 0:
            bi = tex_clean.find(begin_tag, pos)
            ei = tex_clean.find(end_tag, pos)
            if ei == -1:
                break
            if bi != -1 and bi < ei:
                depth += 1
                pos = bi + len(begin_tag)
            else:
                depth -= 1
                if depth == 0:
                    block_end = ei + len(end_tag)
                    results.append((tex_clean[block_start:block_end], block_start))
                    search_from = block_end
                    break
                pos = ei + len(end_tag)
        else:
            break

    return results


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
        # Remove \specialrule{width}{space_above}{space_below}
        s = re.sub(r"\\specialrule\{[^}]*\}\{[^}]*\}\{[^}]*\}\s*", "", s)
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


_NESTED = r"(?:[^{}]|\{(?:[^{}]|\{(?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*\})*\})*"
_MULTICOLUMN_RE = re.compile(
    rf"\\multicolumn\{{(\d+)\}}\{{({_NESTED})\}}\s*\{{({_NESTED})\}}"
)
_MULTIROW_RE = re.compile(
    rf"\\multirow(?:\[[^\]]*\])?\{{(-?[\d.]+)\}}(?:\[[^\]]*\])?(?:\{{[^}}]*\}}|\*)\s*\{{({_NESTED})\}}"
)
_MULTIROWCELL_RE = re.compile(
    rf"\\multirowcell\{{(-?[\d.]+)\}}(?:\[[^\]]*\])?\{{({_NESTED})\}}"
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
                rich_segments=cell.rich_segments,
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
        text = re.sub(r"^\\arrayrulecolor(?:\{[^}]*\}|[a-zA-Z]+)\s*", "", text)
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

    # Strip \parbox wrapper so inner \multirow/\rotatebox can be detected
    m = re.match(rf"\\parbox(?:\[[^\]]*\])?\{{[^}}]*\}}\{{({_NESTED})\}}", text)
    if m:
        text = m.group(1).strip()

    # Strip leading font commands that block .match() for multirow/multicolumn
    text = re.sub(r"^\\(bf|it|rm|tt|sc|bfseries|itshape|rmfamily|ttfamily|scshape|boldmath|bm)\b\s*", "", text).strip()

    # Extract multicolumn (handle nested: outer alignment wins)
    m = _MULTICOLUMN_RE.match(text)
    if m:
        colspan = int(m.group(1))
        alignment = m.group(2).strip()
        text = m.group(3).strip()
        # If inner content is also a \multicolumn, extract it (keep outer alignment)
        m2 = _MULTICOLUMN_RE.match(text)
        if m2:
            text = m2.group(3).strip()
        # Re-extract cellcolor from inside multicolumn content
        for _ in range(3):
            m2 = re.search(r"\\cellcolor(\[[^\]]*\])?\{([^}]*)\}", text)
            if m2:
                if not bg_color:
                    bg_color = _normalize_color(m2.group(2), m2.group(1) or "")
                text = (text[:m2.start()] + text[m2.end():]).strip()
                continue
            break

    # Extract multirowcell (makecell package: \multirowcell{N}{content})
    m = _MULTIROWCELL_RE.match(text)
    if m:
        rowspan = round(float(m.group(1)))
        text = m.group(2).strip()

    # Extract multirow
    m = _MULTIROW_RE.match(text)
    if m:
        rowspan = round(float(m.group(1)))
        text = m.group(2).strip()

    # Strip \makebox BEFORE color conversion so {\color{...} content} inside \makebox
    # doesn't split the text before \makebox can be removed (e.g. \compareyes expansion)
    text = re.sub(r"\\makebox(?:\[[^\]]*\])*\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}", r"\1", text)

    # Convert mid-text \color{name}{content} → \textcolor{name}{content}
    # so the rich_segments logic below can detect it
    text = re.sub(
        rf"\\color(\[[^\]]*\])?\{{({_NESTED})\}}\{{({_NESTED})\}}",
        lambda m: f"\\textcolor{m.group(1) or ''}{{{m.group(2)}}}{{{m.group(3)}}}",
        text,
    )
    # Convert brace-group color switch {\color{name} content} → \textcolor{name}{content}
    # e.g. {\color{Red} \Large $\downarrow$} → \textcolor{Red}{$\downarrow$}
    text = re.sub(
        rf"\{{\\color(\[[^\]]*\])?\{{([^}}]*)\}}\s*({_NESTED})\}}",
        lambda m: f"\\textcolor{m.group(1) or ''}{{{m.group(2)}}}{{{m.group(3).strip()}}}",
        text,
    )

    # Extract \fcolorbox{frame}{bg}{content} and \colorbox{color}{content}
    # Must happen BEFORE rich_segments detection so the color name doesn't leak
    # into plain-text prefixes when _clean_latex is called on partial text.
    for _ in range(10):
        m = re.search(r"\\fcolorbox\{[^}]*\}\{([^}]*)\}\{([^}]*)\}", text)
        if m:
            if not bg_color:
                bg_color = _normalize_color(m.group(1))
            text = text[:m.start()] + m.group(2) + text[m.end():]
            continue
        m = re.search(rf"\\colorbox(\[[^\]]*\])?\{{([^}}]*)\}}\{{({_NESTED})\}}", text)
        if m:
            if not bg_color:
                bg_color = _normalize_color(m.group(2), m.group(1) or "")
            text = text[:m.start()] + m.group(3) + text[m.end():]
            continue
        break

    # Extract \rotatebox{angle}{content} or \rotatebox[origin=c]{angle}{content}
    # Must happen BEFORE rich_segments detection so rotatebox wrapper doesn't leak
    # into plain-text segments (e.g. "rotatebox90■barrier" instead of "■ barrier").
    m = re.search(rf"\\rotatebox(?:\[[^\]]*\])?\{{(\d+)\}}\{{({_NESTED})\}}", text)
    if m:
        rotation = int(m.group(1))
        text = text[:m.start()] + m.group(2) + text[m.end():]

    # Detect multiple \textcolor → rich text segments
    _tc_matches = list(re.finditer(rf"\\textcolor(?:\[[^\]]*\])?\{{({_NESTED})\}}\{{({_NESTED})\}}", text))
    rich_segments = None
    if len(_tc_matches) >= 1:
        segs = []
        pos = 0
        for _m in _tc_matches:
            if _m.start() > pos:
                _pb, _pi, _pu, _, plain = _parse_formatting(text[pos:_m.start()].strip())
                plain = plain.lstrip()
                if plain:
                    segs.append((plain, None, _pb, _pi, _pu))
            color = _normalize_color(_m.group(1))
            sb, si, su, _, content = _parse_formatting(_m.group(2))
            content = content.strip()
            if content:
                segs.append((content, color, sb, si, su))
            pos = _m.end()
        if pos < len(text):
            _raw_rem = text[pos:]
            _rb, _ri, _ru, _, remaining = _parse_formatting(_raw_rem.strip())
            remaining = remaining.rstrip()
            # Preserve a single leading space between colored and plain segments
            if remaining and _raw_rem and _raw_rem[0] == ' ' and not remaining.startswith(' '):
                remaining = ' ' + remaining
            if remaining:
                segs.append((remaining, None, _rb, _ri, _ru))
        if len(segs) > 1:
            rich_segments = tuple(segs)

    # Extract \textcolor{color}{content}
    # BUT: if there's \textcolor inside a subscript/superscript with a STANDARD LaTeX color,
    # preserve as raw LaTeX. Custom colors (like "down") are not safe to preserve because
    # the renderer doesn't output \definecolor commands.
    _subscript_color_match = re.search(r'[_^]\{[^}]*\\textcolor(?:\[[^\]]*\])?\{([^}]+)\}', text)
    if _subscript_color_match:
        _color_name = _subscript_color_match.group(1)
        # Check if it's a standard LaTeX color (NOT from custom colors)
        # Standard colors are those in _LATEX_COLORS dictionary
        from .utils import _LATEX_COLORS
        _is_standard_color = _color_name in _LATEX_COLORS or _color_name.lower() in _LATEX_COLORS
        if _is_standard_color:
            # Standard color - extract color and clean the \textcolor wrapper from text
            if not text_color:
                text_color = _normalize_color(_color_name)
            # Remove the \textcolor wrapper but keep the content
            # Pattern: \textcolor{colorname}{content} -> content
            text = re.sub(
                r'\\textcolor(?:\[[^\]]*\])?\{' + re.escape(_color_name) + r'\}\{([^}]+)\}',
                r'\1',
                text
            )
            return Cell(
                value=text,
                style=CellStyle(raw_latex=True, color=text_color, bg_color=bg_color, alignment=alignment),
                rowspan=rowspan, colspan=colspan,
            )
        # Custom color - let it be extracted normally (color will be lost, but output will be valid)
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

    # Strip redundant outer braces: {{\underline{.451}}} → \underline{.451}
    text = _strip_outer_braces(text)

    # Pre-extract \textcolor from math mode cells BEFORE the math-mode branch
    # e.g. $72.95_{\textcolor{ForestGreen}{+38.82}}$ → extract ForestGreen color
    # This handles cases where \textcolor is inside math mode with subscripts
    _math_textcolor_match = re.search(
        r'\$[^$]*\\textcolor(?:\[[^\]]*\])?\{([^}]+)\}\{([^}]+)\}[^$]*\$',
        text
    )
    if _math_textcolor_match and not text_color:
        _color_name = _math_textcolor_match.group(1)
        _color_content = _math_textcolor_match.group(2)
        text_color = _normalize_color(_color_name)
        # If we extracted a color, also clean the text by removing the \textcolor wrapper
        if text_color:
            _full_match = re.search(
                r'\\textcolor(?:\[[^\]]*\])?\{' + re.escape(_color_name) + r'\}\{' + re.escape(_color_content) + r'\}',
                text
            )
            if _full_match:
                text = text[:_full_match.start()] + _color_content + text[_full_match.end():]

    # If entire cell is a single math expression $...$, try to simplify first
    # e.g. $\textbf{80.11}$ → bold=True, value=80.11
    # e.g. $20.2\pm 0.2$ → value="20.2±0.2"
    # e.g. ${f_{\mathcal{D}}^{l-1}}'=A_{l}V_{l}$ → kept intact as raw LaTeX
    if re.match(r'^\$[^$]+\$$', text):
        _inner = text[1:-1].strip()
        _fmt_bold = False
        _fmt_italic = False
        _fmt_underline = False
        for _ in range(3):
            _m = re.fullmatch(r'\\textbf\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}', _inner)
            if _m:
                _fmt_bold = True
                _inner = _m.group(1).strip()
                continue
            _m = re.fullmatch(r'\\(?:underline|ul)\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}', _inner)
            if _m:
                _fmt_underline = True
                _inner = _m.group(1).strip()
                continue
            _m = re.fullmatch(r'\\(?:textit|emph)\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}', _inner)
            if _m:
                _fmt_italic = True
                _inner = _m.group(1).strip()
                continue
            break
        # Check for color commands inside subscripts/superscripts BEFORE cleaning.
        # e.g. 37.54_{\textcolor{ForestGreen}{+3.80}} should extract color and keep as raw LaTeX.
        _subscript_color_match = re.search(r'[_^]\{[^}]*\\textcolor(?:\[[^\]]*\])?\{([^}]+)\}\{([^}]+)\}', _inner)
        if _subscript_color_match:
            _sub_color_name = _subscript_color_match.group(1)
            _sub_color_content = _subscript_color_match.group(2)
            _sub_color = _normalize_color(_sub_color_name)
            if _sub_color:
                # Extract the color and keep the rest as raw LaTeX
                # Reconstruct the value without the \textcolor wrapper but preserve subscript
                _new_inner = re.sub(
                    r'\\textcolor(?:\[[^\]]*\])?\{' + re.escape(_sub_color_name) + r'\}\{' + re.escape(_sub_color_content) + r'\}',
                    _sub_color_content,
                    _inner
                )
                text = '$' + _new_inner + '$'
                text_color = _sub_color
            else:
                return Cell(
                    value=text,
                    style=CellStyle(raw_latex=True, bg_color=bg_color, alignment=alignment),
                    rowspan=rowspan, colspan=colspan,
                )
        elif re.search(r'[_^]\{[^}]*\\color\b', _inner):
            return Cell(
                value=text,
                style=CellStyle(raw_latex=True, bg_color=bg_color, alignment=alignment),
                rowspan=rowspan, colspan=colspan,
            )
        # Check for \mathcal BEFORE cleaning - it gets converted to Unicode but should be preserved
        # e.g. $\mathcal{I} \rightarrow \mathcal{P}$ should be kept as raw LaTeX
        if re.search(r'\\mathcal\{', _inner):
            return Cell(
                value=text,
                style=CellStyle(raw_latex=True, bg_color=bg_color, alignment=alignment),
                rowspan=rowspan, colspan=colspan,
            )
        _cleaned = _clean_latex(_inner)
        # If cleaned result still has LaTeX commands, keep as raw LaTeX
        # (e.g. ${f_{\mathcal{D}}^{l-1}}'=...$)
        # Note: _/^ alone (e.g. A_p, F_1) are OK — renderer wraps them in $...$
        if re.search(r'\\[a-zA-Z]', _cleaned):
            return Cell(
                value=text,
                style=CellStyle(raw_latex=True, bg_color=bg_color, alignment=alignment),
                rowspan=rowspan, colspan=colspan,
            )
        _num = _try_parse_number(_cleaned)
        return Cell(
            value=_num if _num is not None else _cleaned,
            style=CellStyle(bold=_fmt_bold, italic=_fmt_italic, underline=_fmt_underline,
                            color=text_color, bg_color=bg_color, alignment=alignment),
            rowspan=rowspan, colspan=colspan,
        )

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

    # If text contains \mathcal, preserve as raw LaTeX (Unicode conversion would break rendering)
    if re.search(r'\\mathcal\{', text):
        # Strip \arrayrulecolor{...} which may be prepended from \midrule lines
        clean_text = re.sub(r'\\arrayrulecolor(?:\{[^}]*\}|[a-zA-Z]+)\s*', '', text).strip()
        return Cell(
            value=clean_text,
            style=CellStyle(raw_latex=True, bg_color=bg_color, alignment=alignment),
            rowspan=rowspan, colspan=colspan,
        )

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
    _SIZE_CMDS = r"(?:Large|large|LARGE|small|footnotesize|normalsize|tiny|huge|Huge|scriptsize|normalfont)"
    changed = True
    while changed:
        changed = False
        t = text.strip()

        # Strip leading LaTeX size commands (e.g. \Large \underline{...} → \underline{...})
        m = re.fullmatch(rf"\\{_SIZE_CMDS}\s+(.*)", t, re.DOTALL)
        if m:
            text = m.group(1).strip()
            changed = True
            continue

        # Old-style \bf / \it switch commands (e.g. \bf 94.00 → bold=True, value=94.00)
        m = re.fullmatch(r"\\bf(?![a-zA-Z])\s*(.*)", t, re.DOTALL)
        if m:
            bold = True
            text = m.group(1).strip()
            changed = True
            continue

        m = re.fullmatch(r"\\it(?![a-zA-Z])\s*(.*)", t, re.DOTALL)
        if m:
            italic = True
            text = m.group(1).strip()
            changed = True
            continue

        # \makecell where every \\ line is \textbf{}: hoist bold to cell level
        m = re.fullmatch(rf"\\makecell(?:\[[^\]]*\])?\{{({_NESTED})\}}", t, re.DOTALL)
        if m:
            inner = m.group(1)
            line_parts = _split_by_double_backslash(inner)
            _BF_LINE = re.compile(r"\\textbf\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}")
            if line_parts and all(_BF_LINE.fullmatch(lp.strip()) for lp in line_parts if lp.strip()):
                bold = True
                stripped = [
                    _BF_LINE.fullmatch(lp.strip()).group(1) if lp.strip() else lp
                    for lp in line_parts
                ]
                text = r"\makecell{" + r" \\ ".join(stripped) + "}"
                changed = True
                continue

        m = re.fullmatch(r"\\textbf\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}(.*)", t, re.DOTALL)
        if m:
            bold = True
            text = (m.group(1) + m.group(2)).strip()
            changed = True
            continue

        m = re.fullmatch(r"\\textit\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}(.*)", t, re.DOTALL)
        if m:
            italic = True
            text = (m.group(1) + m.group(2)).strip()
            changed = True
            continue

        m = re.fullmatch(r"\\emph\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}(.*)", t, re.DOTALL)
        if m:
            italic = True
            text = (m.group(1) + m.group(2)).strip()
            changed = True
            continue

        m = re.fullmatch(r"\\(?:underline|ul)\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}(.*)", t, re.DOTALL)
        if m:
            underline = True
            text = (m.group(1) + m.group(2)).strip()
            changed = True
            continue

    # Multi-line bold: if every non-empty line is \textbf{...}, hoist bold to cell level
    if "\n" in text:
        _BF_LINE = re.compile(r"\\textbf\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}")
        lines = text.split("\n")
        if all(_BF_LINE.fullmatch(l.strip()) for l in lines if l.strip()):
            bold = True
            text = "\n".join(
                _BF_LINE.fullmatch(l.strip()).group(1) if l.strip() else l
                for l in lines
            )

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
        rf"\\begin\{{tabular\}}(?:\[[^\]]*\])?\{{({_NESTED})\}}(.*?)\\end\{{tabular\}}",
        lambda m: m.group(2).replace("\\\\", "\n"),
        text, flags=re.DOTALL,
    )
    # \makebox[width][pos]{content} → content
    text = re.sub(r"\\makebox(?:\[[^\]]*\])*\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}", r"\1", text)
    # \makecell{Things-\\EEG} → Things-\nEEG (preserve line breaks as newline)
    text = re.sub(rf"\\makecell(?:\[[^\]]*\])?\{{({_NESTED})\}}", lambda m: m.group(1).replace("\\\\", "\n"), text)
    # \specialcell[t]{content} → content (like makecell)
    text = re.sub(rf"\\specialcell(?:\[[^\]]*\])?\{{({_NESTED})\}}", lambda m: m.group(1).replace("\\\\", "\n"), text)
    # \shortstack{content} → content (like makecell)
    text = re.sub(r"\\shortstack(?:\[[^\]]*\])?\{([^}]*)\}", lambda m: m.group(1).replace("\\\\", "\n"), text)
    # Join hyphenated line breaks: "Conven- \ntional" → "Conventional"
    text = re.sub(r"- *\n *", "", text)
    # \var{$\pm$.005} → $\pm$.005 (strip wrapper, keep content)
    text = re.sub(r"\\var\{([^}]*)\}", r"\1", text)
    # \textcolor{color}{content} → content (strip color, keep content)
    for _ in range(5):
        m = re.search(rf"\\textcolor(?:\[[^\]]*\])?\{{({_NESTED})\}}\{{({_NESTED})\}}", text)
        if not m:
            break
        text = text[:m.start()] + m.group(2) + text[m.end():]
    # Arrow conversions BEFORE math stripping so $\Rightarrow$ adjacent to text works
    # e.g. zh$\Rightarrow$en → zh⇒en (not zhRightarrowen)
    text = re.sub(r"\\uparrow(?![a-zA-Z])", "↑", text)
    text = re.sub(r"\\downarrow(?![a-zA-Z])", "↓", text)
    text = re.sub(r"\\Downarrow(?![a-zA-Z])", "⇓", text)
    text = re.sub(r"\\Uparrow(?![a-zA-Z])", "⇑", text)
    text = re.sub(r"\\rightarrow(?![a-zA-Z])", "→", text)
    text = re.sub(r"\\leftarrow(?![a-zA-Z])", "←", text)
    text = re.sub(r"\\Rightarrow(?![a-zA-Z])", "⇒", text)
    text = re.sub(r"\\Leftarrow(?![a-zA-Z])", "⇐", text)
    # Strip math mode: $D_\text{stage 1}$ → D_stage 1
    text = re.sub(r"(?<!\\)\$([^$]*)\$", lambda m: re.sub(r"\\text\{([^}]*)\}", r"\1", m.group(1)), text)
    # Math symbols
    text = re.sub(r"\\pm(?![a-zA-Z])", "±", text)
    text = text.replace("$\\times$", "×")
    text = re.sub(r"\\texttimes\b", "×", text)
    text = re.sub(r"\\times(?![a-zA-Z])", "×", text)
    text = re.sub(r"\\textemdash\b", "—", text)
    text = re.sub(r"\\alpha\b", "α", text)
    text = re.sub(r"\\beta\b", "β", text)
    text = re.sub(r"\\gamma\b", "γ", text)
    text = re.sub(r"\\mu\b", "μ", text)
    text = re.sub(r"\\sigma\b", "σ", text)
    text = re.sub(r"\\delta\b", "δ", text)
    text = re.sub(r"\\epsilon\b", "ε", text)
    text = re.sub(r"\\theta\b", "θ", text)
    text = re.sub(r"\\lambda\b", "λ", text)
    text = re.sub(r"\\pi\b", "π", text)
    text = re.sub(r"\\omega\b", "ω", text)
    text = re.sub(r"\\tau\b", "τ", text)
    text = re.sub(r"\\triangle\b", "△", text)
    text = re.sub(r"\\star\b", "★", text)
    text = re.sub(r"\\textdaggerdbl\b", "‡", text)
    text = re.sub(r"\\textdagger\b", "†", text)
    text = re.sub(r"\\dagger\b", "†", text)
    text = re.sub(r"\\ddagger\b", "‡", text)
    text = re.sub(r"\\ddag\b", "‡", text)
    text = re.sub(r"\\dag(?![a-zA-Z])", "†", text)
    text = re.sub(r"\\S(?![a-zA-Z])", "§", text)
    text = re.sub(r"\\Sigma\b", "Σ", text)
    text = re.sub(r"\\Omega(?![a-zA-Z])", "Ω", text)
    text = re.sub(r"\\sim\b", "∼", text)
    text = re.sub(r"\\infty\b", "∞", text)
    text = re.sub(r"\\leq\b", "≤", text)
    text = re.sub(r"\\geq\b", "≥", text)
    text = re.sub(r"\\neq\b", "≠", text)
    text = re.sub(r"\\approx\b", "≈", text)
    text = re.sub(r"\\cdots\b", "⋯", text)
    text = re.sub(r"\\cdot(?![a-zA-Z])", "·", text)
    text = re.sub(r"\\ell\b", "ℓ", text)
    # Superscript: annotation symbols drop ^, math superscripts preserve ^
    text = re.sub(r"\^\\text\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\^\\dagger\b", "†", text)
    text = re.sub(r"\^\\ddagger\b", "‡", text)
    text = re.sub(r"\^\\ddag\b", "‡", text)
    text = re.sub(r"\^\\S\b", "§", text)
    text = re.sub(r"\^\\star\b", "★", text)
    # Braced annotation symbols: ^{\dagger} → †, ^{*} → *
    text = re.sub(r"\^\{\\dagger\}", "†", text)
    text = re.sub(r"\^\{\\ddagger\}", "‡", text)
    text = re.sub(r"\^\{\\star\}", "★", text)
    text = re.sub(r"\^\{\\S\}", "§", text)
    text = re.sub(r"\^\{\*\}", "*", text)
    text = re.sub(r"\^\\mathrm\{([^}]*)\}", r"\1", text)
    # General superscript: preserve ^ for math readability
    text = re.sub(r"\^\{([^}])\}", r"^\1", text)  # ^{T} → ^T
    text = re.sub(r"\^\{([^}]+)\}", r"^(\1)", text)  # ^{l-1} → ^(l-1)
    # Subscript: _{content} → _content, _X → _X (preserve underscore)
    # Special case: _{\pm...} or _{±...} → ±... (drop underscore, treat as ± annotation)
    text = re.sub(r"_\{\\pm(?![a-zA-Z])([^}]*)\}", r"±\1", text)
    text = re.sub(r"_\{±([^}]*)\}", r"±\1", text)
    # Special case: _{↓...} or _{↑...} → ↓... (drop underscore, arrow is visual formatting)
    text = re.sub(r"_\{([↓↑⇓⇑][^}]*)\}", r"\1", text)
    text = re.sub(r"_([↓↑⇓⇑])", r"\1", text)
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
    text = text.replace("\\blacksquare", "■")
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
    # Remove \- (LaTeX soft hyphen / discretionary hyphen)
    text = text.replace("\\-", "")
    # Remove \\ inside cells (line break in makecell) — must come before '\ ' conversion
    text = text.replace("\\\\", " ")
    # Remove \, and other spacing; convert '\ ' (forced space) to space
    text = text.replace("\\ ", " ")
    text = re.sub(r"\\[,;!]", "", text)
    text = re.sub(r"\\xspace\b\s*", "", text)
    text = re.sub(r"\\q?quad\s*", " ", text)
    # Remove font style commands: \bf, \rm, \it, \tt, \sc, \bfseries, \boldmath, etc.
    text = re.sub(r"\\(bfseries|rmfamily|itshape|ttfamily|scshape|boldmath|unboldmath)\b\s*", "", text)
    text = re.sub(r"\\(bf|rm|it|tt|sc|bm)(?![a-zA-Z])\s*", "", text)
    # Remove \textsc{...}, \texttt{...}, \textrm{...}, \textit{...}, \textbf{...}
    text = re.sub(r"\\text(sc|tt|rm|it|bf|sf|sl)\{([^}]*)\}", r"\2", text)
    # Remove \text{content} (plain, no style suffix)
    text = re.sub(r"\\text\{([^}]*)\}", r"\1", text)
    # Remove \textsuperscript{...} (supports nested braces)
    text = re.sub(rf"\\textsuperscript\{{({_NESTED})\}}", r"\1", text)
    # \mathcal{X} → Unicode script letters (common in ML papers)
    _mathcal_map = {
        "A": "𝒜", "B": "ℬ", "C": "𝒞", "D": "𝒟", "E": "ℰ", "F": "ℱ",
        "G": "𝒢", "H": "ℋ", "I": "ℐ", "J": "𝒥", "K": "𝒦", "L": "ℒ",
        "M": "ℳ", "N": "𝒩", "O": "𝒪", "P": "𝒫", "Q": "𝒬", "R": "ℛ",
        "S": "𝒮", "T": "𝒯", "U": "𝒰", "V": "𝒱", "W": "𝒲", "X": "𝒳",
        "Y": "𝒴", "Z": "𝒵",
    }
    text = re.sub(r"\\mathcal\{([A-Z])\}", lambda m: _mathcal_map.get(m.group(1), m.group(1)), text)
    # Remove \mathbf{...}, \mathcal{...} (remaining multi-char or lowercase), etc.
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
    text = re.sub(r"\\(firstBest|secondBest|firstcolor|secondcolor|gold|silve|bronze|grel|gbf)(?![a-zA-Z])\s*", "", text)
    # Short prefix commands: \fs53.60 → 53.60, \nd96.5 → 96.5, \ok 90.0 → 90.0
    text = re.sub(r"\\(fs|nd|rd|ok|no)(?![a-zA-Z])\s*", "", text)
    # \up8.09 → ↑8.09, \down0.45 → ↓0.45
    text = re.sub(r"\\up(?![a-zA-Z])\s*", "↑", text)
    text = re.sub(r"\\down(?![a-zA-Z])\s*", "↓", text)
    # \redup{0.66} → ↑0.66, \reddown{0.66} → ↓0.66 (custom arrow commands)
    text = re.sub(r"\\redup\{([^}]*)\}", r"↑\1", text)
    text = re.sub(r"\\reddown\{([^}]*)\}", r"↓\1", text)
    # Logo commands → text labels (these are content, not decoration)
    text = re.sub(r"\\languagelogos?\b\s*", "[Lang]", text)
    text = re.sub(r"\\imagelogo\b\s*", "[Img]", text)
    text = re.sub(r"\\videologo\b\s*", "[Vid]", text)
    text = re.sub(r"\\(pixtralemoji|llamaemoji|Locating|Gray|gray|icono|lgin|hgreen|myred|hlgood|hlbad|freeze|update|best|second|third|UR|icoyes|icohalf|Lgray|first|zzb|zza|Delta|usym|locogpt|Qwenemoji|Googleemoji|glmemoji|Claudeemoji|Openaiemoji|blue|red|green|greyc|gres|grem|grexl|gret|SSR|SR|champmark|champ|Large|lv|name|codeio|codeiopp|model|Ours|negCS|championlogo|silverlogo|bronzelogo|refinv|refeq|lvert|rvert|circ)\b[a-z]*\s*", "", text)
    # Remove \raisebox{...}{content} or \raisebox[...]{...}{content}
    text = re.sub(r"\\raisebox\{[^}]*\}(?:\[[^\]]*\])*\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\raisebox[\d.]+\w*(?:\[[^\]]*\])*\s*", "", text)
    # Remove \parbox{width}{content} or \parbox[align]{width}{content}
    text = re.sub(rf"\\parbox(?:\[[^\]]*\])?\{{[^}}]*\}}\{{({_NESTED})\}}", r"\1", text)
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
    Decimal strings (containing '.') are kept as strings to preserve trailing zeros:
    "94.00" must not become 94.0, which would render as "94" after xlsx roundtrip.
    Only integer strings (e.g. "94", "-3") are converted to float so that openpyxl
    reads them back as int and str(int) reproduces the original exactly.
    """
    t = text.strip()
    if not t:
        return None
    # Preserve strings starting with '+' (e.g., '+1.12') as strings
    if t.startswith('+'):
        return None
    # Preserve decimal strings to avoid losing trailing zeros through xlsx roundtrip.
    # "94.00" → float → 94.0 → xlsx → int(94) → "94" loses precision.
    if '.' in t:
        return None
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
