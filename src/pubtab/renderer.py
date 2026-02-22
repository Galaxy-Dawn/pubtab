"""Jinja2 rendering engine — converts TableData to LaTeX."""

from __future__ import annotations

from dataclasses import fields, replace
from pathlib import Path
from typing import Optional, Union

from jinja2 import Environment

from .models import Cell, CellStyle, SpacingConfig, TableData, ThemeConfig
from .themes import load_theme
from .utils import format_number, hex_to_latex_color, latex_escape


def _cell_to_latex(cell: Cell) -> str:
    """Convert a Cell to its LaTeX string representation."""
    val = cell.value
    if val is None or val == "":
        text = ""
    elif cell.style.raw_latex:
        text = str(val)
    elif cell.style.fmt and isinstance(val, (int, float)):
        text = format_number(val, cell.style.fmt, cell.style.strip_leading_zero)
    else:
        text = latex_escape(str(val))

    # Multi-line cell: convert \n back to \makecell{...\\...}
    if "\n" in text:
        text = "\\makecell{" + text.replace("\n", "\\\\") + "}"

    # Diagbox: diagonal header cell (before styling so bold/italic wraps it)
    if cell.style.diagbox:
        parts = cell.style.diagbox
        text = f"\\diagbox{{{parts[0]}}}{{{parts[1]}}}"

    # Apply styling (skip if raw_latex — user controls formatting)
    if not cell.style.raw_latex:
        if cell.style.bold:
            text = f"\\textbf{{{text}}}"
        if cell.style.italic:
            text = f"\\textit{{{text}}}"
        if cell.style.underline:
            text = f"\\underline{{{text}}}"
        if cell.style.color:
            rgb = hex_to_latex_color(cell.style.color)
            text = f"\\textcolor[RGB]{{{rgb}}}{{{text}}}"
    if cell.rowspan > 1:
        text = f"\\multirow{{{cell.rowspan}}}{{*}}{{{text}}}"
    if cell.colspan > 1:
        align = cell.style.alignment[0] if cell.style.alignment else "c"
        text = f"\\multicolumn{{{cell.colspan}}}{{{align}}}{{{text}}}"

    # cellcolor must be OUTSIDE multirow/multicolumn to color the full cell
    if not cell.style.raw_latex and cell.style.bg_color:
        rgb = hex_to_latex_color(cell.style.bg_color)
        text = f"\\cellcolor[RGB]{{{rgb}}}{text}"

    return text


def _build_col_spec(table: TableData, theme_config: ThemeConfig) -> str:
    """Build the column specification string (e.g. 'cccc')."""
    specs = []
    for c in range(table.num_cols):
        # Use first row's alignment as column default
        if table.cells and c < len(table.cells[0]):
            cell = table.cells[0][c]
            a = cell.style.alignment[0] if cell.style.alignment else "c"
        else:
            a = "c"
        specs.append(a)
    return "".join(specs)


def _auto_cmidrule(header_row_cells: list[Cell], num_cols: int) -> Optional[str]:
    """Auto-generate cmidrule from multicolumn cells in a header row."""
    rules = []
    col = 1
    skip = 0
    for cell in header_row_cells:
        if skip > 0:
            skip -= 1
            continue
        span = cell.colspan if cell.colspan > 1 else 1
        if span > 1:
            skip = span - 1
            if span < num_cols and cell.rowspan <= 1:
                rules.append(f"\\cline{{{col}-{col + span - 1}}}")
        col += 1 if span == 1 else span
    return " ".join(rules) if rules else None


def _normalize_group_separators(gs):
    """Convert List[int] to Dict[int, str] if needed."""
    if isinstance(gs, list):
        return {idx: "\\midrule" for idx in gs}
    return gs or {}


def render(
    table: TableData,
    theme: str = "three_line",
    caption: Optional[str] = None,
    label: Optional[str] = None,
    position: str = "htbp",
    raw_caption: bool = False,
    spacing: Optional[SpacingConfig] = None,
    font_size: Optional[str] = None,
    resizebox: Optional[str] = None,
    col_spec: Optional[str] = None,
    header_sep: Optional[str] = None,
    wide: bool = False,
    span_columns: Optional[bool] = None,
) -> str:
    """Render TableData to a LaTeX string.

    Args:
        table: The table data to render.
        theme: Theme name.
        caption: Table caption (always passed as-is, no escaping).
        label: LaTeX label.
        position: Table float position.
        raw_caption: Deprecated, ignored. Caption is always raw.
        spacing: Override spacing config.
        span_columns: Use table* for two-column spanning (replaces wide).

    Returns:
        LaTeX table string.
    """
    # span_columns takes precedence over deprecated wide
    if span_columns is not None:
        wide = span_columns
    config, template_str = load_theme(theme)
    env = Environment(
        block_start_string="{%",
        block_end_string="%}",
        variable_start_string="{{",
        variable_end_string="}}",
        comment_start_string="{#",
        comment_end_string="#}",
        keep_trailing_newline=True,
    )
    tmpl = env.from_string(template_str)

    # Build vertical merge maps for negative multirow (bg_color + rowspan > 1)
    _vmerge_bg: dict[tuple[int, int], str] = {}
    _vmerge_neg: dict[tuple[int, int], tuple[int, str]] = {}  # last row -> (rowspan, styled_text)
    _vmerge_suppress: set[tuple[int, int]] = set()  # master cells to suppress content
    for ri, row in enumerate(table.cells):
        ci = 0
        for cell in row:
            if cell.rowspan > 1 and cell.style.bg_color:
                _vmerge_suppress.add((ri, ci))
                for dr in range(1, cell.rowspan):
                    _vmerge_bg[(ri + dr, ci)] = cell.style.bg_color
                # Get styled text without multirow/cellcolor for negative multirow
                plain = Cell(value=cell.value,
                             style=replace(cell.style, bg_color=None),
                             rowspan=1, colspan=cell.colspan)
                _vmerge_neg[(ri + cell.rowspan - 1, ci)] = (cell.rowspan, _cell_to_latex(plain))
            ci += 1

    # Convert cells to LaTeX strings, skipping horizontal merge placeholders
    all_rows = []
    has_p_cols = col_spec and "p{" in (col_spec or "")
    for ri, row in enumerate(table.cells):
        latex_row = []
        skip = 0
        ci = 0
        for cell in row:
            if skip > 0:
                skip -= 1
                ci += 1
                continue
            if cell.colspan > 1:
                skip = cell.colspan - 1
            if (ri, ci) in _vmerge_suppress:
                # Master cell with bg_color: emit only cellcolor (content goes to last row)
                rgb = hex_to_latex_color(cell.style.bg_color)
                s = f"\\cellcolor[RGB]{{{rgb}}}"
            elif (ri, ci) in _vmerge_neg:
                # Last placeholder row: negative multirow with content
                rowspan, styled = _vmerge_neg[(ri, ci)]
                rgb = hex_to_latex_color(_vmerge_bg.get((ri, ci), ""))
                s = f"\\cellcolor[RGB]{{{rgb}}}\\multirow{{-{rowspan}}}{{*}}{{{styled}}}"
            elif not cell.value and cell.value != 0 and (ri, ci) in _vmerge_bg:
                # Other placeholder rows: just cellcolor
                rgb = hex_to_latex_color(_vmerge_bg[(ri, ci)])
                s = f"\\cellcolor[RGB]{{{rgb}}}"
            else:
                s = _cell_to_latex(cell)
            if has_p_cols and cell.colspan == 1 and s and not s.startswith("\\multicolumn"):
                s = f"\\multicolumn{{1}}{{c}}{{{s}}}"
            latex_row.append(s)
            ci += 1
        all_rows.append(latex_row)

    raw_header_rows = all_rows[: table.header_rows]
    body_rows = all_rows[table.header_rows :]

    # Build header rows with interleaved separators
    final_header_sep = header_sep
    if isinstance(header_sep, list) and len(header_sep) >= len(raw_header_rows):
        header_rows: list[Union[list[str], str]] = []
        for i, row in enumerate(raw_header_rows):
            header_rows.append(row)
            if i < len(raw_header_rows) - 1:
                header_rows.append(header_sep[i])
        final_header_sep = header_sep[-1]
    elif header_sep is None and table.header_rows > 1:
        # Auto-generate cmidrule between header rows from merged cells
        header_rows = []
        for i, row in enumerate(raw_header_rows):
            header_rows.append(row)
            if i < len(raw_header_rows) - 1:
                rule = _auto_cmidrule(table.cells[i], table.num_cols)
                if rule:
                    header_rows.append(rule)
    else:
        header_rows = raw_header_rows
        if isinstance(header_sep, list):
            final_header_sep = header_sep[-1] if header_sep else None

    # Normalize group_separators: List[int] → Dict[int, str]
    gs = _normalize_group_separators(table.group_separators)

    # Build body rows with group separators and auto-detected section rows
    body_cells = table.cells[table.header_rows:]
    body_rows_with_seps: list[Union[list[str], str]] = []
    for i, row in enumerate(body_rows):
        # Auto-detect section row: single cell spanning all columns
        is_section = (
            i < len(body_cells)
            and body_cells[i]
            and body_cells[i][0].colspan >= table.num_cols
        )
        if is_section and i > 0:
            body_rows_with_seps.append("\\midrule")
        body_rows_with_seps.append(row)
        if is_section:
            body_rows_with_seps.append("\\midrule")
        abs_idx = table.header_rows + i
        if abs_idx in gs:
            sep = gs[abs_idx]
            if isinstance(sep, list):
                body_rows_with_seps.extend(sep)
            else:
                body_rows_with_seps.append(sep)

    computed_col_spec = col_spec or _build_col_spec(table, config)

    # Caption is always passed as-is (no escaping)
    cap = caption

    # Merge spacing: user override > theme > global default
    _default = SpacingConfig()
    _theme = config.spacing or _default
    _user = spacing or _theme
    sp = SpacingConfig(**{
        f.name: getattr(_user, f.name) or getattr(_theme, f.name) or getattr(_default, f.name)
        for f in fields(_default)
    })

    ctx = {
        "col_spec": computed_col_spec,
        "header_rows": header_rows,
        "body_rows": body_rows_with_seps,
        "caption": cap,
        "label": label,
        "position": position,
        "font_size": font_size or config.font_size,
        "caption_position": config.caption_position,
        "spacing": sp,
        "resizebox": resizebox,
        "header_sep": final_header_sep,
        "wide": wide,
    }

    return tmpl.render(**ctx)


def render_to_file(
    table: TableData,
    output: Union[str, Path],
    theme: str = "three_line",
    caption: Optional[str] = None,
    label: Optional[str] = None,
    position: str = "htbp",
    raw_caption: bool = False,
    spacing: Optional[SpacingConfig] = None,
    font_size: Optional[str] = None,
    resizebox: Optional[str] = None,
    col_spec: Optional[str] = None,
    header_sep: Optional[str] = None,
    wide: bool = False,
    span_columns: Optional[bool] = None,
) -> Path:
    """Render TableData and write to a .tex file."""
    output = Path(output)
    tex = render(
        table, theme=theme, caption=caption, label=label,
        position=position, spacing=spacing,
        font_size=font_size, resizebox=resizebox, col_spec=col_spec,
        header_sep=header_sep, wide=wide, span_columns=span_columns,
    )
    output.write_text(tex)
    return output
