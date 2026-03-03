"""Jinja2 rendering engine — converts TableData to LaTeX."""

from __future__ import annotations

import re
from dataclasses import fields, replace
from pathlib import Path
from typing import Optional, Union

from jinja2 import Environment

from .models import Cell, CellStyle, SpacingConfig, TableData, ThemeConfig
from .themes import load_theme
from .utils import format_number, hex_to_latex_color, latex_escape

# Detect math expressions like P_t, F_1, BLEU_1, COT_web, F1_max that need $...$ wrapping.
# Matches: letter + optional alphanumeric base + _ or ^ + alphanumeric subscript/superscript.
_MATH_EXPR_RE = re.compile(r'^[A-Za-z][A-Za-z0-9]*[_^][A-Za-z0-9]+$')
# Detect embedded math sub-expressions like F_1, F1_max inside mixed strings like "F_1-max".
_MATH_SUBEXPR_RE = re.compile(r'([A-Za-z][A-Za-z0-9]*[_^][A-Za-z0-9]+)')
# Detect numeric subscripts like 00.0_00.0 (from $_{...}$ stripped by parser).
_NUMERIC_SUBSCRIPT_RE = re.compile(r'([\d.]+)_([\d.]+)')
# Detect numeric values with signed subscripts like 37.54_+3.80, 94.00_-1.02
# Pattern: NUMBER _ [+-] NUMBER
_NUMERIC_SIGNED_SUBSCRIPT_RE = re.compile(r'^([\d.]+)_([+-][\d.]+)$')
# Detect $...$ math segments mixed with plain text, e.g. "$F_1$-max"
_MATH_DOLLAR_RE = re.compile(r'(\$[^$]+\$)')
# Detect expressions like AR(AP)_LM-O: base with parens + subscript with hyphens
# Pattern: ALPHANUM(ALPHANUM)_ALPHANUM[-ALPHANUM]*
_MATH_PARENS_HYPHEN_RE = re.compile(
    r'^([A-Za-z][A-Za-z0-9]*)\(([A-Za-z0-9]+)\)([_^])([A-Za-z][A-Za-z0-9-]*)$'
)
# Special chars that still need escaping inside $...$
_MATH_INNER_SPECIAL = {"&": r"\&", "%": r"\%", "#": r"\#"}
_MATH_INNER_RE = re.compile("|".join(re.escape(k) for k in _MATH_INNER_SPECIAL))
# Reverse mapping: Unicode mathcal → \mathcal{X}
_UNICODE_TO_LATEX = {
    "𝒜": r"\mathcal{A}", "ℬ": r"\mathcal{B}", "𝒞": r"\mathcal{C}",
    "𝒟": r"\mathcal{D}", "ℰ": r"\mathcal{E}", "ℱ": r"\mathcal{F}",
    "𝒢": r"\mathcal{G}", "ℋ": r"\mathcal{H}", "ℐ": r"\mathcal{I}",
    "𝒥": r"\mathcal{J}", "𝒦": r"\mathcal{K}", "ℒ": r"\mathcal{L}",
    "ℳ": r"\mathcal{M}", "𝒩": r"\mathcal{N}", "𝒪": r"\mathcal{O}",
    "𝒫": r"\mathcal{P}", "𝒬": r"\mathcal{Q}", "ℛ": r"\mathcal{R}",
    "𝒮": r"\mathcal{S}", "𝒯": r"\mathcal{T}", "𝒰": r"\mathcal{U}",
    "𝒱": r"\mathcal{V}", "𝒲": r"\mathcal{W}", "𝒳": r"\mathcal{X}",
    "𝒴": r"\mathcal{Y}", "𝒵": r"\mathcal{Z}",
    "↑": r"\uparrow", "↓": r"\downarrow", "→": r"\rightarrow", "←": r"\leftarrow",
    "↗": r"\nearrow", "↘": r"\searrow",
    "†": r"\dagger", "‡": r"\ddagger",
    # Greek lowercase
    "α": r"\alpha", "β": r"\beta", "γ": r"\gamma", "δ": r"\delta",
    "ε": r"\epsilon", "ζ": r"\zeta", "η": r"\eta", "θ": r"\theta",
    "κ": r"\kappa", "λ": r"\lambda", "μ": r"\mu", "π": r"\pi",
    "ρ": r"\rho", "σ": r"\sigma", "τ": r"\tau", "φ": r"\phi", "ω": r"\omega",
    # Greek uppercase
    "Σ": r"\Sigma", "Ω": r"\Omega", "Δ": r"\Delta",
    # Math symbols
    "∞": r"\infty", "ℓ": r"\ell",
}
_UNICODE_MATH_RE = re.compile("|".join(re.escape(k) for k in _UNICODE_TO_LATEX))


def _cleaned_to_math(s: str) -> str:
    """Convert cleaned math text back to LaTeX $...$ expression."""
    t = s
    t = re.sub(r'\^\(([^)]+)\)', r'^{\1}', t)       # ^(l-1) → ^{l-1}
    t = re.sub(r'\^([A-Za-z0-9])', r'^{\1}', t)      # ^T → ^{T}
    # Subscript handling:
    # Multi-char lowercase subscript at word boundary (e.g., _en, _in) → _{en}, _{in}
    # Use placeholder to avoid double-processing
    t = re.sub(r'_([a-z]{2,})(?=[^a-z]|$|[A-Z])', r'@SUBBRACE@\1@END@', t)
    # Single char subscript (including uppercase, digits, unicode)
    t = re.sub(r'_([A-Za-z0-9]|.)', r'_{\1}', t)
    # Restore multi-char subscripts
    t = t.replace('@SUBBRACE@', '_{').replace('@END@', '}')
    t = _UNICODE_MATH_RE.sub(lambda m: _UNICODE_TO_LATEX[m.group()], t)
    return "$" + _MATH_INNER_RE.sub(lambda m: _MATH_INNER_SPECIAL[m.group()], t) + "$"


def _embed_math_subexprs(s: str) -> str:
    """Wrap embedded math sub-expressions in $...$, escape surrounding plain text.

    e.g. "F_1-max" → "F$_{1}$-max"
    """
    parts = _MATH_SUBEXPR_RE.split(s)
    result = []
    for i, part in enumerate(parts):
        if i % 2 == 1:  # captured math sub-expression
            _m = re.match(r'^([A-Za-z][A-Za-z0-9]*)([_^])([A-Za-z0-9]+)$', part)
            if _m:
                # Keep base in text mode; subscript/superscript in math mode
                base = latex_escape(_m.group(1))
                script = _m.group(2)
                content = _m.group(3)
                result.append(f"{base}${script}{{{content}}}$")
            else:
                result.append("$" + part + "$")
        else:
            result.append(latex_escape(part))
    return "".join(result)


def _process_with_dollar_math(s: str) -> str:
    """Handle strings that contain $...$ math segments mixed with plain text.

    e.g. "$F_1$-max" → "$F_1$-max"  (keep math as-is, escape plain parts)
    """
    parts = _MATH_DOLLAR_RE.split(s)
    result = []
    for i, part in enumerate(parts):
        if i % 2 == 1:  # $...$ math segment — keep as-is
            result.append(part)
        else:
            if _MATH_SUBEXPR_RE.search(part):
                result.append(_embed_math_subexprs(part))
            elif _NUMERIC_SUBSCRIPT_RE.search(part):
                result.append(_NUMERIC_SUBSCRIPT_RE.sub(r"\1$_{\2}$", part))
            else:
                result.append(latex_escape(part))
    return "".join(result)


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
        s = str(val)
        if re.match(r'^\$[^$]+\$$', s):
            text = s  # entire value is a math expression, pass through as-is
        elif _MATH_PARENS_HYPHEN_RE.match(s):
            # Handle AR(AP)_LM-O style: base(parens)_subscript-with-hyphens
            _m = _MATH_PARENS_HYPHEN_RE.match(s)
            # Keep base(parens) in text mode so \textbf/\textit apply;
            # subscript in math mode.
            _base = latex_escape(f"{_m.group(1)}({_m.group(2)})")
            _op, _sub = _m.group(3), _m.group(4)
            text = f"{_base}${_op}{{{_sub}}}$"
        elif _MATH_EXPR_RE.match(s):
            _m = re.match(r'^([A-Za-z][A-Za-z0-9]*)([_^])([A-Za-z0-9]+)$', s)
            if _m and len(_m.group(1)) > 1:
                # Multi-letter base (e.g. CHAIR_S, BLEU_1, hazel_nut): keep base in
                # text mode so \textbf/\textit apply; subscript in math mode.
                _base = latex_escape(_m.group(1))
                _op, _sub = _m.group(2), _m.group(3)
                text = f"{_base}${_op}{{{_sub}}}$"
            elif _m:
                # Single-letter base (e.g. M^2, p_task): keep base in text mode
                _base = latex_escape(_m.group(1))
                text = f"{_base}${_m.group(2)}{{{_m.group(3)}}}$"
            else:
                text = "$" + _MATH_INNER_RE.sub(lambda m: _MATH_INNER_SPECIAL[m.group()], s) + "$"
        elif _MATH_DOLLAR_RE.search(s):
            text = _process_with_dollar_math(s)
        elif re.match(r'^[A-Za-z][A-Za-z0-9-]*\^', s):
            # Base (single or multi-letter, possibly with hyphens) with ^(...) or ^X: keep base in text mode
            _m = re.match(r'^([A-Za-z][A-Za-z0-9-]*)(\^(?:\([^)]+\)|[A-Za-z0-9]+))(.*)', s)
            if _m:
                _base = latex_escape(_m.group(1))
                _sup_raw = _m.group(2)[1:]  # strip leading ^
                _sup = _sup_raw.strip('()')
                _rest_raw = _m.group(3)
                # Convert unicode math symbols in rest (e.g. ↑→\uparrow) into the math expression
                _rest_conv = _UNICODE_MATH_RE.sub(lambda m: _UNICODE_TO_LATEX[m.group()], _rest_raw)
                if _rest_conv != _rest_raw:
                    text = f"{_base}$^{{{_sup}}}{_rest_conv}$"
                elif '^' in _rest_raw or '_' in _rest_raw:
                    # Rest contains more math - use _cleaned_to_math and merge into one expression
                    _rest_math = _cleaned_to_math(_rest_raw).strip('$')
                    text = f"{_base}$^{{{_sup}}}{_rest_math}$"
                else:
                    text = f"{_base}$^{{{_sup}}}${latex_escape(_rest_raw)}"
            else:
                text = _cleaned_to_math(s)
        elif '^' in s:
            text = _cleaned_to_math(s)
        elif _MATH_SUBEXPR_RE.search(s):
            text = _embed_math_subexprs(s)
        elif _NUMERIC_SUBSCRIPT_RE.search(s):
            text = _NUMERIC_SUBSCRIPT_RE.sub(r"\1$_{\2}$", s)
        elif _NUMERIC_SIGNED_SUBSCRIPT_RE.match(s):
            # Handle 37.54_+3.80 style: number with signed subscript (+/-)
            _m = _NUMERIC_SIGNED_SUBSCRIPT_RE.match(s)
            text = f"{_m.group(1)}$_{{{_m.group(2)}}}$"
        elif _UNICODE_MATH_RE.search(s) and ('_' in s or '^' in s):
            text = _cleaned_to_math(s)
        else:
            text = latex_escape(s)

    # Rich segments: per-segment color/bold/italic/underline
    if cell.rich_segments and not cell.style.raw_latex:
        parts = []
        for seg in cell.rich_segments:
            seg_text, seg_color = seg[0], seg[1]
            seg_bold = seg[2] if len(seg) > 2 else False
            seg_italic = seg[3] if len(seg) > 3 else False
            seg_underline = seg[4] if len(seg) > 4 else False
            s = latex_escape(seg_text)
            if seg_bold:
                s = f"\\textbf{{{s}}}"
            if seg_italic:
                s = f"\\textit{{{s}}}"
            if seg_underline:
                s = f"\\underline{{{s}}}"
            if seg_color:
                rgb = hex_to_latex_color(seg_color)
                s = f"\\textcolor[RGB]{{{rgb}}}{{{s}}}"
            parts.append(s)
        text = "".join(parts)

    # Multi-line cell: convert \n back to \makecell{...\\...}
    if "\n" in text:
        text = "\\makecell{" + text.replace("\n", "\\\\") + "}"

    # Diagbox: diagonal header cell (before styling so bold/italic wraps it)
    if cell.style.diagbox:
        parts = cell.style.diagbox
        text = f"\\diagbox{{{parts[0]}}}{{{parts[1]}}}"

    # Apply styling (skip if raw_latex — user controls formatting)
    if not cell.style.raw_latex and not cell.rich_segments:
        if cell.style.bold:
            text = f"\\textbf{{{text}}}"
        if cell.style.italic:
            text = f"\\textit{{{text}}}"
        if cell.style.underline:
            text = f"\\underline{{{text}}}"
        if cell.style.color:
            rgb = hex_to_latex_color(cell.style.color)
            text = f"\\textcolor[RGB]{{{rgb}}}{{{text}}}"
    if not cell.style.raw_latex:
        if cell.style.rotation:
            if cell.rowspan > 1:
                # No [origin=c] for multirow: text extends upward, avoiding bottomrule overflow
                text = f"\\rotatebox{{{cell.style.rotation}}}{{{text}}}"
            else:
                text = f"\\rotatebox[origin=c]{{{cell.style.rotation}}}{{{text}}}"
    if cell.rowspan > 1:
        text = f"\\multirow{{{cell.rowspan}}}{{*}}{{{text}}}"

    if cell.colspan > 1:
        align = cell.style.alignment[0] if cell.style.alignment else "c"
        if cell.style.bg_color:
            # \columncolor in the col-spec colors the entire multicolumn span
            rgb = hex_to_latex_color(cell.style.bg_color)
            text = f"\\multicolumn{{{cell.colspan}}}{{>{{\\columncolor[RGB]{{{rgb}}}}}{align}}}{{{text}}}"
        else:
            text = f"\\multicolumn{{{cell.colspan}}}{{{align}}}{{{text}}}"

    # cellcolor for single-column cells (multicolumn case handled above)
    if cell.style.bg_color and cell.colspan <= 1:
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


def _auto_cmidrule(table_cells: list[list[Cell]], row_idx: int, num_cols: int) -> Optional[str]:
    """Auto-generate cline between row_idx and row_idx+1, skipping multirow cells."""
    skip_cols: set[int] = set()
    for r in range(row_idx + 1):
        for i, cell in enumerate(table_cells[r]):
            col = i + 1
            if cell.rowspan > 1 and r + cell.rowspan > row_idx + 1:
                span = max(cell.colspan, 1)
                for c in range(col, col + span):
                    skip_cols.add(c)

    if len(skip_cols) >= num_cols:
        return None

    rules = []
    start = None
    for c in range(1, num_cols + 1):
        if c not in skip_cols:
            if start is None:
                start = c
        else:
            if start is not None:
                rules.append(f"\\cline{{{start}-{c - 1}}}")
                start = None
    if start is not None:
        rules.append(f"\\cline{{{start}-{num_cols}}}")

    return " ".join(rules) if rules else None


def _normalize_group_separators(gs):
    """Convert List[int] to Dict[int, str] if needed."""
    if isinstance(gs, list):
        return {idx: "\\midrule" for idx in gs}
    return gs or {}


def _row_uniform_bg(cells: list) -> Optional[str]:
    """Return bg_color if all visible cells share the same non-None bg_color, else None."""
    skip = 0
    color: Optional[str] = None
    for cell in cells:
        if skip > 0:
            skip -= 1
            continue
        if cell.colspan > 1:
            skip = cell.colspan - 1
        if cell.style.bg_color is None:
            return None
        if color is None:
            color = cell.style.bg_color
        elif color != cell.style.bg_color:
            return None
    return color


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
    header_cmidrule: bool = True,
    wide: bool = False,
    span_columns: Optional[bool] = None,
    upright_scripts: bool = False,
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
        # Use \rowcolor when all visible cells share the same bg_color.
        # \rowcolor alone only colors the first physical column of each multicolumn span,
        # so multicolumn cells also keep \columncolor in their col-spec for full coverage.
        row_bg = _row_uniform_bg(row)
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
        # Prepend \rowcolor to first cell so the entire row (incl. multicolumn spans) is colored
        if row_bg and latex_row:
            rgb = hex_to_latex_color(row_bg)
            latex_row[0] = f"\\rowcolor[RGB]{{{rgb}}}" + latex_row[0]
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
    elif header_sep is None and table.header_rows > 1 and header_cmidrule:
        # Auto-generate cmidrule between header rows from merged cells
        header_rows = []
        for i, row in enumerate(raw_header_rows):
            header_rows.append(row)
            if i < len(raw_header_rows) - 1:
                rule = _auto_cmidrule(table.cells, i, table.num_cols)
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
        # Auto-detect section row: first cell spans most columns, rest empty
        is_section = False
        if i < len(body_cells) and body_cells[i]:
            c0 = body_cells[i][0]
            if c0.colspan >= table.num_cols:
                is_section = True
            elif c0.colspan >= table.num_cols - 1 and all(
                not c.value and c.value != 0 for c in body_cells[i][1:]
            ):
                is_section = True
        if is_section and i > 0:
            # Only add auto-before if previous item is not already a midrule
            if not body_rows_with_seps or body_rows_with_seps[-1] != "\\midrule":
                body_rows_with_seps.append("\\midrule")
        body_rows_with_seps.append(row)
        abs_idx = table.header_rows + i
        has_group_sep = abs_idx in gs
        if is_section and not has_group_sep:
            body_rows_with_seps.append("\\midrule")
        if has_group_sep:
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

    result = tmpl.render(**ctx)
    if upright_scripts:
        result = re.sub(r'([_^])\{([^}\\]+)\}', r'\1{\\mathrm{\2}}', result)
    return result


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
    header_cmidrule: bool = True,
    wide: bool = False,
    span_columns: Optional[bool] = None,
    upright_scripts: bool = False,
) -> Path:
    """Render TableData and write to a .tex file."""
    output = Path(output)
    tex = render(
        table, theme=theme, caption=caption, label=label,
        position=position, spacing=spacing,
        font_size=font_size, resizebox=resizebox, col_spec=col_spec,
        header_sep=header_sep, header_cmidrule=header_cmidrule,
        wide=wide, span_columns=span_columns, upright_scripts=upright_scripts,
    )
    output.write_text(tex)
    return output
