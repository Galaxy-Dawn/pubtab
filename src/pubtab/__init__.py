"""pubtab — Excel to publication-ready LaTeX tables."""

from __future__ import annotations

from importlib.metadata import version as _version

__version__ = _version("pubtab")

from pathlib import Path
from typing import Callable, Dict, List, Optional, Union

from .models import Cell, SpacingConfig, TableData
from .reader import read_excel
from .renderer import render

__all__ = ["convert", "preview", "tex_to_excel", "SpacingConfig"]

# Sentinel for detecting unset kwargs
_UNSET = object()


def convert(
    input_file: Union[str, Path],
    output: Union[str, Path],
    config: Optional[Union[str, Path]] = None,
    sheet: Optional[Union[str, int]] = _UNSET,
    theme: str = _UNSET,
    caption: Optional[str] = _UNSET,
    label: Optional[str] = _UNSET,
    header_rows: int = _UNSET,
    position: Optional[str] = _UNSET,
    spacing: Optional[SpacingConfig] = _UNSET,
    font_size: Optional[str] = _UNSET,
    resizebox: Optional[str] = _UNSET,
    col_spec: Optional[str] = _UNSET,
    header_sep: Optional[str] = _UNSET,
    header_cmidrule: bool = _UNSET,
    span_columns: bool = _UNSET,
    custom_header: Optional[List[List[Cell]]] = _UNSET,
    group_separators: Optional[Union[List[int], Dict[int, str]]] = _UNSET,
    cell_formatter: Optional[Callable[[int, int, Cell], Cell]] = _UNSET,
    num_cols: Optional[int] = _UNSET,
    preview: bool = _UNSET,
    preamble: Optional[str] = _UNSET,
    dpi: int = _UNSET,
    # Deprecated aliases
    wide: bool = _UNSET,
    raw_caption: bool = _UNSET,
) -> str:
    """Convert an Excel file to LaTeX and write to .tex file.

    Args:
        input_file: Path to .xlsx or .xls file.
        output: Output .tex file path.
        config: Path to YAML config file (explicit kwargs override config values).
        sheet: Sheet name or 0-based index.
        theme: Theme name.
        caption: Table caption (always passed as-is, no escaping).
        label: LaTeX label.
        header_rows: Number of header rows in Excel.
        position: Table float position (e.g. "t", "htbp").
        spacing: Override spacing config.
        font_size: Override font size (e.g. "footnotesize").
        resizebox: Resize width (e.g. "0.8\\textwidth").
        col_spec: Column specification (e.g. "lccc").
        header_sep: Custom header separator (auto-generated from merged cells if None).
        span_columns: Use table* for two-column spanning.
        custom_header: Custom header rows (replaces Excel headers).
        group_separators: List[int] row indices or Dict[int, str] for custom separators.
        cell_formatter: Callback(row_idx, col_idx, cell) -> Cell.
        num_cols: Override column count.
        preview: Generate PNG preview.
        preamble: Extra LaTeX preamble for preview.
        dpi: Preview resolution.
        wide: Deprecated alias for span_columns.
        raw_caption: Deprecated, ignored.

    Returns:
        LaTeX table string.
    """
    # Defaults
    defaults = dict(
        sheet=None, theme="three_line", caption=None, label=None,
        header_rows=None, position="htbp", spacing=None,
        font_size=None, resizebox=None, col_spec=None, header_sep=None,
        header_cmidrule=True, span_columns=False, custom_header=None, group_separators=None,
        cell_formatter=None, num_cols=None, preview=False, preamble=None,
        dpi=300,
    )

    # Load YAML config
    cfg_formatter = None
    if config is not None:
        from .config import load_config
        cfg_kwargs, cfg_formatter = load_config(config)
        defaults.update(cfg_kwargs)

    # Explicit kwargs override config and defaults
    # Handle deprecated wide → span_columns alias
    _span = span_columns if span_columns is not _UNSET else (wide if wide is not _UNSET else _UNSET)
    explicit = dict(
        sheet=sheet, theme=theme, caption=caption, label=label,
        header_rows=header_rows, position=position,
        spacing=spacing, font_size=font_size, resizebox=resizebox,
        col_spec=col_spec, header_sep=header_sep, header_cmidrule=header_cmidrule,
        span_columns=_span,
        custom_header=custom_header,
        group_separators=group_separators,
        cell_formatter=cell_formatter, num_cols=num_cols, preview=preview,
        preamble=preamble, dpi=dpi,
    )
    p = {k: (v if v is not _UNSET else defaults[k]) for k, v in explicit.items()}

    if p["cell_formatter"] is None and cfg_formatter is not None:
        p["cell_formatter"] = cfg_formatter

    # Read Excel
    table = read_excel(input_file, sheet=p["sheet"], header_rows=p["header_rows"])

    if p["custom_header"] is not None:
        data_rows = table.cells[table.header_rows:]
        if p["cell_formatter"]:
            data_rows = [
                [p["cell_formatter"](r, c, cell) for c, cell in enumerate(row)]
                for r, row in enumerate(data_rows)
            ]
        all_cells = p["custom_header"] + data_rows
        hc = len(p["custom_header"])
        nc = p["num_cols"] or table.num_cols
        table = TableData(
            cells=all_cells, num_rows=len(all_cells), num_cols=nc,
            header_rows=hc, group_separators=p["group_separators"] or {},
        )
    elif p["group_separators"]:
        table = TableData(
            cells=table.cells, num_rows=table.num_rows, num_cols=table.num_cols,
            header_rows=table.header_rows, group_separators=p["group_separators"],
        )

    tex = render(
        table, theme=p["theme"], caption=p["caption"], label=p["label"],
        position=p["position"], spacing=p["spacing"],
        font_size=p["font_size"], resizebox=p["resizebox"], col_spec=p["col_spec"],
        header_sep=p["header_sep"], header_cmidrule=p["header_cmidrule"],
        span_columns=p["span_columns"],
    )
    Path(output).write_text(tex)

    if p["preview"]:
        from ._preview import preview as _preview
        png_path = Path(output).with_suffix(".png")
        _preview(output, output=png_path, theme=p["theme"], dpi=p["dpi"],
                 preamble=p["preamble"])

    return tex


def preview(
    tex_input: Union[str, Path],
    output: Optional[Union[str, Path]] = None,
    theme: str = "three_line",
    dpi: int = 300,
    preamble: Optional[str] = None,
) -> Path:
    """Generate a PNG preview from LaTeX content or .tex file.

    Args:
        tex_input: LaTeX string or path to .tex file.
        output: Output PNG path.
        theme: Theme name.
        dpi: Resolution.
        preamble: Extra LaTeX preamble (e.g. custom commands).

    Returns:
        Path to generated PNG.
    """
    from ._preview import preview as _preview

    return _preview(tex_input, output=output, theme=theme, dpi=dpi, preamble=preamble)


def tex_to_excel(
    input_file: Union[str, Path],
    output: Union[str, Path],
) -> Path:
    """Convert a LaTeX .tex file to Excel .xlsx.

    Args:
        input_file: Path to .tex file.
        output: Output .xlsx file path.

    Returns:
        Path to generated .xlsx file.
    """
    from .tex_reader import read_tex
    from .writer import write_excel

    tex_content = Path(input_file).read_text()
    table = read_tex(tex_content)
    return write_excel(table, output)


# Deprecated alias
generate_preview = preview
