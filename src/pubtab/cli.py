"""CLI entry point for pubtab."""

from __future__ import annotations

import click

from .themes import list_themes as _list_themes


@click.group()
def main() -> None:
    """pubtab — Excel to publication-ready LaTeX tables."""


@main.command("xlsx2tex")
@click.argument("input_file")
@click.option("-o", "--output", required=True, help="Output .tex file path.")
@click.option("-c", "--config", default=None, help="YAML config file path.")
@click.option("--sheet", default=None, help="Sheet name or index.")
@click.option("--theme", default=None, help="Theme name.")
@click.option("--caption", default=None, help="Table caption.")
@click.option("--label", default=None, help="LaTeX label.")
@click.option("--header-rows", default=None, type=int, help="Number of header rows.")
@click.option("--span-columns", is_flag=True, default=False, help="Use table* for two-column spanning.")
@click.option("--preview", "do_preview", is_flag=True, help="Generate PNG preview.")
@click.option("--position", default=None, help="Float position [default: htbp].")
@click.option("--font-size", default=None, help="Font size (e.g. footnotesize).")
@click.option("--resizebox", default=None, help="Resize width (e.g. 0.8\\textwidth).")
@click.option("--col-spec", default=None, help="Column spec (e.g. lccc).")
@click.option("--dpi", default=None, type=int, help="Preview DPI [default: 300].")
@click.option("--header-sep", default=None, help="Custom header separator.")
@click.option("--upright-scripts", is_flag=True, default=False, help="Wrap sub/superscript content in \\mathrm{} for upright rendering.")
def xlsx2tex_cmd(
    input_file: str,
    output: str,
    config: str | None,
    sheet: str | None,
    theme: str | None,
    caption: str | None,
    label: str | None,
    header_rows: int | None,
    span_columns: bool,
    do_preview: bool,
    position: str | None,
    font_size: str | None,
    resizebox: str | None,
    col_spec: str | None,
    dpi: int | None,
    header_sep: str | None,
    upright_scripts: bool,
) -> None:
    """Convert an Excel file to LaTeX.

    Use --config to load all settings from a YAML file.
    CLI flags override config values.
    """
    from pathlib import Path
    import pubtab as pt

    # Parse sheet as int if numeric
    sheet_val: str | int | None = sheet
    if sheet is not None:
        try:
            sheet_val = int(sheet)
        except ValueError:
            pass

    # Build kwargs, only pass explicitly set values
    kwargs = {}
    if config is not None:
        kwargs["config"] = config
    if sheet_val is not None:
        kwargs["sheet"] = sheet_val
    if theme is not None:
        kwargs["theme"] = theme
    if caption is not None:
        kwargs["caption"] = caption
    if label is not None:
        kwargs["label"] = label
    if header_rows is not None:
        kwargs["header_rows"] = header_rows
    if span_columns:
        kwargs["span_columns"] = True
    if do_preview:
        kwargs["preview"] = True
    if position is not None:
        kwargs["position"] = position
    if font_size is not None:
        kwargs["font_size"] = font_size
    if resizebox is not None:
        kwargs["resizebox"] = resizebox
    if col_spec is not None:
        kwargs["col_spec"] = col_spec
    if dpi is not None:
        kwargs["dpi"] = dpi
    if header_sep is not None:
        kwargs["header_sep"] = header_sep
    if upright_scripts:
        kwargs["upright_scripts"] = True

    pt.xlsx2tex(input_file, output, **kwargs)
    click.echo(f"Written: {output}")

    if do_preview:
        png_path = Path(output).with_suffix(".png")
        click.echo(f"Preview: {png_path}")


@main.command("themes")
def themes_cmd() -> None:
    """List available themes."""
    for name in sorted(_list_themes()):
        click.echo(f"  {name}")


@main.command("tex2xlsx")
@click.argument("input_file")
@click.option("-o", "--output", required=True, help="Output .xlsx file path.")
def tex2xlsx(input_file: str, output: str) -> None:
    """Convert a LaTeX .tex file to Excel .xlsx."""
    import pubtab as pt

    result = pt.tex_to_excel(input_file, output)
    click.echo(f"Written: {result}")


@main.command("preview")
@click.argument("tex_file")
@click.option("-o", "--output", default=None, help="Output path.")
@click.option("--theme", default="three_line", help="Theme name.")
@click.option("--dpi", default=300, type=int, help="PNG resolution.")
@click.option("--format", "fmt", default="png", type=click.Choice(["png", "pdf"]), help="Output format [default: png].")
@click.option("--preamble", default=None, help="Extra LaTeX preamble (e.g. custom commands).")
def preview_cmd(tex_file: str, output: str | None, theme: str, dpi: int, fmt: str, preamble: str | None) -> None:
    """Generate PNG or PDF from a .tex file."""
    import pubtab as pt

    result = pt.preview(tex_file, output=output, theme=theme, dpi=dpi, preamble=preamble, format=fmt)
    click.echo(f"Output: {result}")


# Hidden alias for backward compatibility
main.add_command(xlsx2tex_cmd, "convert")
