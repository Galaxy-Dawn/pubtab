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
@click.option("--resizebox", default=None, help="(Deprecated) Resize width (e.g. 0.8\\textwidth).")
@click.option("--with-resizebox", "with_resizebox", is_flag=True, help="Wrap tabular with \\resizebox.")
@click.option("--without-resizebox", "without_resizebox", is_flag=True, help="Disable \\resizebox wrapper.")
@click.option("--resizebox-width", default="\\linewidth", show_default=True, help="Resizebox width used with --with-resizebox.")
@click.option("--col-spec", default=None, help="Column spec (e.g. lccc).")
@click.option("--dpi", default=None, type=int, help="Preview DPI [default: 300].")
@click.option("--header-sep", default=None, help="Custom header separator.")
@click.option("--upright-scripts", is_flag=True, default=False, help="Wrap sub/superscript content in \\mathrm{} for upright rendering.")
@click.option(
    "--latex-backend",
    default=None,
    type=click.Choice(["tabular", "tabularray"]),
    help="LaTeX backend used for export.",
)
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
    with_resizebox: bool,
    without_resizebox: bool,
    resizebox_width: str,
    col_spec: str | None,
    dpi: int | None,
    header_sep: str | None,
    upright_scripts: bool,
    latex_backend: str | None,
) -> None:
    """Convert an Excel file to LaTeX.

    Use --config to load all settings from a YAML file.
    CLI flags override config values.
    """
    from pathlib import Path
    import pubtab as pt

    input_path = Path(input_file)
    output_path = Path(output)
    if input_path.is_dir() and output_path.suffix.lower() == ".tex":
        raise click.BadParameter(
            "When INPUT_FILE is a directory, --output must be a directory path.",
            param_hint="--output",
        )

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
    if with_resizebox and without_resizebox:
        raise click.BadParameter(
            "--with-resizebox and --without-resizebox cannot be used together."
        )
    if with_resizebox:
        kwargs["resizebox"] = resizebox_width
    elif without_resizebox:
        # Explicit None overrides config/default resizebox settings.
        kwargs["resizebox"] = None
    elif resizebox is not None:
        kwargs["resizebox"] = resizebox
    if col_spec is not None:
        kwargs["col_spec"] = col_spec
    if dpi is not None:
        kwargs["dpi"] = dpi
    if header_sep is not None:
        kwargs["header_sep"] = header_sep
    if upright_scripts:
        kwargs["upright_scripts"] = True
    if latex_backend is not None:
        kwargs["latex_backend"] = latex_backend

    pt.xlsx2tex(input_file, output, **kwargs)

    if input_path.is_dir():
        excel_files = sorted(
            [p for p in input_path.iterdir() if p.is_file() and p.suffix.lower() in (".xlsx", ".xls")],
            key=lambda p: p.name.lower(),
        )
        if not excel_files:
            click.echo(f"No .xlsx/.xls files found in: {input_path}")
            return
        if sheet is None:
            try:
                from .reader import list_excel_sheets

                tex_total = sum(len(list_excel_sheets(p)) for p in excel_files)
                first_sheet_count = len(list_excel_sheets(excel_files[0]))
            except Exception:
                tex_total = len(excel_files)
                first_sheet_count = 1
        else:
            tex_total = len(excel_files)
            first_sheet_count = 1

        first_tex = output_path / (
            f"{excel_files[0].stem}_sheet01.tex"
            if first_sheet_count > 1
            else f"{excel_files[0].stem}.tex"
        )
        if tex_total <= 1:
            click.echo(f"Written: {first_tex}")
        else:
            click.echo(f"Written: {first_tex} (+{tex_total - 1} additional tex files)")

        if do_preview:
            first_png = first_tex.with_suffix(".png")
            if tex_total <= 1:
                click.echo(f"Preview: {first_png}")
            else:
                click.echo(f"Preview: {first_png} (+{tex_total - 1} additional sheet previews)")
        return

    sheet_count = 1
    if sheet is None:
        try:
            from .reader import list_excel_sheets

            sheet_count = len(list_excel_sheets(input_file))
        except Exception:
            sheet_count = 1

    output_path = Path(output)
    first_tex = output_path
    if sheet_count > 1:
        first_tex = output_path.with_name(f"{output_path.stem}_sheet01.tex")

    if sheet_count <= 1:
        click.echo(f"Written: {first_tex}")
    else:
        click.echo(f"Written: {first_tex} (+{sheet_count - 1} additional sheet tex files)")

    if do_preview:
        png_path = first_tex.with_suffix(".png")
        if sheet_count <= 1:
            click.echo(f"Preview: {png_path}")
        else:
            click.echo(f"Preview: {png_path} (+{sheet_count - 1} additional sheet previews)")


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
    from pathlib import Path

    import pubtab as pt

    input_path = Path(input_file)
    output_path = Path(output)
    if input_path.is_dir() and output_path.suffix.lower() == ".xlsx":
        raise click.BadParameter(
            "When INPUT_FILE is a directory, --output must be a directory path.",
            param_hint="--output",
        )

    result = pt.tex_to_excel(input_file, output)
    if input_path.is_dir():
        tex_files = sorted(
            [p for p in input_path.iterdir() if p.is_file() and p.suffix.lower() == ".tex"],
            key=lambda p: p.name.lower(),
        )
        if not tex_files:
            click.echo(f"No .tex files found in: {input_path}")
            return
        first_xlsx = output_path / f"{tex_files[0].stem}.xlsx"
        if len(tex_files) <= 1:
            click.echo(f"Written: {first_xlsx}")
        else:
            click.echo(f"Written: {first_xlsx} (+{len(tex_files) - 1} additional xlsx files)")
        return

    click.echo(f"Written: {result}")


@main.command("preview")
@click.argument("tex_file")
@click.option("-o", "--output", default=None, help="Output path.")
@click.option("--theme", default="three_line", help="Theme name.")
@click.option(
    "--latex-backend",
    default=None,
    type=click.Choice(["tabular", "tabularray"]),
    help="LaTeX backend used for preview document assembly. Auto-detected from tblr content when omitted.",
)
@click.option("--dpi", default=300, type=int, help="PNG resolution.")
@click.option("--format", "fmt", default="png", type=click.Choice(["png", "pdf"]), help="Output format [default: png].")
@click.option("--preamble", default=None, help="Extra LaTeX preamble (e.g. custom commands).")
def preview_cmd(
    tex_file: str,
    output: str | None,
    theme: str,
    latex_backend: str | None,
    dpi: int,
    fmt: str,
    preamble: str | None,
) -> None:
    """Generate PNG or PDF from a .tex file."""
    from pathlib import Path

    import pubtab as pt

    input_path = Path(tex_file)
    if input_path.is_dir() and output is not None and Path(output).suffix.lower() in (".png", ".pdf"):
        raise click.BadParameter(
            "When TEX_FILE is a directory, --output must be a directory path.",
            param_hint="--output",
        )

    result = pt.preview(
        tex_file,
        output=output,
        theme=theme,
        latex_backend=latex_backend,
        dpi=dpi,
        preamble=preamble,
        format=fmt,
    )
    if input_path.is_dir():
        tex_files = sorted(
            [p for p in input_path.iterdir() if p.is_file() and p.suffix.lower() == ".tex"],
            key=lambda p: p.name.lower(),
        )
        if not tex_files:
            click.echo(f"No .tex files found in: {input_path}")
            return
        output_dir = Path(output) if output is not None else (input_path / f"preview_{fmt}")
        first_out = output_dir / f"{tex_files[0].stem}.{fmt}"
        if len(tex_files) <= 1:
            click.echo(f"Output: {first_out}")
        else:
            click.echo(f"Output: {first_out} (+{len(tex_files) - 1} additional {fmt} files)")
        return

    click.echo(f"Output: {result}")


# Hidden alias for backward compatibility
main.add_command(xlsx2tex_cmd, "convert")
