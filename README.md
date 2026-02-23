# pubtab

<div align="center">

  <p>
    <a href="https://pypi.org/project/pubtab/"><img src="https://img.shields.io/pypi/v/pubtab?style=flat-square&color=blue" alt="PyPI Version"/></a>
    <a href="https://pypi.org/project/pubtab/"><img src="https://img.shields.io/pypi/pyversions/pubtab?style=flat-square" alt="Python Versions"/></a>
    <img src="https://img.shields.io/badge/License-MIT-green?style=flat-square" alt="License"/>
    <a href="https://pypi.org/project/pubtab/"><img src="https://img.shields.io/pypi/dm/pubtab?style=flat-square&color=orange" alt="Downloads"/></a>
  </p>

  <strong>Language</strong>: <a href="README.md">English</a> | <a href="README.zh-CN.md">中文</a>

</div>

> Excel to publication-ready LaTeX tables — bidirectional conversion with full style preservation.

## Highlights

- **Bidirectional** — Excel ↔ LaTeX two-way conversion (`.xlsx`/`.xls` ↔ `.tex`)
- **Style-Preserving** — Colors, bold, italic, merged cells, rotation, diagonal headers fully retained
- **Zero-Config Preview** — Auto-installs TinyTeX on first use; one command to get PNG preview
- **Academic-Optimized** — Leading zero stripping, `\diagbox`, section row detection, `\cmidrule` auto-generation

## Quick Start

```bash
pip install pubtab
```

**CLI:**

```bash
# Excel → LaTeX
pubtab convert table.xlsx -o output.tex

# With options
pubtab convert table.xlsx -o output.tex --theme three_line --caption "Results" --label "tab:results" --preview

# LaTeX → Excel (reverse)
pubtab tex2xlsx paper_table.tex -o recovered.xlsx

# PNG preview from .tex
pubtab preview output.tex -o preview.png --dpi 300
```

**Python API:**

```python
import pubtab

# Excel → LaTeX
pubtab.convert("table.xlsx", output="table.tex", theme="three_line",
               caption="Experimental Results", label="tab:results")

# LaTeX → PNG preview
pubtab.preview("table.tex", dpi=300)

# LaTeX → Excel
pubtab.tex_to_excel("table.tex", "output.xlsx")
```

## Features

### Excel → LaTeX Conversion

Reads `.xlsx` (openpyxl) and `.xls` (xlrd) files, producing publication-quality LaTeX via Jinja2 templates.

**Supported cell features:**
| Feature | LaTeX Output |
|---------|-------------|
| Bold / Italic / Underline | `\textbf{}`, `\textit{}`, `\underline{}` |
| Font color | `\textcolor[RGB]{r,g,b}{}` |
| Background color | `\cellcolor[RGB]{r,g,b}` |
| Merged cells (horizontal) | `\multicolumn{n}{c}{}` |
| Merged cells (vertical) | `\multirow{n}{*}{}` |
| Text rotation | `\rotatebox[origin=c]{angle}{}` |
| Diagonal header | `\diagbox{Row}{Col}` |
| Multi-line content | `\makecell{...\\\\...}` |
| Rich text (per-segment styling) | Per-segment color/bold/italic |

**Table-level features:**
- Auto-generated `\cmidrule` from merged header cells
- Section row detection (full-width first column → auto `\midrule`)
- Group separators via `group_separators` parameter
- Configurable spacing (`tabcolsep`, `arraystretch`, rule widths)
- `resizebox` and `font_size` overrides
- `table*` for two-column spanning (`span_columns=True`)

### LaTeX → Excel Conversion

Parses LaTeX tables back to Excel with robust command support:

- `\multicolumn`, `\multirow` (including negative values)
- `\textbf`, `\textit`, `\underline`, `\emph`
- `\textcolor`, `\cellcolor`, `\rowcolor`, `\rowcolors`
- `\diagbox`, `\makecell`, `\rotatebox`
- `\newcommand`/`\renewcommand` expansion (up to 10 rounds)
- `\definecolor` custom color parsing
- 80+ LaTeX symbol → Unicode mappings (±, ×, →, ✓, α-ω, etc.)
- Nested tabular → `\makecell` conversion

### PNG Preview

Generate publication-quality PNG previews directly from `.tex` files:

```bash
pubtab preview table.tex --dpi 300
```

**TinyTeX auto-installation:** If no system `pdflatex` is found, pubtab automatically downloads and installs TinyTeX (~90 MB) to `~/.pubtab/TinyTeX/`, including required LaTeX packages (booktabs, multirow, xcolor, etc.). This is a one-time setup.

**PDF → PNG pipeline:** `pdf2image` → `qlmanage` (macOS) → `convert` (ImageMagick), using the first available tool.

For best quality, install the optional dependency:

```bash
pip install pubtab[preview]  # installs pdf2image
```

## CLI Reference

| Command | Description |
|---------|-------------|
| `pubtab convert` | Convert Excel to LaTeX |
| `pubtab tex2xlsx` | Convert LaTeX to Excel |
| `pubtab preview` | Generate PNG from .tex |
| `pubtab themes` | List available themes |

<details>
<summary>Full <code>convert</code> options</summary>

```
pubtab convert INPUT -o OUTPUT [OPTIONS]

Options:
  -o, --output TEXT          Output .tex file (required)
  -c, --config TEXT          YAML config file
  --sheet TEXT               Sheet name or 0-based index
  --theme TEXT               Theme name [default: three_line]
  --caption TEXT             Table caption
  --label TEXT               LaTeX label
  --header-rows INTEGER      Number of header rows
  --position TEXT            Float position [default: htbp]
  --font-size TEXT           Font size (e.g. footnotesize)
  --resizebox TEXT           Resize width (e.g. 0.8\textwidth)
  --col-spec TEXT            Column spec (e.g. lccc)
  --span-columns            Use table* for two-column
  --preview                 Generate PNG preview
  --dpi INTEGER             Preview DPI [default: 300]
  --header-sep TEXT          Custom header separator
  --preamble TEXT            Extra LaTeX preamble
```

</details>

## Configuration

All parameters can be set via a YAML config file, with CLI arguments taking precedence:

```yaml
theme: three_line
caption: "Experimental Results"
label: "tab:results"
header_rows: 2
span_columns: false
position: htbp
font_size: footnotesize
spacing:
  tabcolsep: "4pt"
  arraystretch: "1.2"
group_separators: [3, 6]
```

```bash
pubtab convert table.xlsx -o output.tex -c config.yaml
```

## Theme System

pubtab uses a Jinja2-based theme system. The built-in `three_line` theme produces classic booktabs-style tables.

**Custom themes:** Create a directory under `themes/` with `config.yaml` + `template.tex`:

```
my_theme/
├── config.yaml    # packages, spacing, font_size, caption_position
└── template.tex   # Jinja2 template
```

List available themes:

```bash
pubtab themes
```

## Project Structure

<details>
<summary>View project structure</summary>

```
pubtab/
├── pyproject.toml
├── README.md
├── README.zh-CN.md
├── LICENSE
└── src/pubtab/
    ├── __init__.py        # Public API: convert, preview, tex_to_excel
    ├── cli.py             # CLI (click)
    ├── models.py          # Data models (Cell, TableData, SpacingConfig, ThemeConfig)
    ├── reader.py          # Excel reader (.xlsx/.xls)
    ├── renderer.py        # LaTeX renderer (Jinja2)
    ├── tex_reader.py      # LaTeX parser (tex → TableData)
    ├── writer.py          # Excel writer
    ├── _preview.py        # PNG preview (TinyTeX auto-install)
    ├── config.py          # YAML config loader
    ├── utils.py           # LaTeX escaping, color conversion
    └── themes/
        └── three_line/
            ├── config.yaml
            └── template.tex
```

</details>

## Contributing

Issues and pull requests are welcome at [GitHub](https://github.com/gaoruizhang/pubtab).

## License

[MIT](LICENSE)
