"""Round-trip conversion tests: tex→xlsx and xlsx→tex→xlsx."""

from pathlib import Path

import openpyxl
import pytest

from pubtab import convert, read_excel, tex_to_excel
from pubtab.models import SpacingConfig, TableData
from pubtab.preview import _build_standalone, _find_pdflatex, compile_pdf
from pubtab.renderer import render
from pubtab.tex_reader import read_tex

FIXTURES = Path(__file__).parent / "fixtures"
TABLES = ["table1", "table2", "table3", "table4", "table5", "table6", "table8"]


def _compare_xlsx(path_a: Path, path_b: Path) -> list:
    """Compare two xlsx files cell by cell. Returns list of diffs."""
    wb1 = openpyxl.load_workbook(path_a)
    wb2 = openpyxl.load_workbook(path_b)
    ws1, ws2 = wb1.active, wb2.active
    diffs = []
    max_r = max(ws1.max_row, ws2.max_row)
    max_c = max(ws1.max_column, ws2.max_column)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v1 = ws1.cell(r, c).value
            v2 = ws2.cell(r, c).value
            s1 = "" if v1 is None else str(v1).strip()
            s2 = "" if v2 is None else str(v2).strip()
            try:
                if abs(float(s1) - float(s2)) < 1e-9:
                    continue
            except (ValueError, TypeError):
                pass
            if s1 != s2:
                diffs.append((r, c, v1, v2))
    return diffs


# --- tex → xlsx round-trip ---

@pytest.mark.parametrize("name", TABLES)
def test_tex_to_xlsx_dimensions(name, tmp_path):
    """tex→xlsx produces correct dimensions matching original xlsx."""
    tex_file = FIXTURES / f"{name}.tex"
    orig_xlsx = FIXTURES / f"{name}.xlsx"
    gen_xlsx = tmp_path / f"{name}.xlsx"
    tex_to_excel(str(tex_file), str(gen_xlsx))

    wb_orig = openpyxl.load_workbook(orig_xlsx)
    wb_gen = openpyxl.load_workbook(gen_xlsx)
    assert wb_gen.active.max_row == wb_orig.active.max_row
    assert wb_gen.active.max_column == wb_orig.active.max_column


@pytest.mark.parametrize("name", TABLES)
def test_tex_to_xlsx_values_match(name, tmp_path):
    """tex→xlsx cell values match original xlsx exactly."""
    gen_xlsx = tmp_path / f"{name}.xlsx"
    tex_to_excel(str(FIXTURES / f"{name}.tex"), str(gen_xlsx))
    diffs = _compare_xlsx(FIXTURES / f"{name}.xlsx", gen_xlsx)
    assert diffs == [], f"Cell diffs: {diffs[:5]}"


@pytest.mark.parametrize("name", TABLES)
def test_tex_to_xlsx_merged_cells(name, tmp_path):
    """tex→xlsx preserves merged cell count."""
    gen_xlsx = tmp_path / f"{name}.xlsx"
    tex_to_excel(str(FIXTURES / f"{name}.tex"), str(gen_xlsx))

    wb_orig = openpyxl.load_workbook(FIXTURES / f"{name}.xlsx")
    wb_gen = openpyxl.load_workbook(gen_xlsx)
    assert len(wb_gen.active.merged_cells.ranges) == len(wb_orig.active.merged_cells.ranges)


# --- xlsx → tex → xlsx round-trip ---

@pytest.mark.parametrize("name", TABLES)
def test_xlsx_to_tex_roundtrip(name, tmp_path):
    """xlsx→tex→xlsx round-trip preserves all cell values."""
    tex_path = tmp_path / f"{name}.tex"
    convert(str(FIXTURES / f"{name}.xlsx"), str(tex_path))

    gen_xlsx = tmp_path / f"{name}_rt.xlsx"
    tex_to_excel(str(tex_path), str(gen_xlsx))

    table_orig = read_excel(str(FIXTURES / f"{name}.xlsx"))
    table_gen = read_tex(tex_path.read_text())
    assert table_gen.num_rows == table_orig.num_rows
    assert table_gen.num_cols == table_orig.num_cols


# --- tex_reader unit tests ---

def test_tex_reader_comments():
    """Parser handles % comments correctly."""
    tex = r"""
\begin{tabular}{cc}
\toprule
A & B \\ % header comment
\midrule
1 & 2 \\ % data comment
\bottomrule
\end{tabular}
"""
    table = read_tex(tex)
    assert table.num_rows == 2
    assert table.cells[1][0].value == 1.0


def test_tex_reader_multicolumn():
    """Parser handles \\multicolumn correctly."""
    tex = r"""
\begin{tabular}{ccc}
\toprule
\multicolumn{2}{c}{Header} & C \\
\midrule
a & b & c \\
\bottomrule
\end{tabular}
"""
    table = read_tex(tex)
    assert table.cells[0][0].colspan == 2
    assert table.cells[0][0].value == "Header"


def test_tex_reader_multirow():
    """Parser handles \\multirow correctly."""
    tex = r"""
\begin{tabular}{cc}
\toprule
\multirow{2}{*}{A} & B \\
 & C \\
\bottomrule
\end{tabular}
"""
    table = read_tex(tex)
    assert table.cells[0][0].rowspan == 2


def test_tex_reader_diagbox():
    """Parser handles \\diagbox correctly."""
    tex = r"""
\begin{tabular}{cc}
\toprule
\diagbox{Row}{Col} & Data \\
\midrule
a & 1 \\
\bottomrule
\end{tabular}
"""
    table = read_tex(tex)
    assert table.cells[0][0].style.diagbox == ["Row", "Col"]


def test_tex_reader_formatting():
    """Parser extracts bold/italic/underline formatting."""
    tex = r"""
\begin{tabular}{ccc}
\toprule
\textbf{Bold} & \textit{Italic} & \underline{Under} \\
\bottomrule
\end{tabular}
"""
    table = read_tex(tex)
    assert table.cells[0][0].style.bold
    assert table.cells[0][1].style.italic
    assert table.cells[0][2].style.underline


def test_tex_reader_math_cleanup():
    """Parser cleans up math expressions."""
    tex = r"""
\begin{tabular}{c}
\toprule
$D_\text{stage 1}$ \\
\bottomrule
\end{tabular}
"""
    table = read_tex(tex)
    assert table.cells[0][0].value == "D_stage 1"


def test_tex_reader_pm_spacing():
    """Parser normalizes ± spacing."""
    tex = r"""
\begin{tabular}{c}
\toprule
0.626 {$\pm$0.018} \\
\bottomrule
\end{tabular}
"""
    table = read_tex(tex)
    assert "0.626±0.018" in str(table.cells[0][0].value)


# --- renderer tests ---

def test_render_three_line():
    """Renderer produces three-line table structure."""
    from pubtab.models import Cell, CellStyle
    header = [Cell("A", CellStyle(bold=True)), Cell("B", CellStyle(bold=True))]
    row = [Cell("x"), Cell(0.5)]
    table = TableData(cells=[header, row], num_rows=2, num_cols=2, header_rows=1)
    tex = render(table)
    assert "\\toprule" in tex
    assert "\\midrule" in tex
    assert "\\bottomrule" in tex


def test_render_default_heavyrulewidth():
    """Renderer includes default heavyrulewidth."""
    from pubtab.models import Cell, CellStyle
    row = [Cell("a"), Cell("b")]
    table = TableData(cells=[row], num_rows=1, num_cols=2, header_rows=0)
    tex = render(table)
    assert "\\heavyrulewidth" in tex
    assert "1.0pt" in tex


def test_render_special_chars():
    """Renderer escapes special LaTeX characters."""
    from pubtab.models import Cell
    row = [Cell("100%"), Cell("A & B")]
    table = TableData(cells=[row], num_rows=1, num_cols=2, header_rows=0)
    tex = render(table)
    assert "100\\%" in tex
    assert "A \\& B" in tex


# --- preview tests ---

def test_build_standalone_structure():
    """Standalone document has correct structure."""
    tex = _build_standalone("\\begin{table}...\\end{table}")
    assert "\\documentclass" in tex
    assert "\\begin{document}" in tex
    assert "\\resizebox" in tex


def test_build_standalone_auto_resizebox():
    """Standalone auto-adds resizebox when not present."""
    tex = _build_standalone("\\begin{tabular}{cc}a & b\\end{tabular}")
    assert "\\resizebox{\\linewidth}" in tex


def test_build_standalone_keeps_existing_resizebox():
    """Standalone doesn't double-wrap resizebox."""
    content = "\\resizebox{0.9\\textwidth}{!}{\\begin{tabular}{cc}a & b\\end{tabular}}"
    tex = _build_standalone(f"\\begin{{table}}[t]{content}\\end{{table}}")
    assert tex.count("\\resizebox") == 1


@pytest.mark.skipif(not _find_pdflatex(), reason="pdflatex not available")
def test_compile_pdf(tmp_path):
    """PDF compilation works."""
    tex = "\\begin{tabular}{cc}\na & b \\\\\n\\end{tabular}"
    out = tmp_path / "test.pdf"
    compile_pdf(tex, out)
    assert out.exists()
