"""Excel writer — converts TableData to .xlsx."""

from __future__ import annotations

from pathlib import Path
from typing import Union

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from .models import TableData


def write_excel(table: TableData, output: Union[str, Path]) -> Path:
    """Write TableData to an Excel file.

    Args:
        table: The table data to write.
        output: Output .xlsx file path.

    Returns:
        Path to the generated file.
    """
    output = Path(output)
    wb = Workbook()
    ws = wb.active

    # Track merged regions to skip placeholder cells
    merged: set[tuple[int, int]] = set()

    for r, row in enumerate(table.cells):
        col_offset = 0
        for c, cell in enumerate(row):
            excel_row = r + 1
            excel_col = c + 1

            if (r, c) in merged:
                continue

            # Determine value
            if cell.style.diagbox:
                val = " / ".join(cell.style.diagbox)
            else:
                val = cell.value if cell.value != "" else None

            ws.cell(row=excel_row, column=excel_col, value=val)

            # Font styling
            ws.cell(row=excel_row, column=excel_col).font = Font(
                bold=cell.style.bold,
                italic=cell.style.italic,
                underline="single" if cell.style.underline else None,
            )

            # Alignment
            ws.cell(row=excel_row, column=excel_col).alignment = Alignment(
                horizontal=_align_map(cell.style.alignment),
                vertical="center",
            )

            # Merge cells
            if cell.rowspan > 1 or cell.colspan > 1:
                end_row = excel_row + cell.rowspan - 1
                end_col = excel_col + cell.colspan - 1
                ws.merge_cells(
                    start_row=excel_row, start_column=excel_col,
                    end_row=end_row, end_column=end_col,
                )
                for mr in range(r, r + cell.rowspan):
                    for mc in range(c, c + cell.colspan):
                        if (mr, mc) != (r, c):
                            merged.add((mr, mc))

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output))
    return output


def _align_map(latex_align: str) -> str:
    """Map LaTeX alignment to Excel alignment."""
    a = latex_align[0] if latex_align else "c"
    return {"l": "left", "r": "right", "c": "center"}.get(a, "center")
