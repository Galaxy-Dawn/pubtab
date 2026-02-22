"""Excel writer — converts TableData to .xlsx."""

from __future__ import annotations

from pathlib import Path
from typing import Union

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

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
                if cell.value == "" or cell.value is None:
                    continue
                # Cell has content but overlaps a previous merge — remove from merged
                merged.discard((r, c))

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

            # Alignment (wrap_text for multi-line cells, rotation for rotatebox)
            wrap = isinstance(val, str) and "\n" in val
            ws.cell(row=excel_row, column=excel_col).alignment = Alignment(
                horizontal=_align_map(cell.style.alignment),
                vertical="center",
                wrap_text=wrap,
                text_rotation=cell.style.rotation if cell.style.rotation else 0,
            )

            # Merge cells (cap rowspan to avoid overlapping with next content cell)
            if cell.rowspan > 1 or cell.colspan > 1:
                actual_rowspan = min(cell.rowspan, len(table.cells) - r)
                if actual_rowspan > 1:
                    for dr in range(1, actual_rowspan):
                        nr = r + dr
                        if nr < len(table.cells) and table.cells[nr][c].value not in ("", None):
                            actual_rowspan = dr
                            break
                end_row = excel_row + actual_rowspan - 1
                end_col = excel_col + cell.colspan - 1
                if end_row > excel_row or end_col > excel_col:
                    ws.merge_cells(
                        start_row=excel_row, start_column=excel_col,
                        end_row=end_row, end_column=end_col,
                    )
                for mr in range(r, r + actual_rowspan):
                    for mc in range(c, c + cell.colspan):
                        if (mr, mc) != (r, c):
                            merged.add((mr, mc))

    # Auto-adjust column widths based on content
    for col_idx in range(1, table.num_cols + 1):
        max_len = 0
        for row_idx in range(1, len(table.cells) + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is not None:
                # Use max line length for multi-line cells
                cell_len = max(len(str(line)) for line in str(val).split("\n"))
                max_len = max(max_len, cell_len)
        # Add padding, minimum width 4
        ws.column_dimensions[get_column_letter(col_idx)].width = max(max_len + 3, 4)

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output))
    return output


def _align_map(latex_align: str) -> str:
    """Map LaTeX alignment to Excel alignment."""
    a = latex_align[0] if latex_align else "c"
    return {"l": "left", "r": "right", "c": "center"}.get(a, "center")
