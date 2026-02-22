"""Excel writer — converts TableData to .xlsx."""

from __future__ import annotations

from pathlib import Path
from typing import Union

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .models import TableData

# Standard LaTeX/xcolor named colors → RGB hex
_LATEX_COLORS = {
    "red": "FF0000", "blue": "0000FF", "green": "008000", "black": "000000",
    "white": "FFFFFF", "gray": "808080", "grey": "808080", "cyan": "00FFFF",
    "magenta": "FF00FF", "yellow": "FFFF00", "orange": "FF8000",
    "purple": "800080", "brown": "804000", "violet": "8000FF",
    "pink": "FFC0CB", "lime": "00FF00", "olive": "808000", "teal": "008080",
    "darkgray": "404040", "lightgray": "C0C0C0",
}


def _latex_color_to_hex(color: str) -> str | None:
    """Convert LaTeX color spec to 6-digit RGB hex (no '#' prefix).

    Supports: named colors, xcolor mixing (e.g. 'gray!20'), hex codes.
    """
    color = color.strip()
    if not color:
        return None
    # Already hex: #RRGGBB or RRGGBB
    if color.startswith("#"):
        return color[1:].upper()
    if len(color) == 6 and all(c in "0123456789abcdefABCDEF" for c in color):
        return color.upper()
    # xcolor mixing: "color!percent" (e.g. gray!20 = 20% gray + 80% white)
    if "!" in color:
        parts = color.split("!")
        base = _LATEX_COLORS.get(parts[0].lower())
        if base and len(parts) >= 2:
            try:
                pct = float(parts[1]) / 100.0
            except ValueError:
                return None
            br, bg, bb = int(base[0:2], 16), int(base[2:4], 16), int(base[4:6], 16)
            r = int(br * pct + 255 * (1 - pct))
            g = int(bg * pct + 255 * (1 - pct))
            b = int(bb * pct + 255 * (1 - pct))
            return f"{r:02X}{g:02X}{b:02X}"
        return None
    # Named color
    return _LATEX_COLORS.get(color.lower())


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

            # Font styling (with optional text color)
            font_kwargs = dict(
                bold=cell.style.bold,
                italic=cell.style.italic,
                underline="single" if cell.style.underline else None,
            )
            if cell.style.color:
                hex_color = _latex_color_to_hex(cell.style.color)
                if hex_color:
                    font_kwargs["color"] = hex_color
            ws.cell(row=excel_row, column=excel_col).font = Font(**font_kwargs)

            # Background color
            if cell.style.bg_color:
                hex_bg = _latex_color_to_hex(cell.style.bg_color)
                if hex_bg:
                    ws.cell(row=excel_row, column=excel_col).fill = PatternFill(
                        start_color=hex_bg, end_color=hex_bg, fill_type="solid",
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
