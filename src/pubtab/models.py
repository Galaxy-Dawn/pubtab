"""Data models for pubtab."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Union


@dataclass(frozen=True)
class CellStyle:
    """Style attributes for a table cell."""

    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: Optional[str] = None
    bg_color: Optional[str] = None
    alignment: str = "center"
    fmt: Optional[str] = None
    strip_leading_zero: bool = True  # 0.451 → .451
    raw_latex: bool = False  # skip escaping, pass value as-is
    diagbox: Optional[List[str]] = None  # e.g. ["Models", "Datasets"]


@dataclass(frozen=True)
class Cell:
    """A single table cell with value, style, and span info."""

    value: Any = ""
    style: CellStyle = field(default_factory=CellStyle)
    rowspan: int = 1
    colspan: int = 1


@dataclass(frozen=True)
class TableData:
    """2D grid of cells representing a table."""

    cells: List[List[Cell]]
    num_rows: int
    num_cols: int
    header_rows: int = 1
    group_separators: Union[Dict[int, Union[str, List[str]]], List[int]] = field(default_factory=dict)
    # List[int]: row indices where \midrule is inserted AFTER that row.
    # Dict form still supported for custom separators per row.


@dataclass(frozen=True)
class SpacingConfig:
    """Table spacing and rule controls."""

    tabcolsep: Optional[str] = None       # e.g. "4pt"
    arraystretch: Optional[str] = None    # e.g. "1.0"
    heavyrulewidth: Optional[str] = "1.0pt"
    lightrulewidth: Optional[str] = "0.5pt"
    arrayrulewidth: Optional[str] = "0.5pt"


@dataclass(frozen=True)
class ThemeConfig:
    """Configuration for a rendering theme."""

    name: str
    description: str = ""
    packages: List[str] = field(default_factory=list)
    column_sep: str = "1em"
    font_size: Optional[str] = None
    caption_position: str = "top"
    spacing: SpacingConfig = field(default_factory=SpacingConfig)
