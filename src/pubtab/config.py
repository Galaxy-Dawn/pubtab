"""YAML config loader for pubtab."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Tuple, Union

import yaml

from .models import SpacingConfig


def load_config(path: Union[str, Path]) -> Tuple[Dict[str, Any], None]:
    """Load a YAML config file and return (convert_kwargs, None).

    Returns:
        Tuple of (kwargs dict for convert(), None).
    """
    path = Path(path)
    with open(path) as f:
        cfg = yaml.safe_load(f)

    kwargs: Dict[str, Any] = {}

    # Simple pass-through keys
    for key in (
        "theme", "caption", "label", "position",
        "font_size", "resizebox", "col_spec", "header_sep",
        "num_cols", "preview", "preamble", "dpi",
        "header_rows", "sheet",
    ):
        if key in cfg:
            kwargs[key] = cfg[key]

    if "span_columns" in cfg:
        kwargs["span_columns"] = cfg["span_columns"]

    if "spacing" in cfg:
        kwargs["spacing"] = SpacingConfig(**cfg["spacing"])

    if "group_separators" in cfg:
        gs = cfg["group_separators"]
        if isinstance(gs, list):
            kwargs["group_separators"] = gs
        else:
            kwargs["group_separators"] = {
                int(k): v for k, v in gs.items()
            }

    return kwargs, None
