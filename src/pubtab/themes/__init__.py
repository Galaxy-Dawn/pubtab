"""Theme registry and loader."""

from __future__ import annotations

from pathlib import Path
from typing import Dict

import yaml

from ..models import SpacingConfig, ThemeConfig

_THEMES_DIR = Path(__file__).parent
_THEME_CACHE: Dict[str, tuple[ThemeConfig, str]] = {}


def list_themes() -> list[str]:
    """Return names of all available themes."""
    return [d.name for d in _THEMES_DIR.iterdir() if d.is_dir() and (d / "config.yaml").exists()]


def load_theme(name: str) -> tuple[ThemeConfig, str]:
    """Load a theme's config and template.

    Returns:
        Tuple of (ThemeConfig, template_string).
    """
    if name in _THEME_CACHE:
        return _THEME_CACHE[name]

    theme_dir = _THEMES_DIR / name
    if not theme_dir.exists():
        raise ValueError(f"Theme '{name}' not found. Available: {list_themes()}")

    with open(theme_dir / "config.yaml") as f:
        raw = yaml.safe_load(f)

    sp_raw = raw.get("spacing", {}) or {}
    spacing = SpacingConfig(
        tabcolsep=sp_raw.get("tabcolsep"),
        arraystretch=sp_raw.get("arraystretch"),
        heavyrulewidth=sp_raw.get("heavyrulewidth"),
        lightrulewidth=sp_raw.get("lightrulewidth"),
        arrayrulewidth=sp_raw.get("arrayrulewidth"),
    )

    config = ThemeConfig(
        name=raw["name"],
        description=raw.get("description", ""),
        packages=raw.get("packages", []),
        column_sep=raw.get("column_sep", "1em"),
        font_size=raw.get("font_size"),
        caption_position=raw.get("caption_position", "top"),
        spacing=spacing,
    )

    template_str = (theme_dir / "template.tex").read_text()
    _THEME_CACHE[name] = (config, template_str)
    return config, template_str
