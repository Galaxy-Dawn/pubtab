"""Theme registry and loader."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Optional

import yaml

from ..models import SpacingConfig, ThemeConfig

_THEMES_DIR = Path(__file__).parent
_THEME_CACHE: Dict[tuple[str, str], tuple[ThemeConfig, str]] = {}
_BACKEND_VARIANT_SUFFIXES = {"tabularray": "_tabularray"}


def list_themes() -> list[str]:
    """Return names of all available themes."""
    names: set[str] = set()
    for theme_dir in _THEMES_DIR.iterdir():
        if not theme_dir.is_dir() or not (theme_dir / "config.yaml").exists():
            continue
        name = theme_dir.name
        for suffix in _BACKEND_VARIANT_SUFFIXES.values():
            if name.endswith(suffix):
                name = name[: -len(suffix)]
                break
        names.add(name)
    return sorted(names)


def normalize_theme_backend(name: str, backend: Optional[str] = None) -> tuple[str, str]:
    """Normalize a user-facing style-theme/backend pair.

    Supports deprecated backend-specific theme aliases such as
    ``three_line_tabularray`` while keeping ``theme`` and ``backend`` conceptually
    separate for new callers.
    """
    inferred_backend = "tabular"
    canonical_name = name
    for candidate_backend, suffix in _BACKEND_VARIANT_SUFFIXES.items():
        if name.endswith(suffix):
            canonical_name = name[: -len(suffix)]
            inferred_backend = candidate_backend
            break

    if backend is not None and inferred_backend != "tabular" and backend != inferred_backend:
        raise ValueError(
            f"Theme alias '{name}' implies backend '{inferred_backend}', "
            f"but backend '{backend}' was requested."
        )

    resolved_backend = backend or inferred_backend
    if resolved_backend not in {"tabular", "tabularray"}:
        raise ValueError(f"Unsupported LaTeX backend: {resolved_backend}")
    return canonical_name, resolved_backend


def _resolve_theme_dir(name: str, backend: str) -> Path:
    """Resolve a canonical style-theme/backend pair to a concrete theme directory."""
    canonical_name, resolved_backend = normalize_theme_backend(name, backend)

    if resolved_backend == "tabular":
        theme_dir = _THEMES_DIR / canonical_name
        if theme_dir.exists():
            return theme_dir
        raise ValueError(f"Theme '{canonical_name}' not found. Available: {list_themes()}")

    suffix = _BACKEND_VARIANT_SUFFIXES.get(resolved_backend)
    if suffix is None:
        raise ValueError(f"Unsupported LaTeX backend: {resolved_backend}")

    variant_dir = _THEMES_DIR / f"{canonical_name}{suffix}"
    if variant_dir.exists():
        return variant_dir
    raise ValueError(
        f"Theme '{canonical_name}' does not provide a '{resolved_backend}' backend."
    )


def resolve_theme(name: str, backend: str = "tabular") -> str:
    """Resolve a user-facing style-theme/backend pair to an actual theme directory name."""
    return _resolve_theme_dir(name, backend).name


def load_theme(name: str, backend: Optional[str] = None) -> tuple[ThemeConfig, str]:
    """Load a style theme's config and template for a specific backend.

    Returns:
        Tuple of (ThemeConfig, template_string).
    """
    canonical_name, resolved_backend = normalize_theme_backend(name, backend)
    cache_key = (canonical_name, resolved_backend)
    if cache_key in _THEME_CACHE:
        return _THEME_CACHE[cache_key]

    theme_dir = _resolve_theme_dir(canonical_name, resolved_backend)

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
        backend=raw.get("backend", "tabular"),
        packages=raw.get("packages", []),
        preamble_hints=raw.get("preamble_hints", []),
        column_sep=raw.get("column_sep", "1em"),
        font_size=raw.get("font_size"),
        caption_position=raw.get("caption_position", "top"),
        spacing=spacing,
    )

    template_str = (theme_dir / "template.tex").read_text()
    _THEME_CACHE[cache_key] = (config, template_str)
    return config, template_str
