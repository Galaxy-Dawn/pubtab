"""LaTeX compilation and preview generation with auto-install."""

from __future__ import annotations

import logging
import platform
import shutil
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path
from typing import Optional, Union

from .themes import load_theme

logger = logging.getLogger(__name__)

PUBTAB_HOME = Path.home() / ".pubtab"
TEXLIVE_DIR = PUBTAB_HOME / "TinyTeX"

_REQUIRED_PACKAGES = ["booktabs", "multirow", "xcolor", "standalone",
                      "adjustbox", "collectbox", "currfile", "gincltex"]


def _get_tinytex_bin_dir() -> Optional[Path]:
    """Find the TinyTeX bin directory under PUBTAB_HOME."""
    if not TEXLIVE_DIR.exists():
        return None
    # Search for bin/*/pdflatex
    for p in TEXLIVE_DIR.glob("bin/*/pdflatex"):
        return p.parent
    return None


def _find_pdflatex() -> Optional[str]:
    """Find pdflatex: system PATH first, then pubtab's TinyTeX."""
    system = shutil.which("pdflatex")
    if system:
        return system
    bin_dir = _get_tinytex_bin_dir()
    if bin_dir and (bin_dir / "pdflatex").exists():
        return str(bin_dir / "pdflatex")
    return None


def _install_tinytex() -> str:
    """Download and install TinyTeX into ~/.pubtab/TinyTeX.

    Returns:
        Path to pdflatex binary.

    Raises:
        RuntimeError: If installation fails.
    """
    PUBTAB_HOME.mkdir(parents=True, exist_ok=True)
    system = platform.system()
    machine = platform.machine()

    if system == "Darwin":
        if machine == "arm64":
            url = "https://github.com/rstudio/tinytex-releases/releases/download/v2025.03/TinyTeX-1-v2025.03.tgz"
        else:
            url = "https://github.com/rstudio/tinytex-releases/releases/download/v2025.03/TinyTeX-1-v2025.03.tgz"
        archive = PUBTAB_HOME / "tinytex.tgz"
    elif system == "Linux":
        url = "https://github.com/rstudio/tinytex-releases/releases/download/v2025.03/TinyTeX-1-v2025.03.tar.gz"
        archive = PUBTAB_HOME / "tinytex.tar.gz"
    elif system == "Windows":
        url = "https://github.com/rstudio/tinytex-releases/releases/download/v2025.03/TinyTeX-1-v2025.03.zip"
        archive = PUBTAB_HOME / "tinytex.zip"
    else:
        raise RuntimeError(f"Unsupported platform: {system}")

    # Download
    logger.info("Downloading TinyTeX (this only happens once)...")
    print("pubtab: Downloading TinyTeX (~90MB, one-time setup)...", file=sys.stderr)
    urllib.request.urlretrieve(url, archive)

    # Extract
    print("pubtab: Extracting TinyTeX...", file=sys.stderr)
    if TEXLIVE_DIR.exists():
        shutil.rmtree(TEXLIVE_DIR)

    if system == "Windows":
        import zipfile
        with zipfile.ZipFile(archive) as zf:
            zf.extractall(PUBTAB_HOME)
    else:
        subprocess.run(
            ["tar", "-xzf", str(archive), "-C", str(PUBTAB_HOME)],
            check=True, capture_output=True,
        )

    archive.unlink(missing_ok=True)

    # Find pdflatex
    bin_dir = _get_tinytex_bin_dir()
    if not bin_dir:
        raise RuntimeError("TinyTeX installation failed: pdflatex not found after extraction.")

    pdflatex = str(bin_dir / "pdflatex")

    # Install required LaTeX packages
    tlmgr = str(bin_dir / "tlmgr")
    if Path(tlmgr).exists():
        print("pubtab: Installing LaTeX packages...", file=sys.stderr)
        subprocess.run(
            [tlmgr, "install"] + _REQUIRED_PACKAGES,
            capture_output=True,
        )

    print("pubtab: TinyTeX setup complete!", file=sys.stderr)
    return pdflatex


def ensure_pdflatex() -> str:
    """Return path to pdflatex, auto-installing TinyTeX if needed.

    Returns:
        Path to pdflatex binary.
    """
    path = _find_pdflatex()
    if path:
        return path
    return _install_tinytex()


def _strip_table_float(tex: str) -> str:
    """Strip \\begin{table}...\\end{table} float wrapper, keep inner content."""
    import re
    tex = re.sub(r"\\begin\{table\*?\}\[.*?\]\s*", "", tex)
    tex = re.sub(r"\\end\{table\*?\}\s*", "", tex)
    tex = re.sub(r"\\centering\s*", "", tex)
    tex = re.sub(r"\\caption\{", r"\\captionof{table}{", tex)
    return tex.strip()


def _build_standalone(tex_content: str, theme: str = "three_line",
                      preamble: Optional[str] = None) -> str:
    """Wrap table LaTeX in a standalone document."""
    config, _ = load_theme(theme)
    all_pkgs = list(config.packages) + ["caption", "graphicx"]
    pkgs = "\n".join(f"\\usepackage{{{p}}}" for p in all_pkgs)
    extra = f"\n{preamble}" if preamble else ""
    inner = _strip_table_float(tex_content)
    # Auto-wrap in resizebox for preview if not already present
    if "\\resizebox" not in inner:
        inner = f"\\resizebox{{\\linewidth}}{{!}}{{{inner}}}"
    return (
        "\\documentclass[border=10pt]{standalone}\n"
        f"{pkgs}{extra}\n"
        "\\setlength{\\textwidth}{24cm}\n"
        "\\begin{document}\n"
        "\\begin{minipage}{24cm}\n"
        "\\centering\n"
        f"{inner}\n"
        "\\end{minipage}\n"
        "\\end{document}\n"
    )


def compile_pdf(
    tex_content: str,
    output: Union[str, Path],
    theme: str = "three_line",
    preamble: Optional[str] = None,
) -> Path:
    """Compile LaTeX content to PDF.

    Args:
        tex_content: Raw LaTeX table code.
        output: Output PDF path.
        theme: Theme name (for package imports).
        preamble: Extra LaTeX preamble (e.g. custom commands).

    Returns:
        Path to generated PDF.

    Raises:
        RuntimeError: If compilation fails.
    """
    pdflatex = ensure_pdflatex()
    doc = _build_standalone(tex_content, theme, preamble=preamble)
    output = Path(output)

    with tempfile.TemporaryDirectory() as tmpdir:
        tex_path = Path(tmpdir) / "table.tex"
        tex_path.write_text(doc)

        result = subprocess.run(
            [pdflatex, "-interaction=nonstopmode", "-output-directory", tmpdir, str(tex_path)],
            capture_output=True,
            text=True,
        )
        pdf_path = Path(tmpdir) / "table.pdf"
        if not pdf_path.exists():
            raise RuntimeError(f"pdflatex failed:\n{result.stdout}\n{result.stderr}")

        shutil.copy2(pdf_path, output)
    return output


def preview(
    tex_input: Union[str, Path],
    output: Optional[Union[str, Path]] = None,
    theme: str = "three_line",
    dpi: int = 300,
    preamble: Optional[str] = None,
) -> Path:
    """Generate a PNG preview from LaTeX content or .tex file.

    Args:
        tex_input: LaTeX string or path to .tex file.
        output: Output PNG path. Defaults to input stem + .png.
        theme: Theme name.
        dpi: Resolution for PNG output.

    Returns:
        Path to generated PNG.
    """
    tex_path = Path(tex_input)
    if tex_path.exists():
        tex_content = tex_path.read_text()
        if output is None:
            output = tex_path.with_suffix(".png")
    else:
        tex_content = str(tex_input)
        if output is None:
            output = Path("preview.png")

    output = Path(output)

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / "table.pdf"
        compile_pdf(tex_content, pdf_path, theme=theme, preamble=preamble)
        _pdf_to_png(pdf_path, output, dpi=dpi)

    return output


def _pdf_to_png(pdf_path: Path, output: Path, dpi: int = 300) -> None:
    """Convert PDF to PNG."""
    try:
        from pdf2image import convert_from_path

        images = convert_from_path(str(pdf_path), dpi=dpi)
        if images:
            images[0].save(str(output))
            return
    except ImportError:
        pass

    # Fallback: qlmanage (macOS, native high-quality), convert (ImageMagick), sips (last resort)
    if shutil.which("qlmanage"):
        # qlmanage renders PDF natively at target pixel size
        with tempfile.TemporaryDirectory() as tmpdir:
            # Calculate target pixel width from PDF points and DPI
            target_size = max(dpi * 6, 1800)  # at least 1800px wide
            subprocess.run(
                ["qlmanage", "-t", "-s", str(target_size), "-o", tmpdir, str(pdf_path)],
                capture_output=True,
            )
            # qlmanage outputs as <filename>.png in the output dir
            pngs = list(Path(tmpdir).glob("*.png"))
            if pngs:
                shutil.copy2(pngs[0], output)
                return
    elif shutil.which("convert"):
        subprocess.run(
            ["convert", "-density", str(dpi), str(pdf_path), str(output)],
            capture_output=True,
        )
    else:
        raise RuntimeError(
            "No PDF-to-PNG converter found. Install pdf2image (`pip install pubtab[preview]`) "
            "or ImageMagick."
        )
