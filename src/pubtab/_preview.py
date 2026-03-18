"""LaTeX compilation and preview generation with auto-install."""

from __future__ import annotations

import logging
import platform
import re
import ssl
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from pathlib import Path
from typing import Optional, Union

from .themes import load_theme, normalize_theme_backend

logger = logging.getLogger(__name__)

PUBTAB_HOME = Path.home() / ".pubtab"
TEXLIVE_DIR = PUBTAB_HOME / "TinyTeX"

_REQUIRED_PACKAGES = ["booktabs", "multirow", "xcolor", "standalone",
                      "adjustbox", "collectbox", "currfile", "gincltex",
                      "amsfonts", "pifont"]


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


def _build_download_ssl_context() -> Optional[ssl.SSLContext]:
    """Build an SSL context with certifi when available."""
    try:
        import certifi  # type: ignore
    except Exception:
        return None
    return ssl.create_default_context(cafile=certifi.where())


def _download_archive(url: str, archive: Path) -> None:
    """Download TinyTeX archive with cert-friendly SSL handling."""
    req = urllib.request.Request(url, headers={"User-Agent": "pubtab/1.x"})
    context = _build_download_ssl_context()
    try:
        with urllib.request.urlopen(req, context=context) as resp, archive.open("wb") as f:
            shutil.copyfileobj(resp, f)
    except urllib.error.URLError as exc:
        if isinstance(exc.reason, ssl.SSLCertVerificationError):
            raise RuntimeError(
                "TinyTeX download failed due to SSL certificate verification. "
                "Try one of: (1) run '/Applications/Python 3.11/Install Certificates.command' on macOS, "
                "(2) install certifi via 'python3 -m pip install -U certifi', "
                "(3) export SSL_CERT_FILE to certifi.where(), or "
                "(4) install TinyTeX/TeX Live manually and ensure 'pdflatex' is in PATH."
            ) from exc
        raise RuntimeError(
            f"TinyTeX download failed from {url}. Please check network/proxy settings or install TinyTeX manually."
        ) from exc


def _extract_missing_sty(log_text: str) -> Optional[str]:
    """Extract missing .sty package name from pdflatex output."""
    m = re.search(r"LaTeX Error: File `([^`]+)\.sty' not found\.", log_text)
    if not m:
        return None
    return m.group(1).strip()


def _sty_to_tlmgr_package(sty_name: str) -> Optional[str]:
    """Map .sty filename to tlmgr package name."""
    mapping = {
        "pifont": "psnfss",
        "graphicx": "graphics",
    }
    if sty_name in mapping:
        return mapping[sty_name]
    return sty_name


def _find_tlmgr() -> Optional[str]:
    """Find tlmgr command from PATH or pubtab-managed TinyTeX."""
    tlmgr = shutil.which("tlmgr")
    if tlmgr:
        return tlmgr
    bin_dir = _get_tinytex_bin_dir()
    if bin_dir and (bin_dir / "tlmgr").exists():
        return str(bin_dir / "tlmgr")
    return None


def _tlmgr_install_package(package: str) -> None:
    """Install a TeX Live package via tlmgr."""
    tlmgr = _find_tlmgr()
    if not tlmgr:
        raise RuntimeError(
            f"Missing LaTeX package '{package}', but tlmgr is not available. "
            "Please install it manually or install a full TeX distribution."
        )
    subprocess.run([tlmgr, "install", package], check=True, capture_output=True)


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
    _download_archive(url, archive)

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
    tex = re.sub(r"\\begin\{table\*?\}(?:\[[^\]]*\])?\s*", "", tex)
    tex = re.sub(r"\\end\{table\*?\}\s*", "", tex)
    tex = re.sub(r"\\centering\s*", "", tex)
    tex = re.sub(r"\\caption\{", r"\\captionof{table}{", tex)
    return tex.strip()


def _split_leading_setup(inner: str) -> tuple[str, str]:
    """Split non-typesetting setup commands from the actual table body.

    Preview wraps the rendered table in ``\\resizebox``. Commands like
    ``\\definecolor`` or ``\\setlength`` must stay outside that box, otherwise
    tabularray previews with many color definitions can render at the wrong
    scale or with broken merged-cell layout.
    """
    setup_prefixes = (
        "\\definecolor",
        "\\setlength",
        "\\renewcommand",
        "\\SetTblrInner",
        "\\arrayrulecolor",
        "\\rowcolors",
    )

    setup_lines: list[str] = []
    body_lines: list[str] = []
    body_started = False

    for line in inner.splitlines():
        stripped = line.strip()
        is_setup = (
            not stripped
            or stripped.startswith("%")
            or stripped.startswith(setup_prefixes)
        )
        if not body_started and is_setup:
            setup_lines.append(line)
            continue
        body_started = True
        body_lines.append(line)

    setup = "\n".join(setup_lines).strip()
    body = "\n".join(body_lines).strip()
    return setup, body


def _resolve_preview_theme_backend(
    tex_content: str,
    theme: str,
    latex_backend: Optional[str] = None,
) -> tuple[str, str]:
    """Resolve preview theme/backend, inferring tabularray from tblr content."""
    canonical_theme, resolved_backend = normalize_theme_backend(theme, latex_backend)
    if latex_backend is None and "\\begin{tblr}" in tex_content:
        resolved_backend = "tabularray"
    return canonical_theme, resolved_backend


def _build_standalone(
    tex_content: str,
    theme: str = "three_line",
    latex_backend: Optional[str] = None,
    preamble: Optional[str] = None,
) -> str:
    """Wrap table LaTeX in a standalone document."""
    canonical_theme, resolved_backend = _resolve_preview_theme_backend(
        tex_content, theme, latex_backend
    )
    config, _ = load_theme(canonical_theme, backend=resolved_backend)
    all_pkgs = list(config.packages) + ["caption", "graphicx", "fontenc"]
    pkg_lines = []
    for p in all_pkgs:
        if p == "xcolor":
            pkg_lines.append("\\usepackage[dvipsnames,table]{xcolor}")
        elif p == "fontenc":
            pkg_lines.append("\\usepackage[T1]{fontenc}")
        else:
            pkg_lines.append(f"\\usepackage{{{p}}}")
    pkg_lines.extend(config.preamble_hints)
    pkgs = "\n".join(pkg_lines)
    extra = f"\n{preamble}" if preamble else ""
    inner = _strip_table_float(tex_content)
    setup, body = _split_leading_setup(inner)
    # Auto-wrap in resizebox for preview if not already present
    if "\\resizebox" not in body:
        body = f"\\resizebox{{\\linewidth}}{{!}}{{{body}}}"
    inner = f"{setup}\n{body}".strip() if setup else body
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
    latex_backend: Optional[str] = None,
    preamble: Optional[str] = None,
) -> Path:
    """Compile LaTeX content to PDF.

    Args:
        tex_content: Raw LaTeX table code.
        output: Output PDF path.
        theme: Theme name (for package imports).
        latex_backend: Explicit LaTeX backend.
        preamble: Extra LaTeX preamble (e.g. custom commands).

    Returns:
        Path to generated PDF.

    Raises:
        RuntimeError: If compilation fails.
    """
    pdflatex = ensure_pdflatex()
    canonical_theme, resolved_backend = _resolve_preview_theme_backend(
        tex_content, theme, latex_backend
    )
    if resolved_backend == "tabularray" and "\\begin{tblr}" in tex_content:
        tex_content = _sanitize_tblr_for_compile(tex_content)
    doc = _build_standalone(
        tex_content,
        canonical_theme,
        latex_backend=resolved_backend,
        preamble=preamble,
    )
    output = Path(output)

    installed_in_run = set()
    max_retries = 2

    with tempfile.TemporaryDirectory() as tmpdir:
        tex_path = Path(tmpdir) / "table.tex"
        tex_path.write_text(doc)

        for _ in range(max_retries + 1):
            result = subprocess.run(
                [pdflatex, "-interaction=nonstopmode", "-output-directory", tmpdir, str(tex_path)],
                capture_output=True,
            )
            pdf_path = Path(tmpdir) / "table.pdf"
            if pdf_path.exists():
                shutil.copy2(pdf_path, output)
                return output

            out = result.stdout.decode("utf-8", errors="replace")
            err = result.stderr.decode("utf-8", errors="replace")
            log_text = f"{out}\n{err}"
            missing_sty = _extract_missing_sty(log_text)
            pkg = _sty_to_tlmgr_package(missing_sty) if missing_sty else None

            if not pkg or pkg in installed_in_run:
                raise RuntimeError(f"pdflatex failed:\n{out}\n{err}")

            print(
                f"pubtab: Missing LaTeX package '{missing_sty}'. Auto-installing '{pkg}'...",
                file=sys.stderr,
            )
            _tlmgr_install_package(pkg)
            installed_in_run.add(pkg)
    return output


def _sanitize_tblr_for_compile(tex_content: str) -> str:
    """Drop tabular-only rule/color commands that break inside tblr."""
    sanitized = re.sub(r"\\rowcolor(?:\[[^\]]*\])?\{[^}]*\}", "", tex_content)
    sanitized = re.sub(r"\\cmidrule(?:\([^)]*\))?(?:\[[^\]]*\])?\{[^}]*\}\s*", "", sanitized)
    return sanitized


def preview(
    tex_input: Union[str, Path],
    output: Optional[Union[str, Path]] = None,
    theme: str = "three_line",
    latex_backend: Optional[str] = None,
    dpi: int = 300,
    preamble: Optional[str] = None,
) -> Path:
    """Generate a PNG preview from LaTeX content or .tex file.

    Args:
        tex_input: LaTeX string or path to .tex file.
        output: Output PNG path. Defaults to input stem + .png.
        theme: Theme name.
        latex_backend: Explicit LaTeX backend.
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
        compile_pdf(
            tex_content,
            pdf_path,
            theme=theme,
            latex_backend=latex_backend,
            preamble=preamble,
        )
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
    except Exception:
        pass

    # Fallback: PyMuPDF (fitz) pure-Python rasterization.
    try:
        import fitz

        doc = fitz.open(str(pdf_path))
        if doc.page_count > 0:
            page = doc.load_page(0)
            scale = max(float(dpi) / 72.0, 1.0)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            pix.save(str(output))
            doc.close()
            if output.exists():
                return
        else:
            doc.close()
    except Exception:
        pass

    # Fallback: qlmanage (macOS, native high-quality), convert (ImageMagick), sips (last resort)
    if shutil.which("qlmanage"):
        # qlmanage renders PDF natively at target pixel size
        with tempfile.TemporaryDirectory() as tmpdir:
            # Calculate target pixel width from PDF points and DPI
            target_size = max(dpi * 6, 1800)  # at least 1800px wide
            result = subprocess.run(
                ["qlmanage", "-t", "-s", str(target_size), "-o", tmpdir, str(pdf_path)],
                capture_output=True,
            )
            # qlmanage outputs as <filename>.png in the output dir
            pngs = list(Path(tmpdir).glob("*.png"))
            if result.returncode == 0 and pngs:
                shutil.copy2(pngs[0], output)
                return
    elif shutil.which("magick"):
        result = subprocess.run(
            ["magick", str(pdf_path), "-density", str(dpi), str(output)],
            capture_output=True,
        )
        if result.returncode == 0 and output.exists():
            return
    elif platform.system() != "Windows" and shutil.which("convert"):
        result = subprocess.run(
            ["convert", "-density", str(dpi), str(pdf_path), str(output)],
            capture_output=True,
        )
        if result.returncode == 0 and output.exists():
            return
    else:
        raise RuntimeError(
            "No PDF-to-PNG converter found. Reinstall pubtab to ensure default "
            "PNG backends (`pdf2image`, PyMuPDF) are available, or use ImageMagick."
        )

    if not output.exists():
        raise RuntimeError("Failed to convert PDF to PNG: no output image generated.")
