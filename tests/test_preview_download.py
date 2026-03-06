from __future__ import annotations

import io
import ssl
import subprocess
import urllib.error
from pathlib import Path

from pubtab import _preview


class _FakeResp(io.BytesIO):
    def __enter__(self) -> "_FakeResp":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()


def test_download_archive_success(tmp_path: Path, monkeypatch) -> None:
    payload = b"tinytex-archive-bytes"

    def _fake_urlopen(req, context=None):
        return _FakeResp(payload)

    monkeypatch.setattr(_preview.urllib.request, "urlopen", _fake_urlopen)
    out = tmp_path / "tinytex.tgz"
    _preview._download_archive("https://example.com/tinytex.tgz", out)
    assert out.read_bytes() == payload


def test_download_archive_ssl_error_message(tmp_path: Path, monkeypatch) -> None:
    def _fake_urlopen(req, context=None):
        raise urllib.error.URLError(ssl.SSLCertVerificationError("bad cert"))

    monkeypatch.setattr(_preview.urllib.request, "urlopen", _fake_urlopen)

    out = tmp_path / "tinytex.tgz"
    try:
        _preview._download_archive("https://example.com/tinytex.tgz", out)
        assert False, "expected RuntimeError"
    except RuntimeError as exc:
        msg = str(exc)
        assert "SSL certificate verification" in msg
        assert "Install Certificates.command" in msg


def test_extract_missing_sty_and_mapping() -> None:
    log = ".... LaTeX Error: File `booktabs.sty' not found. ...."
    assert _preview._extract_missing_sty(log) == "booktabs"
    assert _preview._sty_to_tlmgr_package("pifont") == "psnfss"
    assert _preview._sty_to_tlmgr_package("booktabs") == "booktabs"


def test_compile_pdf_auto_installs_missing_sty_once(tmp_path: Path, monkeypatch) -> None:
    installed = []
    run_count = {"n": 0}

    def _fake_ensure_pdflatex():
        return "/usr/bin/pdflatex"

    def _fake_install(package: str):
        installed.append(package)

    def _fake_run(cmd, capture_output=True, check=False):
        run_count["n"] += 1
        out_dir = Path(cmd[cmd.index("-output-directory") + 1])
        if run_count["n"] == 1:
            return subprocess.CompletedProcess(
                cmd,
                1,
                stdout=b"LaTeX Error: File `booktabs.sty' not found.",
                stderr=b"",
            )
        (out_dir / "table.pdf").write_bytes(b"%PDF-1.4\n")
        return subprocess.CompletedProcess(cmd, 0, stdout=b"", stderr=b"")

    monkeypatch.setattr(_preview, "ensure_pdflatex", _fake_ensure_pdflatex)
    monkeypatch.setattr(_preview, "_tlmgr_install_package", _fake_install)
    monkeypatch.setattr(_preview.subprocess, "run", _fake_run)

    out = tmp_path / "out.pdf"
    _preview.compile_pdf("\\begin{tabular}{c}A\\end{tabular}", out)
    assert out.exists()
    assert installed == ["booktabs"]
    assert run_count["n"] == 2
