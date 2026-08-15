"""Slide/page rendering to PNG for visual inspection.

Renders a .pptx or .docx to PNG images via LibreOffice + PyMuPDF. This is the
only way to catch visual issues (text overflow, autofit not reflowing) that
parity-based text comparisons cannot detect.
"""

from __future__ import annotations

import os
import shutil
import subprocess
from typing import Any

from ._path_validation import validate_input_file, validate_output_directory

SOFFICE_CANDIDATES = [
    "soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
]


def _find_soffice():
    for candidate in SOFFICE_CANDIDATES:
        if os.path.sep in candidate or ":" in candidate:
            if os.path.exists(candidate):
                return candidate
        else:
            found = shutil.which(candidate)
            if found:
                return found
    return None


def render_preview_pages(
    filepath: str,
    output_dir: str,
    pages: list[int] | None = None,
) -> dict[str, Any]:
    """Render specific pages of a PPTX/DOCX file to PNG images.

    Args:
        filepath: Path to the .pptx or .docx file.
        output_dir: Directory to save rendered PNG images.
        pages: Optional list of 1-based page numbers to render.
            If None, renders all pages (may be slow for large decks).

    Returns:
        {"ok": True, "rendered": ["path1.png", ...]}
        or {"ok": False, "skipped_reason": "soffice not found"}
    """
    validated_file = validate_input_file(filepath, allowed_extensions=(".pptx", ".docx"))
    validated_output_dir = validate_output_directory(output_dir)

    soffice = _find_soffice()
    if not soffice:
        return {
            "ok": False,
            "skipped_reason": "LibreOffice (soffice) not found. Install with: apt install libreoffice",
        }

    ext = os.path.splitext(str(validated_file))[1].lower()
    label = "slide" if ext == ".pptx" else "page"

    os.makedirs(str(validated_output_dir), exist_ok=True)

    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(validated_output_dir), str(validated_file)],
            check=True,
            capture_output=True,
            timeout=120,
        )
    except subprocess.TimeoutExpired:
        return {"ok": False, "skipped_reason": "soffice conversion timed out (>120s)"}
    except subprocess.CalledProcessError as e:
        return {"ok": False, "skipped_reason": f"soffice conversion failed: {e.stderr.decode()[:200]}"}
    except FileNotFoundError:
        return {"ok": False, "skipped_reason": "soffice binary not found"}

    pdf_path = os.path.join(str(validated_output_dir), os.path.splitext(os.path.basename(str(validated_file)))[0] + ".pdf")

    try:
        import fitz
    except ImportError:
        return {"ok": False, "skipped_reason": "PyMuPDF (fitz) not installed"}

    doc = fitz.open(pdf_path)
    out_paths = []
    for i, page in enumerate(doc):
        page_num = i + 1
        if pages is not None and page_num not in pages:
            continue
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        out_png = os.path.join(str(validated_output_dir), f"{label}{page_num}.png")
        pix.save(out_png)
        out_paths.append(out_png)
    doc.close()

    return {"ok": True, "rendered": out_paths}
