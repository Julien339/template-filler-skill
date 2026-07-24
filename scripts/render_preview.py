"""
Renders a .pptx or .docx to one PNG per slide/page via LibreOffice + PyMuPDF,
so the result can actually be looked at (with the Read tool) before calling
a template fill done.

verify_pptx.py/verify_docx.py catch file corruption. Neither catches content
that overflows a box, a table crushed too narrow, or a pptx text box whose
autofit didn't reflow after new (longer) text was written in — that class of
bug only shows up by looking at the rendered pages. This is exactly the risk
this skill's write-back step (setting run.text in place) can introduce.

Usage:
    python render_preview.py <file.pptx|file.docx> <output_dir> [--pages 2,5,7]

--pages takes a comma-separated list of 1-based slide/page numbers to render
only those pages — useful when a large deck/document only had a few slides
touched by apply_pptx.py/apply_docx.py.

Requires LibreOffice (soffice) — install with:
    winget install TheDocumentFoundation.LibreOffice
and the `pymupdf` package — install with:
    pip install pymupdf
"""
import sys
import os
import shutil
import subprocess

SOFFICE_CANDIDATES = [
    "soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
]


def find_soffice():
    for candidate in SOFFICE_CANDIDATES:
        if os.path.sep in candidate or ":" in candidate:
            if os.path.exists(candidate):
                return candidate
        else:
            found = shutil.which(candidate)
            if found:
                return found
    raise FileNotFoundError(
        "LibreOffice (soffice) not found. Install it with:\n"
        "  winget install TheDocumentFoundation.LibreOffice\n"
        "It's required to render a visual preview of the file."
    )


def main():
    in_path, out_dir = sys.argv[1], sys.argv[2]
    only_pages = None
    if "--pages" in sys.argv:
        raw = sys.argv[sys.argv.index("--pages") + 1]
        only_pages = {int(n) for n in raw.split(",") if n.strip()}

    ext = os.path.splitext(in_path)[1].lower()
    label = "slide" if ext == ".pptx" else "page"

    os.makedirs(out_dir, exist_ok=True)
    soffice = find_soffice()

    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, in_path],
        check=True, capture_output=True,
    )
    pdf_path = os.path.join(out_dir, os.path.splitext(os.path.basename(in_path))[0] + ".pdf")

    import fitz
    doc = fitz.open(pdf_path)
    out_paths = []
    for i, page in enumerate(doc):
        page_num = i + 1
        if only_pages is not None and page_num not in only_pages:
            continue
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        out_png = os.path.join(out_dir, f"{label}{page_num}.png")
        pix.save(out_png)
        out_paths.append(out_png)
    doc.close()

    print(f"Rendered {len(out_paths)} {label}(s) to {out_dir}")
    for p in out_paths:
        print(f"  {p}")


if __name__ == "__main__":
    main()
