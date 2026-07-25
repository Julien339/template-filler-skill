"""FastMCP server for template-filler — fill existing PPTX/DOCX templates.

Exposes 8 tools covering the complete extract→apply→verify pipeline:
  - extract_pptx / extract_docx: read template content
  - apply_pptx / apply_docx: write changes in-place
  - verify_pptx / verify_docx: structural integrity
  - verify_parity: confirm only intended changes were made
  - render_preview: visual inspection of rendered slides/pages
"""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from template_filler_mcp._apply import apply_docx_changes, apply_pptx_changes
from template_filler_mcp._extract import extract_docx_content, extract_pptx_content
from template_filler_mcp._parity import verify_parity
from template_filler_mcp._render import render_preview_pages
from template_filler_mcp._verify import verify_docx_structure, verify_pptx_structure

mcp = FastMCP(
    "template-filler",
    instructions=(
        "Fill an existing, already-branded PowerPoint (.pptx) or Word (.docx) "
        "file with new content — swap client names, figures, dates everywhere "
        "they appear, while preserving 100% of the original styling, layout "
        "and masters. The 4-stage pipeline is: extract content → decide changes "
        "→ apply changes → verify parity."
    ),
)


# ── Extract tools ──────────────────────────────────────────────────────


@mcp.tool()
def extract_pptx(template_path: str) -> dict[str, Any]:
    """Extract all text runs from a PPTX file as a structured content map.

    Returns a content map with slides, shapes, and runs — each keyed by a
    stable structural ID like "{slide}/{shape_id}/{paragraph}/{run}".
    Table cells use "{slide}/{shape_id}/r{row}c{col}/{paragraph}/{run}".
    Group shape IDs are dot-joined. Merged cells are deduplicated.

    The content map is the artifact an agent reads to identify which runs
    to change for a new client/project/period. Use paragraph_text (not the
    isolated run.text) to understand what a run is really part of.
    """
    return extract_pptx_content(template_path)


@mcp.tool()
def extract_docx(template_path: str) -> dict[str, Any]:
    """Extract all text runs from a DOCX file as a structured content map.

    Returns a content map with body_runs (paragraphs + tables) and
    header_footer_runs. Body paragraph IDs are "{block}/{run}", table cells
    are "{block}/r{row}c{col}/{paragraph}/{run}", headers/footers are
    "h{section}/{paragraph}/{run}" / "f{section}/{paragraph}/{run}".

    Field-coded runs (dates, page numbers, TOC entries) are flagged with
    "field": true — treat these as read-only; Word regenerates them on open.
    """
    return extract_docx_content(template_path)


# ── Apply tools ────────────────────────────────────────────────────────


@mcp.tool()
def apply_pptx(template_path: str, changes: list[dict[str, str]], output_path: str) -> dict[str, Any]:
    """Apply text changes to a PPTX template by run ID.

    Args:
        template_path: Path to the original .pptx template file.
        changes: List of {"id": "...", "new_text": "..."} objects. Each id
            addresses one run (e.g. "0/5/0/0") or a whole paragraph
            (e.g. "0/5/0"). Only runs named in changes are touched;
            everything else — shapes, formatting, tables — is untouched.
        output_path: Where to save the filled presentation.

    Returns:
        {"applied": N, "total": M, "failed": [...], "output_path": "..."}
    """
    return apply_pptx_changes(template_path, changes, output_path)


@mcp.tool()
def apply_docx(template_path: str, changes: list[dict[str, str]], output_path: str) -> dict[str, Any]:
    """Apply text changes to a DOCX template by run ID.

    Args:
        template_path: Path to the original .docx template file.
        changes: List of {"id": "...", "new_text": "..."} objects.
        output_path: Where to save the filled document.

    Returns:
        {"applied": N, "total": M, "failed": [...], "output_path": "..."}
    """
    return apply_docx_changes(template_path, changes, output_path)


# ── Verify tools ───────────────────────────────────────────────────────


@mcp.tool()
def verify_pptx(filepath: str) -> dict[str, Any]:
    """Validate a PPTX file for structural/XML integrity.

    Checks: ZIP integrity, XML well-formedness, duplicate singleton elements
    (effectLst, xfrm, etc. — violations cause silent PowerPoint repairs),
    illegal XML control characters, dangling media relationships, and that
    python-pptx can open the file.

    Returns {"ok": True, "problems": []} if clean,
    or {"ok": False, "problems": [...]} with issue descriptions.
    """
    return verify_pptx_structure(filepath)


@mcp.tool()
def verify_docx(filepath: str) -> dict[str, Any]:
    """Validate a DOCX file for structural/XML integrity.

    Same checks as verify_pptx but for Word documents — singleton elements
    include pPr, rPr, tcPr, trPr, tblPr, etc.

    Returns {"ok": True, "problems": []} if clean,
    or {"ok": False, "problems": [...]} with issue descriptions.
    """
    return verify_docx_structure(filepath)


@mcp.tool()
def verify_parity_tool(
    original_path: str,
    output_path: str,
    changes: list[dict[str, str]],
) -> dict[str, Any]:
    """Verify that the output file matches the original except for the requested changes.

    This is the definitive check against silent corruption: every run NOT in
    the changes list must be byte-identical; every run in the list must show its
    expected new text; structural counts must match. Treat any FAIL here as a bug.

    Args:
        original_path: Path to the original template.
        output_path: Path to the filled output.
        changes: The same changes list passed to apply_pptx/apply_docx.

    Returns {"ok": True, "problems": []} or {"ok": False, "problems": [...]}.
    """
    return verify_parity(original_path, output_path, changes)


# ── Render tool ────────────────────────────────────────────────────────


@mcp.tool()
def render_preview(
    filepath: str,
    output_dir: str,
    pages: list[int] | None = None,
) -> dict[str, Any]:
    """Render PPTX/DOCX slides or pages to PNG images for visual inspection.

    Uses LibreOffice (soffice) to convert to PDF, then PyMuPDF to rasterize
    to PNG. This catches visual issues (text overflow, autofit not reflowing)
    that text-based parity checks cannot detect.

    Args:
        filepath: Path to .pptx or .docx file.
        output_dir: Directory for rendered PNG images.
        pages: Optional list of 1-based page numbers. None = render all.

    Returns {"ok": True, "rendered": ["path1.png", ...]}
    or {"ok": False, "skipped_reason": "..."} when dependencies are absent.
    """
    return render_preview_pages(filepath, output_dir, pages)


# ── Entry point ────────────────────────────────────────────────────────


def main() -> None:
    """Run the MCP server on stdio transport."""
    mcp.run()


if __name__ == "__main__":
    main()
