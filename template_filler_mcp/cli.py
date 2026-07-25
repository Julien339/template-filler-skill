"""Standalone CLI for template-filler — usable without an MCP client.

Provides subcommands matching the MCP tools for direct command-line use:
    template-filler extract pptx <path>
    template-filler apply pptx <template> --changes <json_file> <output>
    template-filler verify pptx <path>
    template-filler parity <original> <output> --changes <json_file>
    template-filler render <file> <output_dir> [--pages 1,3,5]
"""

from __future__ import annotations

import json
from typing import Optional

import typer

from template_filler_mcp import __version__
from template_filler_mcp._apply import apply_docx_changes, apply_pptx_changes
from template_filler_mcp._extract import extract_docx_content, extract_pptx_content
from template_filler_mcp._parity import verify_parity
from template_filler_mcp._render import render_preview_pages
from template_filler_mcp._verify import verify_docx_structure, verify_pptx_structure

app = typer.Typer(
    name="template-filler",
    help="Fill existing PPTX/DOCX templates with new content while preserving formatting.",
)


def _load_changes(path: str) -> list[dict]:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ── extract ────────────────────────────────────────────────────────────

extract_app = typer.Typer(help="Extract text content from a template file.")
app.add_typer(extract_app, name="extract")


@extract_app.command("pptx")
def extract_pptx_cmd(path: str = typer.Argument(..., help="Path to .pptx template")):
    """Extract all text runs from a PPTX file and print as JSON."""
    result = extract_pptx_content(path)
    typer.echo(json.dumps(result, ensure_ascii=False, indent=2))


@extract_app.command("docx")
def extract_docx_cmd(path: str = typer.Argument(..., help="Path to .docx template")):
    """Extract all text runs from a DOCX file and print as JSON."""
    result = extract_docx_content(path)
    typer.echo(json.dumps(result, ensure_ascii=False, indent=2))


# ── apply ──────────────────────────────────────────────────────────────

apply_app = typer.Typer(help="Apply changes to a template file.")
app.add_typer(apply_app, name="apply")


@apply_app.command("pptx")
def apply_pptx_cmd(
    template: str = typer.Argument(..., help="Path to original .pptx template"),
    changes_file: str = typer.Option(..., "--changes", help="Path to changes.json file"),
    output: str = typer.Argument(..., help="Path for output .pptx"),
):
    """Apply text changes from a changes.json file to a PPTX template."""
    changes = _load_changes(changes_file)
    result = apply_pptx_changes(template, changes, output)
    typer.echo(f"Applied {result['applied']}/{result['total']} changes -> {output}")
    for f in result.get("failed", []):
        typer.echo(f"  FAILED: {f['id']!r}: {f['error']}", err=True)


@apply_app.command("docx")
def apply_docx_cmd(
    template: str = typer.Argument(..., help="Path to original .docx template"),
    changes_file: str = typer.Option(..., "--changes", help="Path to changes.json file"),
    output: str = typer.Argument(..., help="Path for output .docx"),
):
    """Apply text changes from a changes.json file to a DOCX template."""
    changes = _load_changes(changes_file)
    result = apply_docx_changes(template, changes, output)
    typer.echo(f"Applied {result['applied']}/{result['total']} changes -> {output}")
    for f in result.get("failed", []):
        typer.echo(f"  FAILED: {f['id']!r}: {f['error']}", err=True)


# ── verify ─────────────────────────────────────────────────────────────

verify_app = typer.Typer(help="Validate file structural integrity.")
app.add_typer(verify_app, name="verify")


@verify_app.command("pptx")
def verify_pptx_cmd(path: str = typer.Argument(..., help="Path to .pptx file")):
    """Validate a PPTX file for structural/XML integrity."""
    result = verify_pptx_structure(path)
    if result["ok"]:
        typer.echo(f"OK — {path} passed structural validation.")
    else:
        typer.echo(f"FAIL — {len(result['problems'])} problem(s) in {path}:", err=True)
        for p in result["problems"]:
            typer.echo(f"  - {p}", err=True)
        raise typer.Exit(code=1)


@verify_app.command("docx")
def verify_docx_cmd(path: str = typer.Argument(..., help="Path to .docx file")):
    """Validate a DOCX file for structural/XML integrity."""
    result = verify_docx_structure(path)
    if result["ok"]:
        typer.echo(f"OK — {path} passed structural validation.")
    else:
        typer.echo(f"FAIL — {len(result['problems'])} problem(s) in {path}:", err=True)
        for p in result["problems"]:
            typer.echo(f"  - {p}", err=True)
        raise typer.Exit(code=1)


# ── parity ─────────────────────────────────────────────────────────────


@app.command("parity")
def parity_cmd(
    original: str = typer.Argument(..., help="Path to original template"),
    output: str = typer.Argument(..., help="Path to filled output"),
    changes_file: str = typer.Option(..., "--changes", help="Path to changes.json"),
):
    """Verify output matches original except for requested changes."""
    changes = _load_changes(changes_file)
    result = verify_parity(original, output, changes)
    if result["ok"]:
        typer.echo(f"OK — {output} matches {original} everywhere except the requested changes.")
    else:
        typer.echo(f"FAIL — {len(result['problems'])} parity problem(s):", err=True)
        for p in result["problems"]:
            typer.echo(f"  - {p}", err=True)
        raise typer.Exit(code=1)


# ── render ─────────────────────────────────────────────────────────────


@app.command("render")
def render_cmd(
    file: str = typer.Argument(..., help="Path to .pptx or .docx file"),
    output_dir: str = typer.Argument(..., help="Directory for rendered PNGs"),
    pages: Optional[str] = typer.Option(None, "--pages", help="Comma-separated page numbers (e.g. 1,3,5)"),
):
    """Render slides/pages to PNG for visual inspection."""
    page_list = None
    if pages:
        page_list = [int(n.strip()) for n in pages.split(",") if n.strip()]
    result = render_preview_pages(file, output_dir, page_list)
    if result["ok"]:
        typer.echo(f"Rendered {len(result['rendered'])} page(s) to {output_dir}")
        for p in result["rendered"]:
            typer.echo(f"  {p}")
    else:
        typer.echo(f"Skipped — {result['skipped_reason']}", err=True)


@app.command("version")
def version_cmd():
    """Print version and exit."""
    typer.echo(__version__)
