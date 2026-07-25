"""Tests for the extraction module (_extract.py)."""

from pathlib import Path

from template_filler_mcp._extract import extract_docx_content, extract_pptx_content

FIXTURES = Path(__file__).parent / "fixtures"


def test_extract_pptx_returns_valid_structure():
    result = extract_pptx_content(str(FIXTURES / "minimal.pptx"))
    assert result["format"] == "pptx"
    assert result["source"].endswith("minimal.pptx")
    assert result["slide_count"] == 1
    assert isinstance(result["slides"], list)
    assert len(result["slides"]) == 1


def test_extract_pptx_slide_has_runs():
    result = extract_pptx_content(str(FIXTURES / "minimal.pptx"))
    slide = result["slides"][0]
    assert len(slide["runs"]) > 0
    # Each run must have id, text, paragraph_text
    for run in slide["runs"]:
        assert "id" in run
        assert "text" in run
        assert "paragraph_text" in run


def test_extract_pptx_run_id_format():
    result = extract_pptx_content(str(FIXTURES / "minimal.pptx"))
    runs = result["slides"][0]["runs"]
    # All IDs should follow {slide}/{shape_id}/{para}/{run} or /r{row}c{col}/
    for run in runs:
        parts = run["id"].split("/")
        assert len(parts) >= 3, f"ID {run['id']!r} has <3 parts"


def test_extract_docx_returns_body_and_headers():
    result = extract_docx_content(str(FIXTURES / "minimal.docx"))
    assert result["format"] == "docx"
    assert result["section_count"] == 1
    assert len(result["body_runs"]) > 0
    # Header should be present since we added one
    assert len(result["header_footer_runs"]) > 0


def test_extract_docx_paragraph_text_preserved():
    result = extract_docx_content(str(FIXTURES / "minimal.docx"))
    # First body run should be the "Hello DOCX" paragraph
    first = result["body_runs"][0]
    assert first["paragraph_text"] == "Hello DOCX"
    assert first["text"] == "Hello DOCX"


def test_extract_pptx_output_matches_legacy_script():
    """Verify new module output matches original script output."""
    import json
    import subprocess
    import sys

    # Run original script
    out_path = "/tmp/legacy_map.json"
    subprocess.run(
        [sys.executable, "scripts/extract_pptx.py", str(FIXTURES / "minimal.pptx"), out_path],
        check=True,
        capture_output=True,
    )
    with open(out_path) as f:
        legacy = json.load(f)

    # Run new module
    new = extract_pptx_content(str(FIXTURES / "minimal.pptx"))

    assert legacy["slide_count"] == new["slide_count"]
    assert len(legacy["slides"]) == len(new["slides"])
    # Compare each run
    for o_run, n_run in zip(legacy["slides"][0]["runs"], new["slides"][0]["runs"]):
        assert o_run["id"] == n_run["id"]
        assert o_run["text"] == n_run["text"]
        assert o_run["paragraph_text"] == n_run["paragraph_text"]
