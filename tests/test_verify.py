"""Tests for the structural verification module (_verify.py)."""

from pathlib import Path

from template_filler_mcp._verify import verify_docx_structure, verify_pptx_structure

FIXTURES = Path(__file__).parent / "fixtures"


def test_verify_pptx_valid():
    result = verify_pptx_structure(str(FIXTURES / "minimal.pptx"))
    assert result["ok"] is True
    assert result["problems"] == []


def test_verify_docx_valid():
    result = verify_docx_structure(str(FIXTURES / "minimal.docx"))
    assert result["ok"] is True
    assert result["problems"] == []


def test_verify_pptx_corrupt_zip():
    with open("/tmp/test_corrupt.pptx", "wb") as f:
        f.write(b"not a valid zip file at all")
    result = verify_pptx_structure("/tmp/test_corrupt.pptx")
    assert result["ok"] is False
    assert len(result["problems"]) > 0
    assert any("cannot open as ZIP" in p for p in result["problems"])


def test_verify_pptx_missing_file():
    result = verify_pptx_structure("/tmp/nonexistent_file.pptx")
    assert result["ok"] is False
    assert "cannot open" in result["problems"][0].lower()
