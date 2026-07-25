"""Tests for the application module (_apply.py)."""

import shutil
import tempfile
from pathlib import Path

import pytest

from template_filler_mcp._apply import apply_docx_changes, apply_pptx_changes
from template_filler_mcp._extract import extract_docx_content, extract_pptx_content

FIXTURES = Path(__file__).parent / "fixtures"


def _tmp_copy(fixture_name):
    """Copy a fixture to a temp file for mutation."""
    tmp = tempfile.NamedTemporaryFile(suffix=fixture_name, delete=False)
    shutil.copy(FIXTURES / fixture_name, tmp.name)
    return tmp.name


def test_apply_single_run_change():
    out = "/tmp/test_apply_single.pptx"
    cm = extract_pptx_content(str(FIXTURES / "minimal.pptx"))
    first_id = cm["slides"][0]["runs"][0]["id"]

    changes = [{"id": first_id, "new_text": "Hello MCP"}]
    result = apply_pptx_changes(str(FIXTURES / "minimal.pptx"), changes, out)
    assert result["applied"] == 1
    assert result["total"] == 1
    assert len(result["failed"]) == 0

    # Re-extract and verify
    cm2 = extract_pptx_content(out)
    assert cm2["slides"][0]["runs"][0]["text"] == "Hello MCP"


def test_apply_paragraph_level_change():
    out = "/tmp/test_apply_para.pptx"
    cm = extract_pptx_content(str(FIXTURES / "minimal.pptx"))
    first_id = cm["slides"][0]["runs"][0]["id"]
    # Drop the trailing run index for paragraph-level
    para_id = "/".join(first_id.split("/")[:-1])

    changes = [{"id": para_id, "new_text": "Full paragraph replacement"}]
    result = apply_pptx_changes(str(FIXTURES / "minimal.pptx"), changes, out)
    assert result["applied"] == 1

    cm2 = extract_pptx_content(out)
    # All runs in that paragraph should have the new text as paragraph_text
    for run in cm2["slides"][0]["runs"]:
        if run["id"].startswith(para_id + "/"):
            assert run["paragraph_text"] == "Full paragraph replacement"


def test_apply_untouched_runs_preserved():
    out = "/tmp/test_apply_untouched.pptx"
    cm = extract_pptx_content(str(FIXTURES / "minimal.pptx"))
    runs = cm["slides"][0]["runs"]
    # Change only the first run
    first_id = runs[0]["id"]
    original_texts = {r["id"]: r["text"] for r in runs}

    changes = [{"id": first_id, "new_text": "Changed"}]
    apply_pptx_changes(str(FIXTURES / "minimal.pptx"), changes, out)

    cm2 = extract_pptx_content(out)
    runs2 = cm2["slides"][0]["runs"]
    for r2 in runs2:
        if r2["id"] == first_id:
            assert r2["text"] == "Changed"
        elif r2["id"] in original_texts:
            assert r2["text"] == original_texts[r2["id"]], f"Run {r2['id']} changed unexpectedly!"


def test_apply_invalid_id_handled():
    out = "/tmp/test_apply_invalid.pptx"
    changes = [{"id": "999/999/0/0", "new_text": "should fail"}]
    result = apply_pptx_changes(str(FIXTURES / "minimal.pptx"), changes, out)
    assert result["applied"] == 0
    assert len(result["failed"]) == 1


def test_apply_docx_table_cell():
    out = "/tmp/test_apply_docx_cell.docx"
    cm = extract_docx_content(str(FIXTURES / "minimal.docx"))
    # Find a table cell run
    cell_run = None
    for run in cm["body_runs"]:
        if run["text"] == "Key":
            cell_run = run
            break
    assert cell_run is not None

    changes = [{"id": cell_run["id"], "new_text": "NewKey"}]
    result = apply_docx_changes(str(FIXTURES / "minimal.docx"), changes, out)
    assert result["applied"] == 1

    cm2 = extract_docx_content(out)
    for run in cm2["body_runs"]:
        if run["id"] == cell_run["id"]:
            assert run["text"] == "NewKey"
            break
    else:
        pytest.fail("Changed run not found in output")
