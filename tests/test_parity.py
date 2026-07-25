"""Tests for the parity verification module (_parity.py)."""

from pathlib import Path

from template_filler_mcp._apply import apply_pptx_changes
from template_filler_mcp._extract import extract_pptx_content
from template_filler_mcp._parity import verify_parity

FIXTURES = Path(__file__).parent / "fixtures"


def test_parity_clean_apply():
    """A clean apply with correct changes should pass parity."""
    cm = extract_pptx_content(str(FIXTURES / "minimal.pptx"))
    first_id = cm["slides"][0]["runs"][0]["id"]
    changes = [{"id": first_id, "new_text": "Clean change"}]

    out = "/tmp/test_parity_clean.pptx"
    apply_pptx_changes(str(FIXTURES / "minimal.pptx"), changes, out)

    result = verify_parity(str(FIXTURES / "minimal.pptx"), out, changes)
    assert result["ok"] is True, f"Unexpected problems: {result['problems']}"


def test_parity_detects_extra_change():
    """If a run changed but isn't in changes.json, parity should fail."""
    cm = extract_pptx_content(str(FIXTURES / "minimal.pptx"))
    runs = cm["slides"][0]["runs"]

    # Change only the first run in changes.json
    changes = [{"id": runs[0]["id"], "new_text": "Changed"}]
    out = "/tmp/test_parity_extra.pptx"
    apply_pptx_changes(str(FIXTURES / "minimal.pptx"), changes, out)

    # But verify_parity claims a DIFFERENT run was changed
    wrong_changes = [{"id": runs[1]["id"], "new_text": "Wrong"}]
    result = verify_parity(str(FIXTURES / "minimal.pptx"), out, wrong_changes)
    assert result["ok"] is False, "Should detect that a run changed outside changes.json"
    assert len(result["problems"]) > 0


def test_parity_detects_nonexistent_id():
    """A change id that doesn't exist in the original should fail."""
    bad_changes = [{"id": "nonexistent/999/0/0", "new_text": "X"}]
    result = verify_parity(
        str(FIXTURES / "minimal.pptx"),
        str(FIXTURES / "minimal.pptx"),  # same file = no actual changes
        bad_changes,
    )
    assert result["ok"] is False
    assert any("matches no run" in p for p in result["problems"])
