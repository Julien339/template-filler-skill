"""Tests for the MCP server module (server.py)."""

from pathlib import Path

from template_filler_mcp.server import mcp

FIXTURES = Path(__file__).parent / "fixtures"


def test_server_name():
    assert mcp.name == "template-filler"


def test_all_tools_registered():
    tools = mcp._tool_manager._tools
    expected = {
        "extract_pptx",
        "extract_docx",
        "apply_pptx",
        "apply_docx",
        "verify_pptx",
        "verify_docx",
        "verify_parity_tool",
        "render_preview",
    }
    actual = set(tools.keys())
    assert actual == expected, f"Missing: {expected - actual}, Extra: {actual - expected}"


def test_extract_tool_callable():
    tool = mcp._tool_manager._tools["extract_pptx"]
    result = tool.fn(str(FIXTURES / "minimal.pptx"))
    assert result["format"] == "pptx"
    assert result["slide_count"] == 1


def test_verify_tool_callable():
    tool = mcp._tool_manager._tools["verify_pptx"]
    result = tool.fn(str(FIXTURES / "minimal.pptx"))
    assert result["ok"] is True


def test_render_tool_callable():
    tool = mcp._tool_manager._tools["render_preview"]
    result = tool.fn(str(FIXTURES / "minimal.pptx"), "/tmp/mcp_render_test", [1])
    # Should either render or gracefully skip
    assert "ok" in result
