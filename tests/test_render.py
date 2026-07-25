"""Tests for the render preview module (_render.py)."""

import os
import shutil
from pathlib import Path

from template_filler_mcp._render import _find_soffice, render_preview_pages

FIXTURES = Path(__file__).parent / "fixtures"


def test_find_soffice():
    """_find_soffice should return path or None."""
    result = _find_soffice()
    # On this system, soffice is available
    if shutil.which("soffice"):
        assert result is not None
    else:
        assert result is None


def test_render_pptx():
    out_dir = "/tmp/test_render_pptx"
    shutil.rmtree(out_dir, ignore_errors=True)

    result = render_preview_pages(str(FIXTURES / "minimal.pptx"), out_dir, pages=[1])

    if shutil.which("soffice"):
        assert result["ok"] is True
        assert len(result["rendered"]) == 1
        assert os.path.exists(result["rendered"][0])
        assert result["rendered"][0].endswith(".png")
    else:
        assert result["ok"] is False
        assert "skipped_reason" in result
