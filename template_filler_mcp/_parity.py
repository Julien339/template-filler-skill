"""Parity verification for template-fill operations.

Re-extracts BOTH the original template and the written output and asserts that
the apply step changed exactly what the changes list asked for and nothing else.
This is the direct answer to "did my edit silently break something else."
"""

from __future__ import annotations

from typing import Any

from template_filler_mcp._extract import extract_docx_content, extract_pptx_content


def _runs_by_id(content_map, fmt):
    if fmt == "pptx":
        return {run["id"]: run for slide in content_map["slides"] for run in slide["runs"]}
    return {run["id"]: run for run in content_map["body_runs"] + content_map["header_footer_runs"]}


def _structure_signature(content_map, fmt):
    if fmt == "pptx":
        return {
            "slide_count": content_map["slide_count"],
            "run_count": sum(len(s["runs"]) for s in content_map["slides"]),
        }
    return {
        "section_count": content_map["section_count"],
        "body_run_count": len(content_map["body_runs"]),
        "header_footer_run_count": len(content_map["header_footer_runs"]),
    }


def verify_parity(original_path: str, output_path: str, changes: list[dict[str, str]]) -> dict[str, Any]:
    """Verify that the output file matches the original except for requested changes.

    This is the definitive check: every run NOT named in the changes list must
    be byte-identical text before/after; every run that IS in the changes list
    must show its expected new text; structural counts must match.

    Args:
        original_path: Path to the original template file (.pptx or .docx).
        output_path: Path to the filled output file.
        changes: The same changes list passed to apply_*_changes().

    Returns:
        {"ok": True, "problems": []} or {"ok": False, "problems": ["..."]}
    """
    fmt = "pptx" if original_path.lower().endswith(".pptx") else "docx"

    expected = {c["id"]: c["new_text"] for c in changes}

    original_map = (
        extract_pptx_content(original_path) if fmt == "pptx" else extract_docx_content(original_path)
    )
    output_map = extract_pptx_content(output_path) if fmt == "pptx" else extract_docx_content(output_path)

    original_runs = _runs_by_id(original_map, fmt)
    output_runs = _runs_by_id(output_map, fmt)

    problems = []

    # Determine which original runs are covered by the changes list.
    # A change id addresses either one exact run (present as-is) or a whole
    # paragraph (a prefix of one or more run ids — the paragraph-level fallback
    # from _apply.py).
    covered = set()
    run_level_checks = []
    para_level_checks = []

    for change_id, new_text in expected.items():
        if change_id in original_runs:
            run_level_checks.append((change_id, new_text))
            covered.add(change_id)
            continue
        prefix = change_id + "/"
        matched = [rid for rid in original_runs if rid.startswith(prefix)]
        if not matched:
            problems.append(f"changed id {change_id!r} matches no run or paragraph in the original")
            continue
        para_level_checks.append((change_id, new_text, matched))
        covered.update(matched)

    # Untouched runs must be identical.
    for run_id, old_run in original_runs.items():
        if run_id in covered:
            continue
        new_run = output_runs.get(run_id)
        if new_run is None:
            problems.append(f"run {run_id!r} present in original but missing in output")
        elif new_run["text"] != old_run["text"]:
            problems.append(
                f"run {run_id!r} changed but wasn't in changes list: "
                f"{old_run['text']!r} -> {new_run['text']!r}"
            )

    # Changed runs must show expected new text.
    for run_id, new_text in run_level_checks:
        actual = output_runs.get(run_id)
        if actual is None:
            problems.append(f"changed run {run_id!r} not found in output at all")
        elif actual["text"] != new_text:
            problems.append(f"run {run_id!r} expected {new_text!r}, got {actual['text']!r}")

    for prefix, new_text, matched in para_level_checks:
        sample = output_runs.get(matched[0])
        if sample is None:
            problems.append(f"paragraph {prefix!r} not found in output at all")
        elif sample["paragraph_text"] != new_text:
            problems.append(f"paragraph {prefix!r} expected {new_text!r}, got {sample['paragraph_text']!r}")

    # Structural counts must match.
    orig_sig = _structure_signature(original_map, fmt)
    out_sig = _structure_signature(output_map, fmt)
    if orig_sig != out_sig:
        problems.append(f"structure mismatch: original {orig_sig} vs output {out_sig}")

    return {"ok": len(problems) == 0, "problems": problems}
