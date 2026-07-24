"""
Re-extracts BOTH the original template and the written output with the same
extraction logic used by extract_pptx.py/extract_docx.py and asserts that
apply_pptx.py/apply_docx.py changed exactly what changes.json asked for and
nothing else. This is the direct answer to "did my edit silently break
something else" -- the exact question manual before/after backup copies
are otherwise the only defense against.

Compares TEXT and STRUCTURE (run/section counts), not raw file bytes --
python-pptx/python-docx can rewrite incidental XML (relationship ID order,
whitespace) on any save even when zero content changed, so a byte diff
would false-positive on a clean run.

Usage:
    python verify_parity.py <original.pptx|docx> <output.pptx|docx> <changes.json>

Exits 0 and prints "OK" only if every run not named in changes.json is
identical text before/after, every run named in changes.json now shows its
expected new text, and structural counts match. Exits 1 and lists every
discrepancy otherwise. Run this after every apply_pptx.py/apply_docx.py,
alongside verify_pptx.py/verify_docx.py and render_preview.py.
"""
import sys
import os
import json
import subprocess


def extract(fmt, path, tmp_json):
    script = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"extract_{fmt}.py")
    subprocess.run([sys.executable, script, path, tmp_json], check=True, capture_output=True)
    with open(tmp_json, encoding="utf-8") as f:
        return json.load(f)


def runs_by_id(content_map, fmt):
    if fmt == "pptx":
        return {
            run["id"]: run
            for slide in content_map["slides"]
            for run in slide["runs"]
        }
    return {
        run["id"]: run
        for run in content_map["body_runs"] + content_map["header_footer_runs"]
    }


def structure_signature(content_map, fmt):
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


def main():
    original_path, output_path, changes_path = sys.argv[1], sys.argv[2], sys.argv[3]
    fmt = "pptx" if original_path.lower().endswith(".pptx") else "docx"

    with open(changes_path, encoding="utf-8") as f:
        changes = json.load(f)
    expected = {c["id"]: c["new_text"] for c in changes}

    orig_tmp = original_path + ".orig_map.json"
    out_tmp = output_path + ".out_map.json"
    original_map = extract(fmt, original_path, orig_tmp)
    output_map = extract(fmt, output_path, out_tmp)

    original_runs = runs_by_id(original_map, fmt)
    output_runs = runs_by_id(output_map, fmt)

    problems = []

    # An expected change id addresses either one exact run (present as-is in
    # the ORIGINAL map) or a whole paragraph (a prefix of one or more run
    # ids -- apply_pptx.py/apply_docx.py's paragraph-level fallback). Every
    # original run covered by either form is exempt from the "must be
    # unchanged" check below.
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

    for run_id, old_run in original_runs.items():
        if run_id in covered:
            continue
        new_run = output_runs.get(run_id)
        if new_run is None:
            problems.append(f"run {run_id!r} present in original but missing in output")
        elif new_run["text"] != old_run["text"]:
            problems.append(
                f"run {run_id!r} changed but wasn't in changes.json: "
                f"{old_run['text']!r} -> {new_run['text']!r}"
            )

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

    orig_sig = structure_signature(original_map, fmt)
    out_sig = structure_signature(output_map, fmt)
    if orig_sig != out_sig:
        problems.append(f"structure mismatch: original {orig_sig} vs output {out_sig}")

    for tmp in (orig_tmp, out_tmp):
        try:
            os.remove(tmp)
        except OSError:
            pass

    if problems:
        print(f"FAIL — {len(problems)} parity problem(s):", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        sys.exit(1)

    print(f"OK — {output_path} matches {original_path} everywhere except the requested changes.")


if __name__ == "__main__":
    main()
