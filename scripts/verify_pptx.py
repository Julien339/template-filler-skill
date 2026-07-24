"""
Validates a .pptx for the class of structural error that PowerPoint enforces
more strictly than python-pptx (or a naive well-formed-XML check) catches —
the kind that makes PowerPoint silently "repair" the file on open and drop
content, with no error raised anywhere in the Python pipeline.

Generic: works on any .pptx regardless of how it was produced. Shared
verbatim with the sibling pptx-builder-skill, where this exact failure mode
happened for real once (a shadow-effect helper appended a second
<a:effectLst> sibling instead of reusing the existing one — invalid per the
OOXML schema, but silent everywhere except PowerPoint's own repair step).

Usage:
    python verify_pptx.py <file.pptx>

Exits 0 and prints "OK" only if the file is clean. Exits 1 and lists every
problem otherwise. Always run this immediately after apply_pptx.py.
"""
import sys
import zipfile
from lxml import etree

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

# Elements the OOXML schema allows at most once among a given element's
# children. Duplicating any of these is exactly the kind of error that
# triggers a silent PowerPoint repair-and-drop-content on open.
SINGLETON_CHILDREN = {
    A + "effectLst", A + "effectDag", A + "xfrm", A + "ln",
    A + "scene3d", A + "sp3d", A + "custGeom", A + "prstGeom",
}


def check_zip_integrity(z, problems):
    bad = z.testzip()
    if bad:
        problems.append(f"zip integrity: corrupt entry {bad}")


def check_xml_wellformed(z, problems):
    parts = {}
    for name in z.namelist():
        if name.endswith(".xml") or name.endswith(".rels"):
            try:
                parts[name] = etree.fromstring(z.read(name))
            except Exception as e:
                problems.append(f"malformed XML: {name}: {e}")
    return parts


def check_duplicate_singletons(parts, problems):
    for name, root in parts.items():
        if not (name.startswith("ppt/slides/slide") and name.endswith(".xml")):
            continue
        for el in root.iter():
            counts = {}
            for child in el:
                counts[child.tag] = counts.get(child.tag, 0) + 1
            for tag, count in counts.items():
                if tag in SINGLETON_CHILDREN and count > 1:
                    problems.append(
                        f"{name}: <{etree.QName(el).localname}> has {count} "
                        f"<{etree.QName(tag).localname}> children (schema allows at most 1)"
                    )


def check_illegal_chars(parts, problems):
    for name, root in parts.items():
        if not (name.startswith("ppt/slides/slide") and name.endswith(".xml")):
            continue
        text = etree.tostring(root).decode("utf-8", errors="replace")
        for ch in text:
            if ord(ch) < 0x20 and ch not in "\t\n\r":
                problems.append(f"{name}: illegal XML control character U+{ord(ch):04X}")
                break


def check_media_rels(z, problems):
    names = set(z.namelist())
    for name in names:
        if not name.startswith("ppt/slides/_rels/"):
            continue
        root = etree.fromstring(z.read(name))
        for rel in root:
            target = rel.get("Target")
            if not target or rel.get("TargetMode") == "External":
                continue
            resolved = "ppt/" + target[3:] if target.startswith("../") else target
            if resolved not in names:
                problems.append(f"{name}: dangling relationship target {target!r}")


def check_opens_with_python_pptx(path, problems):
    try:
        from pptx import Presentation
        prs = Presentation(path)
        if len(prs.slides) == 0:
            problems.append("presentation has 0 slides")
    except Exception as e:
        problems.append(f"python-pptx failed to open the file: {e}")


def main():
    path = sys.argv[1]
    problems = []

    with zipfile.ZipFile(path) as z:
        check_zip_integrity(z, problems)
        parts = check_xml_wellformed(z, problems)
        if parts:
            check_duplicate_singletons(parts, problems)
            check_illegal_chars(parts, problems)
        check_media_rels(z, problems)

    check_opens_with_python_pptx(path, problems)

    if problems:
        print(f"FAIL — {len(problems)} problem(s) in {path}:", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        sys.exit(1)

    print(f"OK — {path} passed structural validation.")


if __name__ == "__main__":
    main()
