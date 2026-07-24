---
name: template-filler
description: Use this skill to reuse an EXISTING, already-branded PowerPoint (.pptx) or Word (.docx) file for a new client/project — swap a client name, figures, dates, or other content everywhere it appears, while preserving 100% of the original styling, layout and masters. Different from pptx-builder-skill/docx-builder, which build a NEW document from an HTML design: this skill never touches formatting, it only rewrites the text of an existing file. Trigger on "fill in this template with...", "reuse last year's deck/report for...", "update this pptx/docx with the new client's info", "swap the numbers in this existing presentation/document".
---

# Template Filler

Fills an **existing** `.pptx`/`.docx` template with new content — no HTML,
no Playwright, no design stage. The file already has its branding; this
skill only changes what the file *says*, never how it looks.

**Zero template prep required.** There's no `{{placeholder}}` syntax to
insert first — the agent reads the template's current content and decides
what to change based on what it actually says (e.g. spotting last year's
client name, or last year's figures, wherever they appear).

## Stage 1 — Extract the current content

From this skill's directory, run the extractor matching the template's
format:

```bash
python scripts/extract_pptx.py <template.pptx> work/content_map.json
# or
python scripts/extract_docx.py <template.docx> work/content_map.json
```

This walks the **existing** file (via python-pptx/python-docx, reading the
real object model — not scraping HTML) and produces a flat JSON list of
every text run's current text, each with a stable ID:

- pptx: `{slide}/{shape_id}/{paragraph}/{run}` for text, `{slide}/{shape_id}
  /r{row}c{col}/{paragraph}/{run}` for table cells.
- docx: `{block}/{run}` for body paragraphs, `{block}/r{row}c{col}/{paragraph}
  /{run}` for table cells, `h{section}/{paragraph}/{run}` /
  `f{section}/{paragraph}/{run}` for each section's default header/footer.

Read `work/content_map.json` (it can be large — grep/filter it rather than
dumping the whole thing into context if the template has many slides/pages).
Each entry also carries a `paragraph_text` field: the full text of the
paragraph that run belongs to, because PowerPoint/Word routinely split what
looks like one phrase into several runs (autocorrect, a prior manual edit) —
read `paragraph_text` to understand what a run is really part of, don't
judge a fragment in isolation.

## Stage 2 — Decide what changes, and how

This is where "matching by content" actually happens, and it's the agent's
job, not a script's: read the content map, identify which runs/paragraphs
hold the values that need to change for the new client/project, and note
that a repeated value (a client name, a date) usually appears in **more
than one place** — check the whole map, don't stop at the first match.

Write `work/changes.json` as a list of `{"id": ..., "new_text": ...}`:

- Address a **single run** by its exact ID (`"2/5/0/0"`) when you want
  surgical precision.
- Address a **whole paragraph** by leaving off the trailing run index
  (`"2/5/0"`) when its `paragraph_text` shows it was split into more runs
  than are meaningful — the first run gets the new text, every other run in
  that paragraph is blanked. This is the normal way to replace a phrase that
  autocorrect fragmented, and is usually simpler than reconstructing exact
  run boundaries.
- Never invent an ID that isn't in the content map, and never rely on
  search-and-replace across the raw file — a value like "12" can appear
  as a real figure in one place and something unrelated (a page number, an
  unrelated count) elsewhere; only IDs seen in the content map are safe to
  target.

**Present an old → new diff table to the human and get explicit approval
before running Stage 3.** This mirrors the "approve before compile" rule
the sibling skills already use for HTML designs — nothing gets written
until the human has seen exactly what's about to change.

## Stage 3 — Write the new file

```bash
python scripts/apply_pptx.py <template.pptx> work/changes.json <output.pptx>
# or
python scripts/apply_docx.py <template.docx> work/changes.json <output.docx>
```

Opens the **original** file and writes the new text into only the runs
named in `changes.json`, in place — every other run, every shape,
paragraph, table, and all formatting is left completely untouched. Prints a
warning (not a silent failure) for any change ID it couldn't resolve; read
stderr and fix `changes.json` rather than ignoring it.

## Stage 4 — Verify (never skip this)

Three checks, all required before calling the fill done:

```bash
python scripts/verify_pptx.py <output.pptx>          # or verify_docx.py
python scripts/verify_parity.py <template.pptx> <output.pptx> work/changes.json
python scripts/render_preview.py <output.pptx> work/preview --pages <touched slide/page numbers>
```

- `verify_pptx.py`/`verify_docx.py` catch the same class of structural/XML
  corruption the sibling skills guard against (PowerPoint/Word silently
  "repairing" a file by dropping content, with no Python-side error).
- `verify_parity.py` is specific to this skill: it re-extracts both the
  original and the output and asserts that **every run not in
  `changes.json` is byte-identical text before/after**, every run that *is*
  in `changes.json` shows the expected new text, and slide/paragraph/table
  counts match. This is the direct check against "did the fill silently
  break something else" — treat any FAIL here as a bug, not a warning.
- `render_preview.py` renders the touched slides/pages to PNG — use
  `--pages` with just the numbers that changed, no need to re-render an
  entire 80-slide deck for a 2-slide edit. **Actually look at the rendered
  PNGs with the Read tool** before reporting success — this is the only way
  to catch a pptx text box whose autofit didn't reflow after much longer
  text was written in (a visual issue verify_parity.py cannot see; it only
  checks text, not layout).

## Known limits (design around these, don't fight them)

- **No image/logo swapping** — needs relationship-level (blip) XML surgery
  not implemented in this version. Swap images by hand in PowerPoint/Word
  after the text fill.
- **Only per-slide (pptx) / per-document-body+headers+footers (docx)
  content is addressable** — text inherited from a slide master/layout, or
  from non-default (first-page/even-page) headers/footers, is out of scope.
- **Field-coded runs are read-only** — dates, page numbers, and TOC entries
  built from Word/PowerPoint auto-fields are either invisible to the
  extractor (pptx) or flagged `"field": true` in the content map (docx);
  writing to a flagged run looks successful but Word regenerates it on next
  open, so don't target these IDs.
- **docx table cells** assume simple paragraph/run content — a table
  nested inside another table's cell isn't addressed.
- **Autofit isn't reflowed.** If new text is much longer than the original
  in a pptx text box with shrink-to-fit, the box may look visually off even
  though the write itself is correct — this is exactly why the
  render-preview review step in Stage 4 isn't optional.
