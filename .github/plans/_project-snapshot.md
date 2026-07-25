# Project Snapshot: template-filler-skill

**Generated**: 2026-07-25
**Repo**: /home/gw/opt/template-filler-skill
**Git**: git repo present

## Architecture

Single-directory Python script pipeline wrapped as a Claude Code skill. No package structure. No MCP server.

```
template-filler-skill/
├── SKILL.md              # Claude Code skill definition (4-stage pipeline)
├── README.md             # User-facing docs
├── requirements.txt      # python-pptx, python-docx, lxml, pymupdf
├── scripts/
│   ├── extract_pptx.py   # PPTX → JSON content map (125 lines)
│   ├── extract_docx.py   # DOCX → JSON content map (142 lines)
│   ├── apply_pptx.py     # JSON changes → PPTX output (113 lines)
│   ├── apply_docx.py     # JSON changes → DOCX output (105 lines)
│   ├── verify_pptx.py    # PPTX structural validation (127 lines)
│   ├── verify_docx.py    # DOCX structural validation (139 lines)
│   ├── verify_parity.py  # Before/after parity check (149 lines)
│   └── render_preview.py # LibreOffice + PyMuPDF render (89 lines)
└── docs/
    ├── demo-before.png
    └── demo-after.png
```

## Key Modules

| Module | Responsibility | Dependencies |
|--------|---------------|--------------|
| extract_pptx | Walk PPTX shapes/tables → flat JSON with stable IDs | python-pptx |
| extract_docx | Walk DOCX body+headers/footers → flat JSON | python-docx, lxml |
| apply_pptx | Mutate specific runs in-place, preserve formatting | python-pptx |
| apply_docx | Mutate specific runs in-place, preserve formatting | python-docx |
| verify_pptx | XML integrity, singleton children, media refs | lxml, python-pptx |
| verify_docx | XML integrity, singleton children, media refs | lxml, python-docx |
| verify_parity | Re-extract both files, assert untouched runs unchanged | Both extract scripts |
| render_preview | soffice → PDF → PyMuPDF → PNG per slide/page | LibreOffice, PyMuPDF |

## Test Commands

No existing test suite. Scripts are CLI-only: `python scripts/<name>.py <args>`.

## CI

No CI configuration present.

## Dependencies

- python-pptx >= 0.6.23
- python-docx >= 1.2.0
- lxml >= 4.9
- pymupdf >= 1.23
- External: LibreOffice (soffice) for render_preview
