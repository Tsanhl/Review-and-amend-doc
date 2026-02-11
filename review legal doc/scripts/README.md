# DOCX refinement (style-preserving)

If you have:
- an **original** `.docx` (correct formatting/footnotes/fonts), and
- an **amended** `.docx` (better wording, but formatting got messed up),

use `refine_docx_from_amended.py` to copy only the *text* improvements into the
original document while preserving the original formatting.

## Usage

### Option A — Amended DOCX (recommended)

```bash
python3 "review docx/scripts/refine_docx_from_amended.py" \
  --original "/path/to/original.docx" \
  --amended  "/path/to/amended.docx" \
  --out      "/path/to/original_refined.docx"
```

Clean output (no bold/highlight on changes):
- `python3 "review docx/scripts/refine_docx_from_amended.py" --no-markup --original "/path/to/original.docx" --amended "/path/to/amended.docx" --out "/path/to/original_refined.docx"`

### Option B — Amended text file

`amended.txt` must be UTF-8 and use **blank lines** between paragraphs so the
paragraph count matches the original DOCX.

```bash
python3 "review docx/scripts/refine_docx_from_amended.py" \
  --original "/path/to/original.docx" \
  --amended-txt "/path/to/amended.txt" \
  --out      "/path/to/original_refined.docx"
```

## Guarantees

- Preserves the original DOCX package (styles, fonts/sizes, spacing/layout, headers/footers, footnotes, etc).
- Marks inserted/replaced wording as **bold + yellow highlight**.

## Notes / limitations

- The original and amended files must have the **same paragraph count** in `word/document.xml` (i.e., the amended DOCX must preserve paragraph breaks). If not, the script will fail with a clear error.
- Paragraphs containing hyperlinks/fields are skipped (reported as “skipped”).

## Review report DOCX

To produce the mandatory report artifact as `.docx`:

```bash
python3 "review legal doc/scripts/generate_review_report_docx.py" \
  --input "/path/to/review_report.md" \
  --out   "/path/to/review_report.docx"
```
Report generator updated: clean numbered ledger labels, markdown cleanup, no duplicated numbering.
