---
name: review-general-docx
description: Perform a professional, top-mark review for the user-requested professional area/topic in DOCX documents (non-legal), including grammar, logic, coherence, structure, factual/citation integrity, and final quality assurance. Use when the user asks to "review" or "review and amend" a DOCX in any non-legal professional domain. For amend requests, deliver both an amended DOCX and a review report.
---

# Workflow

## Architecture (required)

Use one orchestrator plus specialist passes executed in order. The orchestrator is accountable for completeness and output quality.

### Main agent responsibilities

- Parse user instructions first (scope, exclusions, target/limit words, citation style request, clean-vs-marked output).
- Treat the user-provided question/prompt/rubric (if provided) as the primary benchmark for content evaluation.
- Determine the governing question source using this default priority:
  1. question/prompt pasted by the user in chat or terminal;
  2. question/prompt explicitly stated as included inside the DOCX (for example a `Question:` block at the top);
  3. if no question is provided, infer topic/thesis from the document and use that as the benchmark.
- Detect the document's topic/professional area from the question, headings, terminology, and references, and run review/amend as that area expert.
- Extract body text, footnotes/endnotes, and bibliography/reference list if present.
- Keep source DOCX formatting unchanged except amendment markup rules.
- Run all passes in order: Grammar -> Coherence/Structure -> Accuracy/Citations -> Final Holistic.
- Enforce word count target/limit when provided; otherwise keep near original length.
- Produce required outputs by mode (review-only vs review+amend).

### Specialist passes

1) Grammar pass  
2) Fluency, coherence, and structure pass  
3) Accuracy and citation/reference integrity pass  
4) Final holistic QA pass

If sub-agents are unavailable, emulate these passes sequentially.

---

## Operation Modes (mandatory)

### Mode A — Review + Amend (or amend request)

Must always deliver both outputs:

1. **Amended DOCX** implementing improvements at top-mark/10-10 standard.
2. **Review report DOCX** with change summary and verification notes/ledger.

Both files must be generated and present on disk before final reply.

### Mode B — Review only

Must deliver **review report DOCX** only (no implemented DOCX edits unless user asks to amend).

---

## Non-negotiables

- **Top-mark standard for amend.** Treat amend requests as a 10/10 refinement target.
- **Auto-domain-expert mode is mandatory.** For each document, identify the likely professional area/topic and apply that domain’s professional standards, terminology expectations, and reasoning quality during review and amendment.
- **Tone default and override.** Default tone is formal, academic, and professional. If the user explicitly requests a different tone, follow the user’s tone request.
- **Zero introduced errors.** Never degrade correctness.
- **Preserve author intent and voice** unless user asks for substantial stylistic shift.
- **No formatting mutations.** Do not alter layout/style/typography except approved amendment markup.
- **Amendment markup default is mandatory.** Changed wording must be **bold + yellow highlight** by default.
- **Clean mode exception.** Only skip markup when user explicitly requests clean final/no highlight.
- **Non-destructive workflow.** Never overwrite source DOCX; output a new file.
- **Word count compliance.** Respect user max/target; if absent, remain near original.
- **Citation/reference verification is mandatory when present.**
- **No fake sources.** Never invent citations, references, page numbers, dates, identifiers, or URLs.
- **User-source and amendment-source parity.** Apply the same strict verification standard to existing sources and newly added/replaced sources.
- **No unresolved verification for amend output.** If unverifiable, replace with verified support or flag and remove unsupported claim before final amended DOCX.
- **Default full-verification on every review request.** For both `review` and `review + amend`, verify all substantive content claims and all footnotes/references/bibliography entries (if present).
- **No-fake-source default for all outputs.** In every final output (review report and/or amended DOCX), do not present fake or unverified sources as valid. Anything unverified must be clearly marked unverified and must not be treated as established fact.
- **Question/prompt alignment is mandatory.** If the user provides a question/prompt/rubric, evaluate and amend content against it directly. If none is provided, infer the likely question/thesis from the document and state that inference in the review report.
- **Question-source flexibility is default.** The governing question may come from chat/terminal text or from inside the DOCX when the user says it is included there.
- **Citation style rule:** follow the user-requested citation style. If none is requested, preserve dominant existing style and normalise internal consistency.
- **Output location guardrail.** Confirm destination before writing outputs to Desktop; otherwise use safe temp path and report it.
- **Report artifact is mandatory.** A review request is not complete until a report file is created as `.docx` and its absolute path is returned.
- **First-pass completeness gate.** For `review + amend`, never send final completion unless both the amended DOCX and review report DOCX exist at final paths.
- **Report format default.** Report must be generated as DOCX by default; in-chat summary alone is never a substitute.

---

## Pass 1 — Grammar

Check spelling, punctuation, agreement, tense consistency, article/preposition usage, pronoun clarity, fragments/run-ons, and sentence correctness.

**Output artifact:** grammar fixes list (original -> corrected).

## Pass 2 — Fluency, Coherence, Structure

Check sentence flow, paragraph coherence, transitions, argument progression, section structure, intro/body/conclusion alignment, and removal of redundancy.

**Output artifact:** coherence/structure fixes list with rationale.

### 2D — Question alignment (mandatory when provided)

- Does each section directly address the user-provided question/prompt?
- Are all required parts of the question answered?
- Is the thesis/position responsive to the requested task?
- Flag and correct off-topic sections or missing required points.

## Pass 3 — Accuracy and Citation/Reference Integrity

### 3A Content accuracy

- Verify substantive facts, dates, names, numbers, and domain claims against reliable sources where possible.
- Correct unsupported or inaccurate claims.

### 3B Footnotes/endnotes and in-text citations

- Verify existence/authenticity of cited sources.
- Verify metadata fields for source type (author, title, venue, year, volume/issue/pages, court/case number, statute details, URL/date when required).
- If fake/unverifiable:
  1. Flag clearly.
  2. Replace with a real, relevant, accurate source where possible.
  3. Update citation in user-requested style.
  4. Log replacement in verification notes.

### 3C Bibliography/reference list (if present)

- Ensure source-list and citations are consistent (no orphan/missing items).
- Correct metadata and formatting inconsistencies.
- Remove duplicates.

### 3D Verification ledger (mandatory)

Track each substantive citation/claim as:
- `Verified`
- `Corrected+Verified`
- `Unverified`

For **review+amend**, final `Unverified` count must be zero.

## Pass 4 — Final Holistic QA

Read full document for final polish: logic continuity, clarity, consistency, tone, structure balance, and overall quality.

---

## Implementation (DOCX)

Use `scripts/refine_docx_from_amended.py` to transfer refined wording while preserving document styling and structure.

Recommended command pattern:

```bash
python3 "review general doc/scripts/refine_docx_from_amended.py" \
  --original "/path/to/original.docx" \
  --amended  "/path/to/amended.docx" \
  --out      "/path/to/output_refined.docx"
```

For clean final (only if explicitly requested):

```bash
python3 "review general doc/scripts/refine_docx_from_amended.py" \
  --no-markup \
  --original "/path/to/original.docx" \
  --amended  "/path/to/amended.docx" \
  --out      "/path/to/output_refined.docx"
```

### Review report generation (DOCX required)

Use `scripts/generate_review_report_docx.py` to write the report as a `.docx` file.

```bash
python3 "review general doc/scripts/generate_review_report_docx.py" \
  --input "/path/to/report.md" \
  --out   "/path/to/review_report.docx"
```

---

## Output Format

## Part A — Review Report (always required)

Include:

- Mode used: `Review only` or `Review + amend`
- Change summary by pass (Grammar / Coherence-Structure / Accuracy-Citations / Final QA)
- Key issues fixed
- Content improvement roadmap to reach 10/10 excellence:
  - what to add
  - what to strengthen/deepen
  - what to reduce/remove
- Verification notes and ledger summary: `Verified / Corrected+Verified / Unverified`
- Word count summary: original vs final (for amend)

Delivery requirements:
- Must be delivered as a `.docx` file with absolute output path.
- In-chat report text is optional and supplementary only.
- Report DOCX typography default: Calibri (body) at 12pt throughout (including headings, lists, and ledger lines), unless the user explicitly requests otherwise.

Citation-ledger formatting (clean default):
- Use a single numbered list level only (no duplicated numbering like `1. 1.`).
- For in-text reference mapping entries: `1. [1] ...`
- For footnote checks: `1. Footnote 12: ...`
- For bibliography checks: `1. [12]. ...`
- Put source links on the next line as `Source: <url>` (no extra numbering prefix).

## Part B — Amended DOCX (required for amend mode)

Return final amended file path and confirm:
- changed wording is bold + yellow (unless clean final requested)
- original formatting preserved outside amended text
- no unresolved verification items remain

## Part C — Delivery Gate (mandatory)

Before final response, verify required artifacts exist:
- `review only`: review report DOCX exists.
- `review + amend`: amended DOCX exists and review report DOCX exists.

If any required file is missing, generate it first and only then return final completion.
