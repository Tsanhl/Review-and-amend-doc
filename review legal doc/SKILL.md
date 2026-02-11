---
name: review-docx
description: Perform a professional, lawyer-grade review of a DOCX essay or document. Use when the user asks to "review a docx", "check my essay", "proofread my paper", or "review my assignment". Covers grammar, fluency/coherence/structure, content/footnote/bibliography accuracy, and a final holistic pass — outputting a polished, 10/10 standard essay that stays near a user-requested target/limit (or near the original if none is given), while keeping all original DOCX formatting unchanged. If the user requests implemented changes, only the changed wording is marked in **bold + yellow highlight** (or, if the user requests a clean final, no visual markup is added).
---

# Workflow


## Architecture (required)

Use one main agent (orchestrator) plus specialist sub-agents run sequentially. Each pass builds on the previous one. The orchestrator owns the final output and ensures nothing is missed.

### Main agent (orchestrator) responsibilities

- Receive the user's DOCX file (and essay question / marking rubric, if provided).
- Determine the governing question source using this default priority:
  1. question/prompt pasted by the user in chat or terminal;
  2. question/prompt explicitly stated as included inside the DOCX (for example a `Question:` block at the top);
  3. if no question is provided, infer topic/thesis from the document and use that as the benchmark.
- Parse terminal prompt instructions first (scope, exclusions, word-count target/limit, and any "exclude bibliography/abbreviations" directions).
- Extract the full text, footnotes, and bibliography from the DOCX.
- Keep all original document formatting unchanged (font family, font size, spacing, indentation, heading styles, numbering, margins, and layout); do not normalise or restyle any content.
- Run all four review passes in order: Grammar → Fluency/Coherence/Structure → Accuracy → Final Holistic Check.
- Merge all corrections into a single, clean final output.
- Enforce word count targeting: keep the amended output near the user's requested target or word-limit for the whole essay. If no target/limit is given, preserve the original length (default ±2%).
- Present a summary of all changes made, grouped by category.
- When generating a refined DOCX (Pass 5), preserve the original DOCX styling exactly. For any implemented amendment, apply **bold + yellow highlight** to changed wording by default. Only skip markup if the user explicitly requests a clean final (no highlight/bold).

### Sub-agents (specialists)

1) **Grammar sub-agent (Pass 1)**
2) **Fluency, Coherence & Structure sub-agent (Pass 2)**
3) **Accuracy & Citations sub-agent (Pass 3)**
4) **Final Holistic sub-agent (Pass 4)**

If sub-agents are not available, emulate this architecture by running these roles sequentially and labelling outputs clearly.

---

## Non-negotiables

- **Zero tolerance for introduced errors.** Every correction must improve the original — never introduce new mistakes.
- **Preserve the author's voice and intent.** Do not rewrite the essay in a different style. Enhance, do not replace.
- **Tone default and override.** Default tone is formal, academic, and professional (lawyer-grade). If the user explicitly requests a different tone, follow the user’s tone request.
- **Preserve DOCX typography.** In refined DOCX output, retain the font family/size/style used by the user (including heading/body differences). Do not switch to a different default font.
- **User-original style is the only style source.** All amendments must be based on the user’s original DOCX font, font size, and paragraph style at the exact local position. Never substitute styles from another file or from application defaults.
- **Run-level style inheritance is mandatory (global rule).** For every inserted/amended DOCX segment (including bibliography lines), inherit typography from the nearest unchanged local run/paragraph in the same section: same font family, font size, paragraph style, spacing, indentation, and alignment. Never apply direct font-name or font-size overrides that differ from the user’s local style.
- **Paragraph property inheritance is mandatory.** For inserted/amended lines, clone paragraph properties (`style`, line spacing, space before/after, indents, alignment) from adjacent unchanged paragraphs in the same section. Never leave inserted lines with default/blank paragraph properties when neighbors use explicit settings.
- **No formatting mutations are allowed.** Improve content only. Do not change any existing formatting attribute (font, size, colour, italics, underline, spacing, alignment, indentation, list formatting, page layout, headers/footers, tables, captions, or styles). Only permitted formatting change: for implemented amendments, apply **bold + yellow highlight** to changed/added wording only (unless the user explicitly requests clean final mode).
- **Clean final mode.** If the user requests “clean final / no amendments / no highlight”, do not add any bold/highlight; implement wording changes only.
- **Output location guardrail.** Do not write amended/refined DOCX outputs to the skill workspace folder (`/Users/hltsang/Desktop/Skills`) by default. Ask the user first whether saving to Desktop is allowed. Only save to Desktop after explicit permission; otherwise save to a safe temporary path and report it.
- **Path disambiguation is mandatory.** If similarly named DOCX files exist in multiple Desktop folders, resolve and use the exact user-intended absolute path before editing. Do not assume Desktop root when an active project subfolder contains the target file.
- **Non-destructive amendment workflow is mandatory.** Never directly amend the user’s source DOCX file in place. First create a copy, apply amendments to that copy, and deliver a new output DOCX.
- **Word count compliance.** Keep the amended output near the user-requested target or declared word limit for the whole essay.
  - Terminal prompt constraints override defaults for amendment scope, but quality checks (accuracy, coherence, structure, citation integrity) still apply to the full submitted material unless the user explicitly says review scope is limited.
  - If the user gives a target number (for example, 2500 words), aim to stay within about ±2% unless the user asks for exact matching.
  - If the user gives a maximum limit, do not exceed it; target the upper band (about 95-100% of the limit) unless instructed otherwise.
  - If no target/limit is provided, keep the final output approximately the same length as the original (±2%).
  - Do not pad or cut content without purpose.
- **Exclusion handling is mandatory.** If the user provides bibliography and/or abbreviations content but marks it as excluded from amendment, do not edit those excluded sections; still run checks on them and report any accuracy/coherence/citation issues.
- **Every footnote and bibliography entry must be verified.** If a citation cannot be verified, flag it clearly — never silently remove or fabricate citations.
- **Absolute source integrity is mandatory.** Final delivered output must contain zero fake content, zero fake citations, zero fake footnotes, and zero fake bibliography entries.
- **Real-source integrity is mandatory.** Check that each reference exists and that its metadata is internally correct for that source type (for example: author-title-journal-volume/issue/year/pages; case name-neutral citation/court/year; legislation title/year/jurisdiction).
- **User-source and amendment-source parity is mandatory.** Apply the same 100% real-and-accurate verification standard to: (a) all sources already present in the user's original footnotes/bibliography, and (b) every source added, replaced, corrected, or amended during review. Never keep or introduce an unverified source.
- **No hallucinated references.** Never invent sources, page numbers, dates, or case names.
- **Default full-verification on every review request.** For both `review` and `review + amend`, verify all substantive content claims and all footnotes/references/bibliography entries (if present).
- **No-fake-source default for all outputs.** In every final output (review report and/or amended DOCX), do not present fake or unverified sources as valid. Anything unverified must be clearly marked unverified and must not be treated as established fact.
- **If the user provides the essay question or marking rubric, every structural and content decision must be evaluated against it.**
- **Question-source flexibility is default.** The governing question may come from chat/terminal text or from inside the DOCX when the user says it is included there.
- **All four passes are mandatory.** Do not skip or merge passes.
- **The final output must be publication-ready.** It must read as a polished, 10/10 piece of academic writing at professional lawyer standard.
- **When the user requests amendments (implemented output):** deliver only fully verified content/citations. No unresolved verification items may remain in the amended DOCX.
- **Review + amend is dual-delivery mandatory.** If the user asks to review and amend (or asks to amend), always deliver both: (1) an amended DOCX, and (2) a review report containing change summary + verification notes/ledger.
- **Review report artifact is mandatory.** A review request is not complete until a review report is created as a `.docx` file and its absolute path is returned.
- **First-pass completeness gate.** For `review + amend`, never send final completion unless both required files exist on disk: amended DOCX + review report DOCX.
- **Report format default.** Report must be produced as DOCX by default; in-chat summary text alone never satisfies report delivery.
- **Amendment markup default is mandatory.** Every implemented amendment output must visibly mark changed wording in **bold + yellow highlight**. Use clean, no-markup output only when the user explicitly asks for clean final/no highlight.
- **Citation style lock:** If the user requests a citation style, follow that style. If no style is requested, default to OSCOLA.
- **Bibliography-only requests are check-first.** If the user asks to review/amend the bibliography specifically, the primary task is to verify accuracy (real source, correct author/title/year/journal/court/legislation details, and compliance with the active citation style). Do not rewrite entries for style unless a concrete error is found.
- **When OSCOLA is the active style:** do not add a trailing full stop at the end of bibliography entries (for example, `Author, Title (Year)` not `Author, Title (Year).`).
- **When OSCOLA is the active style:** case names must be italicised in both footnotes and bibliography entries (for example, `*Donoghue v Stevenson*`).
- **Quote/apostrophe style preference.** Use typographic curly quotes/apostrophes in edited bibliography text (for example, `‘…’` and `’`), unless the user explicitly asks for straight quotes.

---

## Pass 1 — Grammar Check

**Agent: Grammar sub-agent**

Systematically review every sentence for:

1. **Spelling** — correct all misspellings; respect British vs American English consistency (detect which the user uses and maintain it throughout).
2. **Punctuation** — commas, semicolons, colons, apostrophes, quotation marks (single vs double), hyphens vs en-dashes vs em-dashes, full stops.
3. **Subject-verb agreement** — singular/plural concordance.
4. **Tense consistency** — maintain the dominant tense; flag and fix unwarranted shifts.
5. **Article usage** — correct missing, extra, or wrong articles (a/an/the).
6. **Pronoun reference** — ensure every pronoun has a clear, unambiguous antecedent.
7. **Parallelism** — fix faulty parallel structures in lists and comparisons.
8. **Sentence fragments and run-ons** — fix without altering meaning.
9. **Word choice / malapropisms** — flag words used incorrectly (e.g., "effect" vs "affect", "principle" vs "principal").
10. **Preposition usage** — correct non-standard or awkward prepositional phrases.

**Output:** Corrected text with a changelog listing every grammar fix (original → corrected, with line/sentence reference).

---

## Pass 2 — Fluency, Coherence & Structure

**Agent: Fluency/Coherence/Structure sub-agent**

Using the essay question (if provided) as the benchmark, review:

### 2A — Sentence-level fluency
- Eliminate awkward phrasing, redundancy, and wordiness.
- Improve readability without changing meaning.
- Vary sentence length and structure to avoid monotony.
- Replace vague language with precise terms.

### 2B — Paragraph-level coherence
- Each paragraph must have a clear topic sentence.
- Logical flow between sentences within each paragraph.
- Effective use of transition words and linking phrases.
- No abrupt jumps in logic.

### 2C — Essay-level structure
- **Introduction:** Does it clearly state the thesis/argument? Does it outline the structure of the essay? Does it engage the reader?
- **Body paragraphs:** Does each paragraph advance the argument? Are they in the most logical order? Is there a clear progression of ideas?
- **Conclusion:** Does it summarise the key arguments? Does it answer the essay question directly? Does it avoid introducing new material?
- **Overall arc:** Is the argument sustained and developed throughout? Is the essay balanced (no section disproportionately long or short)?

### 2D — Question alignment (if essay question provided)
- Does every section directly address the question asked?
- Are all parts of the question answered?
- Is the thesis responsive to the specific prompt?
- Flag any off-topic or tangential sections.

**Output:** Revised text with a changelog listing every fluency/coherence/structure change, grouped by sub-category (2A/2B/2C/2D).

---

## Pass 3 — Accuracy of Content, Footnotes & Bibliography

**Agent: Accuracy & Citations sub-agent**

### 3A — Content accuracy
- Verify factual claims, dates, statistics, case names, legislation titles, and named principles.
- Verify each material factual/legal claim against reliable primary or authoritative sources where possible.
- If a claim appears incorrect or unsupported, flag it with a correction or a note requesting the user to verify.
- Ensure legal/technical terminology is used correctly (if applicable).
- Check that any quoted material matches the cited source (to the extent verifiable).

### 3B — Footnote accuracy & fake-source replacement
- Every footnote must be checked for:
  - **Scope parity (mandatory)** — this applies equally to user-original footnotes and any new/replacement/amended footnotes created during review.
  - **Existence and authenticity** — verify that the cited source actually exists. Use web search to confirm the source is real (correct author, title, publication, year). If a footnote references a **fabricated, non-existent, or hallucinated source**:
    1. Flag it clearly as fake/unverifiable.
    2. **Find a real, relevant, and accurate replacement source** that supports the same claim in the text. Search for genuine academic sources, case law, legislation, or authoritative publications that make the same or a closely related point.
    3. Replace the fake footnote with the verified real source, formatted in the active citation style (default OSCOLA if no style is requested).
    4. Log the replacement in the changelog: `Fake: [original fake citation] → Replaced with: [real citation]`.
    5. If no suitable real source can be found to support the claim, flag the claim itself as unsupported and recommend the user either remove the claim or provide their own source.
  - **Correct citation format** — follow the user-requested citation style; if no style is requested, use OSCOLA. Convert mixed citations to the active style consistently across all footnotes.
  - **Metadata match (mandatory)** — ensure cited metadata matches the real source record:
    - Journal/article sources: author(s), article title, journal title, year, volume/issue, and first page/pinpoint must align.
    - Cases: case name, neutral citation or report citation, court, and year must align exactly.
    - Legislation: instrument title, year, section/regulation reference, and jurisdiction must align.
    - Books/chapters: author/editor, title, edition/year, publisher, and pinpoint (if used) must align.
  - **Author name(s)** — correctly spelled and in correct order for the active citation style.
  - **Title** — italicised or quoted correctly per source type (book, journal article, case, legislation, online source).
  - **Year, volume, issue, page numbers** — present and correctly formatted.
  - **Pinpoint references** — if a specific page or paragraph is cited, check it appears reasonable in context.
  - **Court and jurisdiction** — for case citations, ensure the court name and year are correct format.
  - **URL and access date** — for online sources, ensure URL is present and access date is included if required by the style.
  - **Cross-reference shorthand** — used correctly and referring to the right prior footnote. When OSCOLA is active: use 'ibid' (lowercase, not italicised) for immediately preceding source; use '(n X)' for earlier footnotes; do not use 'supra' or 'op cit' in OSCOLA.
  - **Sequential numbering** — footnotes must be numbered sequentially with no gaps or duplicates.
- Cross-reference: every in-text citation must have a corresponding footnote, and every footnote must correspond to a claim in the text.

### 3C — Bibliography / Reference List accuracy
- Every source cited in footnotes must appear in the bibliography (and vice versa — flag orphan entries).
- Apply the same verification standard to both user-original entries and any entries added/replaced/amended during review.
- Bibliography entries must follow the active citation style (default OSCOLA if no style is requested).
- For bibliography-focused requests, run an **accuracy audit first** and make only error-driven amendments:
  - Check for fake/non-existent sources.
  - Check for wrong author/title/year/volume/issue/page/court/jurisdiction metadata.
  - Check for wrong source type formatting under the active citation style.
  - Avoid unnecessary stylistic rewrites when an entry is already accurate.
- **When OSCOLA is the active style, apply OSCOLA bibliography rules:**
  - Divide into sections by source type: **Primary Sources** (Cases, Legislation, Treaties) and **Secondary Sources** (Books, Chapters in edited books, Journal articles, Online sources, etc.).
  - Within each section, list alphabetically by author surname (or case name / legislation title for primary sources).
  - Do NOT include pinpoint page references in bibliography entries (those belong only in footnotes).
  - Do NOT include 'ibid' or '(n X)' references in the bibliography.
  - Do NOT end bibliography entries with a full stop.
- Correct formatting per source type (book, journal, case, legislation, treaty, online source, etc.).
- No duplicate entries.
- Consistent punctuation and formatting across all entries, including typographic curly quotes/apostrophes where applicable.
- If any fake/replaced footnotes were corrected in 3B, update the bibliography to reflect the real replacement sources.

### 3D — Bibliography generation (if none exists)
- If the essay has **no bibliography / reference list** and the user requests one (or if the essay is academic and a bibliography is expected):
  1. Compile every unique source cited across all footnotes.
  2. Generate a complete bibliography in the active citation style (default OSCOLA).
  3. If OSCOLA is active, organise into OSCOLA sections: **Primary Sources** (Cases, Legislation, Treaties) and **Secondary Sources** (Books, Chapters, Journal Articles, Online Sources, etc.).
  4. Sort alphabetically within each section.
  5. Append the bibliography at the end of the essay.
  6. Log in the changelog: `Bibliography generated from X footnote sources.`

### 3E — Cross-referencing
- Verify internal cross-references ("as discussed in Part II above" — does Part II actually discuss that?).
- Check that any "see above" / "see below" / "supra" / "infra" references are accurate.

**Output:** Revised text with footnotes and bibliography corrected in the active citation style (default OSCOLA). Changelog listing every citation fix. A separate "Verification flags" list for any claims or citations that could not be fully verified (with suggestions).

### 3F — Verification gate before amendment delivery
- Build a verification ledger for all substantive claims and footnotes with statuses: `Verified`, `Corrected+Verified`, or `Unverified`.
- For standard review output, report all `Unverified` items clearly in Verification flags.
- If the user requests implemented amendments/refined DOCX, resolve all `Unverified` items first by:
  1. replacing with verified sources, or
  2. rewriting/removing unsupported claims.
- Do not deliver amended final DOCX while any `Unverified` item remains.

---

## Pass 4 — Final Holistic Check

**Agent: Final Holistic sub-agent**

This is the quality-assurance pass. Read the entire essay as a unified piece and check:

1. **Read-through:** Read the full essay start to finish. Does it flow naturally? Does it read as a polished, professional piece?
2. **Argument strength:** Is the argument convincing? Are there any logical gaps, unsupported claims, or weak links in reasoning?
3. **Consistency:** Terminology, spelling conventions (British/American), formatting, heading styles, numbering — all consistent throughout.
4. **Tone:** Appropriate academic register throughout. No informal language, contractions (unless stylistically intentional), or colloquialisms.
5. **Formatting:**
   - Heading hierarchy is consistent and logical.
   - Quotations are correctly formatted (short quotes inline, long quotes block-indented, per style guide).
   - Lists and enumerations are formatted consistently.
6. **Word count check:** Confirm the final word count stays near the user-requested target or within the user-specified limit. If no target/limit is provided, confirm original parity (±2%).
7. **10/10 standard test:** Would this essay receive top marks from a demanding marker? If not, identify precisely what prevents this and make targeted improvements.
8. **Final proofread:** One last character-by-character scan for any remaining typos, spacing issues, or formatting glitches.

**Output:** The final, polished essay — ready for submission.

---

## Pass 5 (Optional) — Implement Improvements Into Refined DOCX

**Agent: Main agent (orchestrator), only when user explicitly requests implementation**

This pass runs **after** the review is complete and only if the user asks in terminal to apply the improvements (for example: "implement improvements", "apply changes", "generate final refined docx").

**Implementation tip (local):** If you already have an "amended" DOCX whose wording is correct but whose formatting is broken, prefer copying only the text changes back into the original using `review docx/scripts/refine_docx_from_amended.py` (preserves fonts/spacing/footnotes; marks changes as bold + yellow highlight by default, or use `--no-markup` for a clean final).

1. Use the fully reviewed result (all accepted changes from Passes 1-4) as the source of truth.
2. Create the refined output as a copy of the original DOCX and edit only the required wording runs; do not rebuild/reflow the document from plain text.
   - Do not edit the user’s original file in place at any stage.
   - Keep the original file unchanged and produce a separate amended file path.
3. Preserve all original formatting exactly:
   - Keep the same font family, font size, style hierarchy, paragraph formatting, and page/layout settings everywhere.
   - For each amended run, keep all pre-existing run attributes and add **bold + yellow highlight** only to the changed wording.
   - For newly inserted wording or lines, clone local run/paragraph properties from adjacent unchanged content in the same section (do not fall back to document defaults).
   - Preserve mixed inline styling patterns where present (for example, italicised titles inside bibliography entries); do not flatten an inserted entry into one uniform run style.
   - For bibliography amendments specifically, ensure inserted/replaced entries inherit the same paragraph spacing and line spacing as neighboring bibliography entries.
   - Do not perform any document-wide style replacement or formatting cleanup.
4. Apply all refinements to the document text, footnotes, and bibliography.
5. Mark every refinement in **bold + yellow highlight**:
   - Bold + highlight only changed/added wording where possible.
   - Do not bold or highlight unchanged text.
   - Do not remove or alter existing formatting on unchanged text.
   - Do not add any formatting other than bold + yellow highlight to changed wording.
6. Keep word count compliance: stay near the user-requested target or within the user-specified word limit for the whole essay; if neither is provided, remain within ±2% of original.
7. Confirm output location before writing the refined DOCX:
   - Ask whether saving to Desktop is allowed.
   - Do not write output to `/Users/hltsang/Desktop/Skills` unless the user explicitly asks for that location.
   - If Desktop is approved, save to Desktop.
   - If Desktop is not approved (or no consent is given), save to a safe temporary path (for example, under `/tmp`) and report the exact path.
8. Save as a new file, never overwrite the original (recommended suffix: `_refined.docx`).
9. Before delivery, confirm there are zero unresolved verification items (all claims/citations verified or corrected).
10. Return the final file path/name in terminal and confirm that refinements are **bolded + highlighted (yellow)** and the original font styling is preserved.
11. Run a final formatting integrity check: confirm that no style/layout change occurred outside amended runs.
12. Run an amended-line style parity check before delivery: for each amended paragraph/run, verify parity with neighboring unchanged content for font family, font size, paragraph style, line spacing, and space before/after. If mismatch exists, fix before output.
13. Run a final path/output integrity check: verify the delivered file exists at the exact target path and is a real `.docx` document (not only a temporary lock file like `~$...docx`).

**Output:** Two mandatory deliverables for amendment requests:
1. A new refined DOCX with all implemented improvements in **bold + yellow highlight**.
2. A review report DOCX containing change summary and verification notes/ledger.

---

## Output Format

### Part A — Change Summary

Present a structured summary of all changes, organised by pass:

```
## Change Summary

### Pass 1 — Grammar (X changes)
- [list key changes with original → corrected]

### Pass 2 — Fluency/Coherence/Structure (X changes)
- [list key changes with brief rationale]

### Pass 3 — Citations & Accuracy (X changes)
- [list key corrections]
- **Verification flags:** [any items requiring user confirmation]
- **Verification ledger summary:** [Verified: X | Corrected+Verified: Y | Unverified: Z]
- **Citation-style compliance check:** [Pass/Fail, with any corrected non-compliant entries under active style]

### Pass 4 — Final Holistic (X changes)
- [list any final adjustments]

### Word Count
- Requested target (if any): XXXX words
- Original: XXXX words
- Final: XXXX words
```

### Part B — Final Reviewed Essay

Output the complete, final essay with all corrections applied. This includes:
- Full body text
- All footnotes (corrected and properly formatted)
- Complete bibliography / reference list (corrected and properly formatted)

**The final essay must be presented in full — no truncation, no summaries, no "rest remains the same". Every word of the reviewed essay must be output.**

### Part C (Only if requested) — Refined DOCX Delivery

- If the user asks to implement/apply improvements, generate and deliver a new refined `.docx`.
- Confirm:
  - Output filename/path
  - Output location permission status (Desktop approved or fallback temporary path used)
  - Original file kept unchanged
  - All refinements are **bold + highlighted (yellow)** in the refined document
  - All original formatting/style/layout is unchanged except amended wording
  - Zero unresolved verification items remain in the amended output

### Part D — Report Delivery (mandatory for all review requests)

- Deliver a review report as a `.docx` file and return its absolute path.
- The report must include:
  - change summary
  - content improvement roadmap to reach 10/10 excellence (what to add, strengthen/deepen, and reduce/remove)
  - verification notes and verification ledger
- In-chat report text is optional and supplementary only.
- Report DOCX typography default: Calibri (body) at 12pt throughout (including headings, lists, and ledger lines), unless the user explicitly requests otherwise.

### Part D1 — Citation Ledger Formatting (clean default)

- Use one numbered list level only; never output duplicated numbering (`1. 1.`).
- Reference-index entries: `1. [1] ...`
- Footnote entries: `1. Footnote 12: ...`
- Bibliography entries: `1. [12]. ...`
- Place verification URL on the next line as `Source: <url>` without extra numeric labels.

### Part E — Delivery Gate (mandatory)

Before final response, confirm required artifacts exist on disk:
- `review only`: review report DOCX exists.
- `review + amend` or `amend`: amended DOCX exists and review report DOCX exists.

If any required file is missing, generate it first and only then return completion.

---

## Special Instructions

- **If the user provides the essay question:** Use it as the primary benchmark for structural and content evaluation. State at the top of your review what question is being answered.
- **If no essay question is provided:** Infer the topic and thesis from the essay itself, and evaluate structure and content against that inferred thesis. State your inference and ask the user to confirm.
- **Terminal prompt scope is authoritative.** Apply amendments according to the latest terminal prompt instructions first (including include/exclude directives and word-count constraints).
- **If bibliography/abbreviations are supplied but marked as excluded:** do not amend those excluded sections, but still check and report accuracy, coherence, structure relevance, and citation integrity issues found there.
- **Citation style:** follow the user-requested style; if no style is requested, default to OSCOLA. If mixed styles are found, normalise to the active style.
- **Bibliography-only mode (when requested):** prioritise verification and error detection over rewriting. Amend only entries with identified issues (fake source, wrong metadata, wrong citation-style format).
- **When OSCOLA is the active style:** do not place a full stop at the end of bibliography entries.
- **Quote style in bibliography edits:** prefer typographic curly quotes/apostrophes (`‘…’`, `’`) rather than straight quotes.
- **Jurisdiction awareness:** If the essay is legal in nature, detect the jurisdiction from context (cases cited, legislation referenced) and ensure authorities and legal terminology are jurisdiction-accurate while keeping the active citation style formatting (default OSCOLA).
- **No word count limit on the review process.** However many words the user wrote is how many words you review and output.
- **If the user requests a specific final word count or gives an essay word limit:** Keep the amended output near that target/limit (default tolerance about ±2% unless the user asks for stricter matching). If a maximum limit is given, do not exceed it.
- **If the user asks to amend/implement improvements:** treat this as a request for a top-mark (10/10), professionally excellent lawyer-standard refined version, subject to strict verification and the active citation-style gate above (default OSCOLA if not requested otherwise).
- **Output location consent is mandatory for amended DOCX.** Before saving any refined document, ask whether Desktop output is allowed. Do not default to saving in `/Users/hltsang/Desktop/Skills`.
- **General DOCX amendment rule:** inserted/amended text must visually match user local style in font and size; amendment markup must be additive only and must not alter local typography.
- **If the DOCX contains images, tables, or charts:** Note their presence and positions but focus review on text content. Flag if any caption or label contains errors.
