# GitHub Issue List

GitHub-ready issue drafts based on the `board-memo-refresh` benchmark and follow-up product discussion.

---

## 1) Add first-class footnote insertion: `deword footnote add`

**Title**

`feat: add first-class footnote insertion command`

**Problem**

`deword` is already good at changing existing text, but adding a new footnote currently requires raw OOXML plumbing:
- split a run at the anchor word
- insert `<w:footnoteReference>`
- allocate the next footnote id
- append a new `<w:footnote>` block
- preserve nearby formatting

This is too low-level for agent workflows and too fragile for routine document work.

**Proposed CLI**

```bash
deword footnote add doc.docx \
  --anchor "assessment" \
  --text "Assessment drafted April 2, 2026 from confirmed distributor staffing assumptions."
```

Optional flags:

```bash
--output output.docx
--scope paragraph|document
--occurrence 1
--after "assessment"
--before "assessment"
```

**Expected behavior**

- Find the anchor text uniquely in visible document text
- Split the containing run(s) safely
- Insert a properly formatted footnote reference at the anchor
- Allocate the next available footnote id automatically
- Create `word/footnotes.xml` if missing
- Preserve the surrounding paragraph/run formatting
- Fail loudly if the anchor is ambiguous or missing

**Why this matters**

This is the cleanest next step from the current benchmark. It turns a hard XML surgery task into a single high-level editing primitive.

**Acceptance criteria**

- Passes `benchmarks/board-memo-refresh` basic challenge without raw XML editing
- Result opens in Word without repair
- New superscript footnote marker renders correctly
- New footnote text appears at bottom of page in Word
- Existing header/footer/table formatting remains intact

---

## 2) Build an internal text-position insertion engine

**Title**

`feat: add run-aware text-position insertion engine for structural edits`

**Problem**

A lot of missing authoring features share the same hard problem:
- map visible text to OOXML runs
- split runs safely at character offsets
- preserve `w:rPr`
- insert non-text OOXML elements between text fragments

Without this abstraction, every new feature becomes custom XML surgery.

**Proposed internal API**

Examples:

- `findTextPosition(xml, text)`
- `splitRunAtOffset(xml, paragraphRef, charOffset)`
- `insertRunAfter(xml, runRef, newRunXml)`
- `cloneRunProperties(runXml)`
- `nextFootnoteId(files)`

**Expected behavior**

- Works across fragmented runs
- Preserves styling from surrounding text
- Supports insertion of:
  - footnote references
  - hyperlinks
  - comments
  - images
  - bookmarks

**Why this matters**

This is the foundation that makes future editing and authoring features fast to implement and safe to use.

**Acceptance criteria**

- Footnote insertion can be implemented on top of this engine
- No duplicate paragraphs or broken XML after insertion
- Existing `edit` / `replace` behavior remains unchanged

---

## 3) Add professional table insertion: `deword table add`

**Title**

`feat: add high-level table insertion command`

**Problem**

Agents and users often need to add structured tables to proposals, reports, contracts, and memos. Today this requires raw OOXML editing. That is exactly the kind of plumbing `deword` should hide.

**Proposed CLI**

```bash
deword table add doc.docx \
  --after-heading "Financial Summary" \
  --csv data.csv
```

Or:

```bash
deword table add doc.docx \
  --after "Financial Summary" \
  --json rows.json
```

Possible flags:

```bash
--style professional|grid|minimal
--header-row
--autofit
--caption "Quarterly Metrics"
--output output.docx
```

**Expected behavior**

- Insert a table after a matched paragraph/heading/anchor text
- Build valid OOXML table structure automatically
- Apply a reasonable professional default style
- Preserve document formatting around the insertion point
- Support basic data inputs from CSV/JSON

**Why this matters**

Tables are common in real human document work:
- board memos
- proposals
- operating reviews
- legal exhibits
- onboarding packets

**Acceptance criteria**

- User can insert a 2-column or multi-column table without touching XML
- Table opens/render correctly in Word
- Table styling looks professional by default
- Inserted table survives round-trip edits and `deword read`

---

## 4) Add professional image insertion: `deword image add`

**Title**

`feat: add high-level image insertion command`

**Problem**

Documents often need logos, screenshots, charts, signatures, and product images. Inserting an image cleanly in OOXML is error-prone and currently requires plumbing.

**Proposed CLI**

```bash
deword image add doc.docx \
  --after-heading "Market Overview" \
  --file chart.png \
  --alt "Regional sales chart"
```

Possible flags:

```bash
--width 480
--height 320
--align left|center|right
--wrap inline|block
--caption "Figure 1. Regional sales chart"
--output output.docx
```

**Expected behavior**

- Copy the image into `word/media/`
- Create/update relationships automatically
- Insert a valid drawing block at the requested location
- Support sensible defaults for professional document layout
- Preserve readability and surrounding styles

**Why this matters**

If `deword` can add tables and pictures professionally, it crosses from “edit tool” into “real document work tool” while still avoiding low-level Word plumbing.

**Acceptance criteria**

- Image renders correctly in Word
- No broken rels or missing media entries
- Alignment and sizing defaults look professional
- `deword read` still extracts the image correctly afterward

---

## 5) Add a high-level insertion/composition command to avoid plumbing

**Title**

`feat: add high-level insert/compose primitives so agents never need OOXML plumbing`

**Problem**

The current escape hatch (`deword xml`) is valuable, but it still exposes ZIP/XML details to the agent. The product goal should be:

> the agent says what to change, `deword` handles the OOXML plumbing.

**Proposed CLI direction**

A single family of commands such as:

```bash
deword insert footnote ...
deword insert table ...
deword insert image ...
```

Or a composition command:

```bash
deword compose doc.docx --plan edits.json
```

Where `edits.json` could express:
- replacements
- footnotes
- tables
- images
- paragraph insertion points

**Expected behavior**

- High-level agent-friendly commands
- Loud failure on ambiguous anchors
- Strong defaults for professional layout
- No need for users or agents to touch OOXML directly for common tasks

**Acceptance criteria**

- The `board-memo-refresh` benchmark can be solved with only high-level `deword` commands
- A proposal/customization benchmark can add at least one table and one image without raw XML editing

---

## 6) Add comments/endnotes/hyperlink insertion on top of the same engine

**Title**

`feat: support additional structural annotations using the same insertion engine`

**Problem**

Once `deword` can insert footnotes reliably, the same machinery should support other common document annotations and references.

**Candidate features**

- `deword comment add`
- `deword endnote add`
- `deword hyperlink add`

**Why this matters**

These are adjacent authoring primitives that humans use constantly in Word workflows.

**Acceptance criteria**

- Shared code path with the insertion engine
- Clean Word round-tripping
- No raw OOXML required for common annotation tasks

---

## 7) Add benchmark-oriented acceptance tests for human document tasks

**Title**

`test: add benchmark-driven regression tests for human document workflows`

**Problem**

We now have a concrete benchmark that represents real document work, not just internal feature coverage. The product should protect that workflow with regression tests.

**Initial benchmark set**

- `benchmarks/board-memo-refresh`
  - body/header/footer edits
  - table value edits
  - existing footnote edit
  - new footnote insertion

Future benchmark ideas:
- HR onboarding packet
- contract amendment
- proposal customization

**Acceptance criteria**

- A repeatable script can run the benchmark workflow
- Semantic assertions are machine-checkable
- Visual/manual review steps are documented
- Future authoring features are validated against realistic document tasks

---

## Suggested priority

1. `feat: add first-class footnote insertion command`
2. `feat: add run-aware text-position insertion engine for structural edits`
3. `feat: add high-level table insertion command`
4. `feat: add high-level image insertion command`
5. `feat: add high-level insert/compose primitives so agents never need OOXML plumbing`
6. `feat: support additional structural annotations using the same insertion engine`
7. `test: add benchmark-driven regression tests for human document workflows`
