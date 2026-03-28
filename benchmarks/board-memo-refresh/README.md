# Board Memo Refresh (Basic)

## Why this benchmark exists

This is a realistic **basic** Word task:

- update a draft memo to a final board-review version
- fix dates and ownership details
- update numbers in a table
- update a confidentiality footer
- update one existing footnote and add one new footnote
- avoid changing protected historical text

A human executive assistant, finance ops person, chief of staff, or analyst could plausibly do this task in Word.

## Scenario

You have a 3-page board memo in Word. Leadership has approved a revised version for board review. Update the document without breaking formatting.

This is the **basic challenge**.

## Target task

The agent receives `prompt.txt` and the input `.docx`, and must produce an updated `.docx`.

## Scoring

### Semantic checks
Must:
- change the title and header text from `Draft Board Memo` to `Final Board Memo`
- change the board meeting date from `March 21, 2026` to `April 18, 2026`
- change `Prepared by: Lena Ortiz, Interim CFO` to `Prepared by: Maya Chen, CFO`
- change `FY2026 Revenue Target` value from `$12.4M` to `$14.1M`
- change `Launch Budget` value from `$2.8M` to `$3.2M`
- change footer text from `INTERNAL DRAFT — DO NOT DISTRIBUTE` to `CONFIDENTIAL — BOARD REVIEW`
- change the first footnote date from `January 12, 2026` to `March 30, 2026`
- add a new second footnote on the word `assessment`
- the new second footnote must read: `Assessment drafted April 2, 2026 from confirmed distributor staffing assumptions.`

Must not:
- leave behind stale values above
- modify the protected phrase `draft scenario model`

### Visual checks
- same overall layout as gold doc
- same page count as gold doc
- header/footer still present
- table still looks like a table
- no obvious broken spacing or missing lines

### File health
- opens in Microsoft Word
- exports to PDF successfully

## Files you will create

- `input/board-memo-input.docx`
- `gold/board-memo-gold.docx`
- optionally `gold/board-memo-gold.pdf`
- later, candidate outputs in `candidate/`

## Recommended first-pass workflow

1. Build the input doc in Word using `fixture-spec.md`
2. Save it as `input/board-memo-input.docx`
3. Duplicate it and manually apply the requested edits in Word
4. Save that as `gold/board-memo-gold.docx`
5. Export the gold doc as PDF from Word
6. Use `prompt.txt` as the benchmark prompt for the agent
7. Save candidate outputs in `candidate/`
8. Export candidate docs to PDF for visual comparison
