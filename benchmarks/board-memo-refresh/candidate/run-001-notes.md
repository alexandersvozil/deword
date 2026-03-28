# Run 001

## Input
- `../input/board-memo-input.docx`

## Candidate output
- `board-memo-output.docx`

## Method
Batch text replacement with `deword replace --map replacements.json`

## Command
```bash
cp benchmarks/board-memo-refresh/input/board-memo-input.docx \
  benchmarks/board-memo-refresh/candidate/board-memo-output.docx

deword replace benchmarks/board-memo-refresh/candidate/board-memo-output.docx \
  --map benchmarks/board-memo-refresh/candidate/replacements.json
```

## Observed result
- Body content matches gold semantically
- Header updated
- Footer updated
- Footnote date updated
- Markdown diff vs gold only differed in metadata timestamps
- XML diffs vs gold were structural/run-level only, not semantic

## Next manual check
Open `board-memo-output.docx` in Microsoft Word and export:
- `board-memo-output.pdf`

Then compare visually against:
- `../gold/board-memo-gold.pdf`
