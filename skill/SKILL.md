---
name: deword
description: Create, read, inspect, edit, fill, patch, unpack, and add real Word equations to Word documents (.docx and supported .doc) with the deword CLI. Use when you need to create or modify Word files without breaking formatting.
license: MIT
compatibility: Requires the deword CLI (installed globally or available via npx) and shell access.
---

# deword

Use `deword` for Word documents instead of raw ZIP/XML parsing or ad-hoc scripts whenever possible.

## First-time agent workflow

If this is your first time seeing `deword`, do this first:

```bash
deword help
```

Then choose one of these safe starting points.

### Existing document

```bash
deword read <file> -f summary
deword read <file> -f model
```

### New document

```bash
deword new <file>
deword read <file> -f summary
```

Always re-read after changes:

```bash
deword read <file>
# or
deword read <file> -f summary
```

## Command chooser

| Need | Command |
|---|---|
| Create a new `.docx` | `deword new <file>` |
| Inspect a document quickly | `deword read <file> -f summary` |
| Inspect full content | `deword read <file>` |
| Get paragraph/table IDs for structured work | `deword read <file> -f model` |
| Insert a real Word equation | `deword formula <file> ...` |
| Replace one exact unique string | `deword edit <file> --old ... --new ...` |
| Replace all occurrences | `deword replace <file> --old ... --new ...` or `--map` |
| List form fields | `deword fields <file>` |
| Fill fields / checkboxes | `deword fill <file> ...` |
| Do multi-step structural edits | `deword patch <file> -p patch.json` |
| Unpack for deep inspection | `deword unpack <file>` |
| Use raw XML as an escape hatch | `deword xml <file> ...` |

## Installation / availability check

```bash
command -v deword || npx deword --version
```

If needed:

```bash
brew install alexandersvozil/tap/deword
# or
npm install -g deword
```

## Recommended workflow

### 1) Inspect first

Quick overview:

```bash
deword read <file> -f summary
```

Normal reading:

```bash
deword read <file>
```

Structured agent work:

```bash
deword read <file> -f model
```

Use `model` when you expect multi-step edits or need stable IDs for paragraphs and tables.

### 2) Choose the right edit mode

#### A. Create a new document

```bash
deword new document.docx
deword new document.docx --text "Project kickoff notes"
```

New documents come from a bundled template. By default they contain a single `<empty>` placeholder so you can immediately target it with `edit`, `replace`, `formula`, or `patch`.

#### B. Insert a real Word equation

```bash
deword formula document.docx --replace "<empty>" --latex "E = mc^2"
deword formula document.docx --after "Financial Model" --latex "\\frac{a+b}{c}"
```

Important:
- this inserts a **real Microsoft Word equation** in **OMML / Math Mode**
- use this instead of raw XML whenever you need math formulas
- current focus is **displayed equation paragraphs**

Supported syntax is a practical LaTeX-style subset, including:
- `x^2`, `x_1`, `x_1^2`
- `\\frac{a+b}{c}`
- `\\sqrt{x}`, `\\sqrt[n]{x}`
- `\\sin(x)`, `\\log(x)`
- `\\alpha`, `\\beta`, `\\pi`, `\\leq`, `\\geq`, `\\neq`, `\\infty`

Location modes:
- `--replace "..."`
- `--after "..."`
- `--before "..."`
- `--append`

#### C. Single exact replacement

Use when one unique string should change:

```bash
deword edit <file> --old "exact old text" --new "new text"
```

Notes:
- edits in-place by default
- use `-o <output.docx>` to write a copy
- if match count is 0 or >1, deword fails; add more context and retry

#### D. Replace all occurrences / template mode

Use for placeholders or global replacements:

```bash
deword replace <file> --old "{{NAME}}" --new "John Smith"
deword replace <file> --map replacements.json
```

Map format:

```json
{
  "{{NAME}}": "John Smith",
  "{{DATE}}": "2026-03-31"
}
```

#### E. Fill form fields / checkboxes

List fields first:

```bash
deword fields <file>
deword fields <file> -f json
```

Fill them:

```bash
deword fill <file> --field "Employee Name" --value "John Smith"
deword fill <file> --field "I agree" --check
deword fill <file> --json fields.json
```

Batch JSON format:

```json
{
  "Employee Name": "John Smith",
  "Start Date": "2026-03-31",
  "I agree": true
}
```

#### F. Multi-step structured edits

For complex edits, use `patch`:

```bash
deword read <file> -f model
deword patch <file> -p patch.json -o updated.docx
```

Minimal patch example:

```json
{
  "version": "1.0",
  "operations": [
    {
      "op": "replace_text",
      "target": { "by_text": { "text": "Draft Report", "match": "exact" } },
      "new_text": "Final Report"
    }
  ]
}
```

Common patch ops:
- `replace_text`
- `replace_all`
- `edit_paragraph`
- `insert_paragraph`
- `insert_footnote`
- `edit_footnote`
- `insert_table`
- `edit_table_cell`
- `append_table_row`
- `insert_image`
- `set_region_text`
- `fill_field`
- `set_checkbox`

### 3) Verify after changes

Always re-read after editing:

```bash
deword read <file>
# or
deword read <file> -f summary
```

When layout matters, recommend a final visual check in Word.

## Output formats

```bash
deword read <file>                 # markdown
deword read <file> -f json         # structured JSON
deword read <file> -f summary      # metadata + preview
deword read <file> -f model        # agent-friendly structure
```

Images can be extracted to a specific directory:

```bash
deword read <file> -i ./images
```

## Deep inspection / escape hatches

### Unpack a document

```bash
deword unpack <file> -o ./docx-work
```

Useful for examining generated markdown, JSON, XML, media, and manifest files.

### Work directly with internal XML

```bash
deword xml <file>
deword xml <file> --list
deword xml <file> -p word/styles.xml
deword xml <file> --set word/document.xml -i modified.xml -o output.docx
```

Use XML mode only when higher-level commands cannot express the change.

## Supported formats

- `.docx` (Word 2007+): full read + edit support
- `.doc` (MHTML/Web Page): read only
- legacy binary `.doc`: not supported

## Rules of thumb for agents

- Prefer `deword` over manual XML parsing for Word tasks.
- Start with `deword help` if unsure.
- Start with `read -f summary` or `read -f model` before editing.
- Use `new` when no document exists yet.
- Use `formula` whenever the user wants a real Word math equation.
- Use `edit` for a single unique replacement.
- Use `replace` for placeholders and bulk replacements.
- Use `fields` + `fill` for forms.
- Use `patch` for multi-step or structural edits.
- Use `-o` to avoid modifying originals unless the user clearly wants in-place edits.
- Re-read to verify semantic correctness after each change.
- If a command fails due to ambiguous text, use a more specific target.
- Use `xml` only as a last resort.

## If invoked with /skill:deword <file>

Treat the appended `User: <file>` text as the Word document path.
Immediately inspect it with:

```bash
deword read <file> -f summary
```

Then choose the appropriate follow-up workflow based on the user's task.
