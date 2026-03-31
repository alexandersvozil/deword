# 🪱 deword

**Create, read, edit, fill, patch, and add real equations to Word documents — without breaking formatting.**

Word docs are awful for agents: ZIP archives, fragmented runs, hidden media, and XML everywhere. `deword` turns them into something agents can actually work with, and now also lets them create new `.docx` files and insert real Microsoft Word equations.

## Install

```bash
# Homebrew (macOS/Linux)
brew install alexandersvozil/tap/deword

# npm (any platform with Node.js 18+)
npm install -g deword

# Zero-install
npx deword --help
```

## First-time agent quick start

If you are seeing `deword` for the first time, start here:

```bash
deword help
deword new draft.docx
deword read draft.docx -f summary
deword formula draft.docx --replace "<empty>" --latex "E = mc^2"
deword read draft.docx
```

For an existing document:

```bash
deword read report.docx -f summary
deword read report.docx -f model
```

## Pick the right command

| Need | Command |
|---|---|
| Create a new `.docx` | `deword new <file>` |
| Inspect a document | `deword read <file>` |
| Get structure for agent work | `deword read <file> -f model` |
| Insert a real Word equation | `deword formula <file> ...` |
| Replace one exact unique string | `deword edit <file> --old ... --new ...` |
| Replace all occurrences / placeholders | `deword replace <file> --old ... --new ...` or `--map` |
| List form fields | `deword fields <file>` |
| Fill fields / checkboxes | `deword fill <file> ...` |
| Do multi-step structural edits | `deword patch <file> -p patch.json` |
| Inspect unpacked XML/media | `deword unpack <file>` |
| Use raw XML as an escape hatch | `deword xml <file> ...` |

## Quick start

```bash
# Create a new document from the bundled template
deword new report.docx

# Read
deword read report.docx

# Insert a real Word equation (OMML / Math Mode)
deword formula report.docx --replace "<empty>" --latex "E = mc^2"

# Edit one exact unique text occurrence
deword edit report.docx --old "Draft Report" --new "Final Report"

# Replace all occurrences / template placeholders
deword replace report.docx --map replacements.json

# Fill form fields
deword fill form.docx --field "Employee Name" --value "John Smith"
deword fill form.docx --field "I agree" --check

# Read an agent-friendly document model
deword read report.docx -f model

# Apply a high-level JSON patch plan
deword patch report.docx -p patch.json -o final.docx
```

## Creating new documents

```bash
deword new report.docx
deword new report.docx --text "Quarterly Board Memo"
deword create drafts/note.docx
```

Creates a new `.docx` from the bundled template.

By default the document contains a single `<empty>` placeholder so agents can immediately target it with:
- `deword edit`
- `deword replace`
- `deword formula`
- `deword patch`

Use `--text` to replace the placeholder during creation.

## Formulas / equations

```bash
deword formula report.docx --replace "<empty>" --latex "E = mc^2"
deword formula report.docx --after "Financial Model" --latex "\\frac{Revenue - Cost}{Revenue}"
deword math report.docx --append --latex "x_1 + x_2 = y"
```

This inserts a **real Microsoft Word equation** in **OMML / Math Mode**. The result opens in Word as an actual editable equation, not plain text and not hand-written XML.

### Supported syntax today

Practical LaTeX-style subset:

- superscripts/subscripts: `x^2`, `x_1`, `x_1^2`
- fractions: `\\frac{a+b}{c}`
- radicals: `\\sqrt{x}`, `\\sqrt[n]{x}`
- functions: `\\sin(x)`, `\\log(x)`
- common Greek letters / symbols: `\\alpha`, `\\beta`, `\\pi`, `\\leq`, `\\geq`, `\\neq`, `\\infty`

### Location modes

- `--replace "..."` replace a uniquely matched body paragraph
- `--after "..."` insert after a uniquely matched body paragraph
- `--before "..."` insert before a uniquely matched body paragraph
- `--append` append to the end of the document body

### Current scope

`formula` currently inserts **displayed equation paragraphs**. It is ideal when you want a standalone formula block in Word. For more advanced layout or unsupported math syntax, use multiple formula insertions, `patch`, or `xml` as a last resort.

## Reading

```bash
deword read report.docx              # markdown (default)
deword read report.docx -f json      # structured JSON
deword read report.docx -f model     # agent-friendly structure with IDs
deword read report.docx -f summary   # metadata + preview
deword read report.docx -i ./imgs    # custom image extraction dir
```

What you get:
- markdown with headings, bold, italic, tables, and links preserved
- images auto-extracted to a temp directory by default
- metadata in YAML frontmatter
- merged fragmented Word runs for cleaner reading
- formula previews rendered as readable plain text in `read`

## Editing

### `edit` — surgical single replacement

```bash
deword edit report.docx --old "Q4 2025" --new "Q4 FY2025"
deword edit report.docx --old "Draft" --new "Final" -o final.docx
```

Use `edit` when the old text should match **exactly once** across the document. If it matches 0 or more than 1 times, deword fails and tells you to add more context.

Search scope includes:
- body
- headers
- footers
- footnotes
- endnotes

### `replace` — batch/template replacement

```bash
deword replace report.docx --old "{{NAME}}" --new "John Smith"
deword replace report.docx --map replacements.json
```

Replaces **all** occurrences.

JSON map format:

```json
{
  "{{NAME}}": "John Smith",
  "{{DATE}}": "2025-01-15",
  "{{COMPANY}}": "Acme"
}
```

### `fields` / `fill` — form fields and checkboxes

```bash
deword fields form.docx
deword fields form.docx -f json
deword fill form.docx --field "Name" --value "John"
deword fill form.docx --field "I agree" --check
deword fill form.docx --json data.json
```

Batch JSON format:

```json
{
  "Employee Name": "John Smith",
  "Start Date": "2025-01-15",
  "I agree": true
}
```

## `patch` — high-level agent editing plan

```bash
deword patch report.docx -p patch.json
deword patch report.docx -p patch.json -o final.docx
```

Use `patch` for multi-step or structural edits when `edit` or `replace` is too narrow.

Example:

```json
{
  "version": "1.0",
  "operations": [
    {
      "op": "replace_text",
      "target": { "by_text": { "text": "Draft Report", "match": "exact" } },
      "new_text": "Final Report"
    },
    {
      "op": "insert_footnote",
      "anchor": { "by_text": { "text": "assumption", "match": "exact" } },
      "footnote_text": "Confirmed with finance on 2026-03-28."
    },
    {
      "op": "insert_table",
      "location": { "after": { "by_heading": { "text": "Financial Summary", "level": 2 } } },
      "data": [["Metric", "Value"], ["Revenue", "$14.1M"]],
      "table_style": "professional"
    },
    {
      "op": "insert_image",
      "location": { "after": { "by_heading": { "text": "Market Overview", "level": 1 } } },
      "file_path": "chart.png",
      "width_px": 480,
      "alignment": "center",
      "caption": "Figure 1. Regional sales chart"
    }
  ]
}
```

Current patch ops:
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

## `unpack` — deep inspection

```bash
deword unpack report.docx -o ./docx-work
```

Creates agent-friendly files such as:
- `CONTENT.md`
- `content.json`
- `xml/`
- `media/`
- `manifest.json`

## `xml` — safe escape hatch

```bash
deword xml report.docx
deword xml report.docx -p word/styles.xml
deword xml report.docx --list
deword xml report.docx --set word/document.xml -i new.xml
```

Use XML mode only when higher-level commands cannot express the change.

## Recommended agent workflow

### Existing document

```bash
deword read <file> -f summary
deword read <file> -f model
# choose edit / replace / fill / formula / patch
deword read <file>
```

### New document

```bash
deword new <file>
deword formula <file> --replace "<empty>" --latex "E = mc^2"
# or deword edit / replace / patch
deword read <file>
```

## Formats

| Input | Read | Edit |
|---|---|---|
| `.docx` (Word 2007+) | ✅ Full | ✅ Full |
| `.doc` (MHTML/Web Page) | ✅ Full | ❌ |
| `.doc` (Legacy binary) | ❌ Clear error | ❌ |

Detection is by magic bytes, not extension.

## Agent integration

### `deword help`

The CLI help is designed to be enough for a first-time agent:

```bash
deword help
deword help formula
deword help patch
```

### `AGENTS.md`

This repo includes an `AGENTS.md` file at the root. Agents that support that convention can pick it up automatically.

### Pi skill

Copy the `skill/` directory into your Pi skills folder:

```bash
cp -r skill/ ~/.pi/agent/skills/deword
```

Or add it to your project’s `.pi/skills/` directory.

### Zero-install via `npx`

```bash
npx deword read document.docx
npx deword edit document.docx --old "foo" --new "bar"
```

## Why not just `python-docx`?

| | Raw XML | python-docx script | deword |
|---|---|---|---|
| Tokens (1.1MB docx) | 198,613 | 5,600 | 6,255 |
| Formatting preserved | ❌ | ❌ | ✅ |
| Tables | ❌ | Bare text | ✅ Markdown |
| Images | ❌ | ❌ | ✅ Auto-extracted |
| Real Word equations | ❌ | Fragile/manual | ✅ |
| Editing | Manual XML surgery | Fragile | ✅ One command |
| Round-trips | 0 | 2-3 | 1 |

## License

MIT
