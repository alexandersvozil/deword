# 🪱 deword — Agent Instructions

You have access to `deword`, a CLI for creating, reading, editing, filling, patching, unpacking, and adding real Word equations to Word documents without breaking formatting.

## First thing to do if you are new to deword

```bash
deword help
```

If you already have a document:

```bash
deword read <file> -f summary
deword read <file> -f model
```

If you need a fresh document:

```bash
deword new <file>
```

Always re-read after changes:

```bash
deword read <file>
# or
deword read <file> -f summary
```

## Installation check

```bash
command -v deword || npx deword --version
```

If not installed, install with one of:

```bash
brew install alexandersvozil/tap/deword   # macOS/Linux
npm install -g deword                      # any platform with Node.js
```

## Command chooser

| Need | Command |
|---|---|
| Create a new `.docx` | `deword new <file>` |
| Inspect content quickly | `deword read <file> -f summary` |
| Read full content | `deword read <file>` |
| Get paragraph/table IDs for structured edits | `deword read <file> -f model` |
| Insert a real Word equation | `deword formula <file> ...` |
| Replace one exact unique string | `deword edit <file> --old ... --new ...` |
| Replace all occurrences / placeholders | `deword replace <file> ...` |
| List form fields | `deword fields <file>` |
| Fill fields / checkboxes | `deword fill <file> ...` |
| Do multi-step structural edits | `deword patch <file> -p patch.json` |
| Deep inspection | `deword unpack <file>` |
| Last-resort raw XML access | `deword xml <file> ...` |

## Creating a new Word document

```bash
deword new <file>
deword new <file> --text "Hello"
```

New documents are created from a bundled template. By default they contain a single `<empty>` placeholder so you can immediately target it with `edit`, `replace`, `formula`, or `patch`.

## Inserting a math formula / equation

```bash
deword formula <file> --replace "<empty>" --latex "E = mc^2"
deword formula <file> --after "Financial Model" --latex "\\frac{a+b}{c}"
```

This inserts a **real Word equation** in **OMML / Math Mode**, not plain text.

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

## Reading a Word document

```bash
deword read <file>                    # clean markdown to stdout
deword read <file> -f json            # structured JSON with metadata
deword read <file> -f summary         # quick metadata + preview
deword read <file> -f model           # paragraph/table IDs for agent work
deword read <file> -i ./images        # extract images to specific dir
```

Images are auto-extracted to a temp directory by default. The markdown output contains absolute paths to extracted images.

## Editing text

```bash
deword edit <file> --old "exact old text" --new "new text"
deword edit <file> --old "old" --new "new" -o output.docx
```

- In-place by default
- Text must match exactly once
- Handles fragmented runs, headers, footers, footnotes, endnotes

## Replacing text in batch

```bash
deword replace <file> --old "{{NAME}}" --new "John Smith"
deword replace <file> --map replacements.json
```

Replaces **all** occurrences.

Example JSON:

```json
{"{{NAME}}": "John", "{{DATE}}": "2025-01-15"}
```

## Form fields

```bash
deword fields <file>
deword fields <file> -f json
deword fill <file> --field "Name" --value "John"
deword fill <file> --field "Agree" --check
deword fill <file> --json fields.json
```

Batch JSON format:

```json
{"Name": "John", "Agree": true, "Date": "2025-01-15"}
```

## Multi-step structural edits

```bash
deword read <file> -f model
deword patch <file> -p patch.json
```

Use `patch` when you need more than a single text replacement.

## Unpacking for deep inspection

```bash
deword unpack <file> -o ./work
```

Creates agent-friendly artifacts such as `CONTENT.md`, `content.json`, `xml/`, `media/`, and `manifest.json`.

## XML escape hatch

```bash
deword xml <file>
deword xml <file> -p word/styles.xml
deword xml <file> --list
deword xml <file> --set word/document.xml -i modified.xml
```

Use XML mode only when higher-level commands cannot express the change.

## Supported formats

| Format | Read | Edit |
|---|---|---|
| `.docx` (Word 2007+) | ✅ | ✅ |
| `.doc` (MHTML/Web Page) | ✅ | ❌ |
| `.doc` (Legacy binary) | ❌ | ❌ |

## Rules of thumb

- Prefer `deword` over manual XML parsing for Word tasks.
- Start with `deword help` if unsure.
- Use `deword new` when you need a fresh `.docx`.
- Use `deword formula` when you need a real Word math equation.
- Use `deword read -f model` before structural edits.
- Use `deword edit` for one exact unique change.
- Use `deword replace` for placeholders and repeated text.
- Use `deword fields` + `deword fill` for forms.
- Use `deword patch` for multi-step edits.
- Use `deword xml` only as a last resort.
- Re-read after every change.
