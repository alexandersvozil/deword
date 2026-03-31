---
name: deword
description: Create, read, edit, and fill Word documents (.docx, .doc). Converts to markdown, replaces text, fills form fields, checks checkboxes — all without breaking formatting. Use when you need to create, read, modify, fill out, or update any Word document.
---

# deword — Word Document Reader & Editor

## Setup

```bash
brew install alexandersvozil/tap/deword   # macOS/Linux
npm install -g deword                      # any platform with Node.js
```

## Create a new document

```bash
deword new document.docx
deword new document.docx --text "Project kickoff notes"
```

New documents are created from a bundled template. By default they contain a single `<empty>` placeholder so you can immediately replace it with `edit`, `replace`, or `patch`.

## Read a document

```bash
deword read document.docx              # markdown to stdout
deword read document.docx -f json      # structured JSON output
deword read document.docx -f summary   # metadata + preview only
```

## Edit text (like Pi's edit tool)

```bash
deword edit document.docx --old "old text" --new "new text"
```

- **In-place** by default (use `-o output.docx` to write elsewhere)
- Text must be **unique** — errors if found 0 or >1 times
- Handles Word's fragmented runs automatically (spell-check splits, etc.)
- Searches body, headers, footers, footnotes, endnotes

## Replace text (batch/template mode)

```bash
deword replace document.docx --old "{{NAME}}" --new "John Smith"
deword replace document.docx --map replacements.json
```

Unlike `edit`, replaces **ALL** occurrences. JSON format:
```json
{"{{NAME}}": "John Smith", "{{DATE}}": "2025-01-15", "{{COMPANY}}": "Acme"}
```

## List form fields

```bash
deword fields document.docx            # human-readable list
deword fields document.docx -f json    # structured JSON
```

Shows all SDT content controls: name, type, current value, placeholder text.

## Fill form fields

```bash
deword fill document.docx --field "Employee Name" --value "John Smith"
deword fill document.docx --field "I agree" --check
deword fill document.docx --field "I agree" --uncheck
deword fill document.docx --json fields.json
```

JSON format for batch fill:
```json
{
  "Employee Name": "John Smith",
  "Start Date": "2025-01-15",
  "I agree to terms": true,
  "Department": "Engineering"
}
```

String values fill text fields. Boolean values check/uncheck checkboxes.

## XML escape hatch (advanced)

When you need to do something deword doesn't support directly:

```bash
deword xml document.docx                          # show word/document.xml
deword xml document.docx -p word/styles.xml       # show specific file
deword xml document.docx --list                   # list all files in ZIP
deword xml document.docx --set word/document.xml -i modified.xml  # replace file
```

The XML commands handle ZIP repacking safely — validates XML before writing.

## Workflow: create/read → edit → verify

```bash
deword new document.docx                          # 0. create a new file when needed
deword read document.docx                         # 1. see what's there
deword edit document.docx --old "..." --new "..."   # 2. make changes
deword read document.docx                         # 3. verify
```

## Supported formats

- `.docx` (Word 2007+) — full read + edit support
- `.doc` (MHTML/Web Page) — read only
- `.doc` (legacy binary) — clear error with guidance
