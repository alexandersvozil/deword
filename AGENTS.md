# 🪱 deword — Agent Instructions

You have access to `deword`, a CLI tool that reads and edits Word documents (.docx, .doc) without breaking formatting.

## Installation check

```bash
command -v deword || npx deword --version
```

If not installed, install with one of:

```bash
brew install alexandersvozil/tap/deword   # macOS/Linux
npm install -g deword                      # any platform with Node.js
```

## Reading a Word document

```bash
deword read <file>                    # clean markdown to stdout
deword read <file> -f json            # structured JSON with metadata
deword read <file> -f summary         # quick metadata + preview
deword read <file> -i ./images        # extract images to specific dir
```

**Default behavior:** Images are auto-extracted to a temp directory. The markdown output contains absolute paths to the extracted images so you can view them with your image viewing tool.

## Editing text (like Pi's edit tool)

```bash
deword edit <file> --old "exact old text" --new "new text"
deword edit <file> --old "old" --new "new" -o output.docx  # write to different file
```

- In-place by default. Text must match exactly once (add context to disambiguate).
- Handles fragmented runs, XML encoding, headers/footers automatically.

## Replacing text (batch/template mode)

```bash
deword replace <file> --old "{{NAME}}" --new "John Smith"
deword replace <file> --map replacements.json
```

Replaces ALL occurrences. JSON: `{"{{NAME}}": "John", "{{DATE}}": "2025-01-15"}`

## Form fields

```bash
deword fields <file>                              # list all fields
deword fields <file> -f json                      # structured JSON
deword fill <file> --field "Name" --value "John"  # fill a text field
deword fill <file> --field "Agree" --check        # check a checkbox
deword fill <file> --json fields.json             # batch fill from JSON
```

JSON format: `{"Name": "John", "Agree": true, "Date": "2025-01-15"}`

## XML escape hatch

```bash
deword xml <file>                                 # show word/document.xml
deword xml <file> -p word/styles.xml              # show specific file
deword xml <file> --list                          # list all files in ZIP
deword xml <file> --set word/document.xml -i modified.xml  # replace safely
```

## Unpacking for deep inspection

```bash
deword unpack <file> -o ./work        # full extraction to working dir
```

Creates: `CONTENT.md`, `content.json`, `xml/` (prettified XML), `media/` (images), `manifest.json`

## What deword handles

| Feature | Details |
|---------|---------|
| Headings | Preserved with proper `#` levels |
| Bold/Italic | `**bold**`, `*italic*` |
| Tables | Proper markdown tables |
| Images | Auto-extracted with paths in output |
| Links | Preserved as `[text](url)` |
| Metadata | Title, author, dates in YAML frontmatter |
| Fragmented runs | Spell-check fragments merged automatically |
| Text editing | In-place, format-preserving, unique-match |
| Batch replace | Template/mail-merge style replacement |
| Form fields | SDT content controls: text, checkbox, dropdown |
| XML access | Safe extract/replace with ZIP handling |

## Supported formats

| Format | Read | Edit |
|--------|------|------|
| `.docx` (Word 2007+) | ✅ | ✅ |
| `.doc` (MHTML/Web Page) | ✅ | ❌ |
| `.doc` (Legacy binary) | ❌ | ❌ |

## Tips

- **Always use `deword read` first** to understand the document before editing
- Use `deword edit` for surgical single replacements (like Pi's `edit` tool)
- Use `deword replace` for batch/template replacements (replaces all occurrences)
- Use `deword fields` + `deword fill` for form documents
- Use `deword xml` as an escape hatch for anything the other commands can't do
- Image paths in output are absolute — use your file/image viewing tool to inspect them
