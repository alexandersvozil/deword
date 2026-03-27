# 🪱 deword

**De-Words your documents for AI agents.**

Word docs are the worst format for LLMs — bloated XML, fragmented text, images trapped in ZIP archives. `deword` rips that away and gives your agent clean markdown + images. And now it can **edit them too** — text replacement, form filling, batch replace — all without breaking formatting.

## Install

```bash
# Homebrew (macOS/Linux)
brew install alexandersvozil/tap/deword

# npm (any platform with Node.js 18+)
npm install -g deword

# Or just run it directly — no install needed
npx deword read report.docx
```

## Quick start

```bash
# Read
deword read report.docx

# Edit (like Pi's edit tool — unique text, in-place)
deword edit report.docx --old "Draft Report" --new "Final Report"

# Replace all occurrences (batch/template mode)
deword replace report.docx --map replacements.json

# Fill form fields
deword fill form.docx --field "Employee Name" --value "John Smith"
deword fill form.docx --field "I agree" --check
```

## Reading

```bash
deword read report.docx              # markdown (default)
deword read report.docx -f json      # structured JSON
deword read report.docx -f summary   # metadata + preview
deword read report.docx -i ./imgs    # custom image output dir
```

- **Markdown** with headings, **bold**, *italic*, tables, links preserved
- **Images** auto-extracted to `$TMPDIR/deword/<file>/images/`
- **Tables** as proper markdown tables
- **Metadata** (title, author, dates) in YAML frontmatter
- Fragmented Word runs merged (spell-check splits `"Hello"` across 3 XML elements — we fix that)

## Editing

### `edit` — surgical single replacement

```bash
deword edit report.docx --old "Q4 2025" --new "Q4 FY2025"
deword edit report.docx --old "Draft" --new "Final" -o final.docx
```

Like Pi's `edit` tool: the old text must match **exactly once** across the entire document (body, headers, footers, footnotes). If it matches 0 or >1 times, you get a clear error telling you to add more context. Edits in-place by default.

### `replace` — batch/template replacement

```bash
deword replace report.docx --old "{{NAME}}" --new "John Smith"
deword replace report.docx --map replacements.json
```

Replaces **all** occurrences. JSON map format:

```json
{"{{NAME}}": "John Smith", "{{DATE}}": "2025-01-15", "{{COMPANY}}": "Acme"}
```

### `fill` — form fields and checkboxes

```bash
deword fields form.docx                              # list all fields
deword fields form.docx -f json                      # structured JSON
deword fill form.docx --field "Name" --value "John"  # fill text field
deword fill form.docx --field "I agree" --check      # check checkbox
deword fill form.docx --json data.json               # batch fill
```

Batch JSON format (strings for text, booleans for checkboxes):

```json
{"Employee Name": "John Smith", "Start Date": "2025-01-15", "I agree": true}
```

### `xml` — safe escape hatch

For anything the other commands can't do, work with the raw XML directly:

```bash
deword xml report.docx                                    # show document.xml
deword xml report.docx -p word/styles.xml                 # show specific file
deword xml report.docx --list                             # list all ZIP contents
deword xml report.docx --set word/document.xml -i new.xml # replace a file
```

Handles ZIP repacking safely. Validates XML before writing.

## Formats

| Input | Read | Edit |
|---|---|---|
| `.docx` (Word 2007+) | ✅ Full | ✅ Full |
| `.doc` (MHTML/Web Page) | ✅ Full | ❌ |
| `.doc` (Legacy binary) | ❌ Clear error | ❌ |

Detection is by magic bytes, not extension.

## Agent Integration

deword is built to be used by AI coding agents. There are several ways to integrate it:

### AGENTS.md (automatic for many agents)

The repo includes an `AGENTS.md` file at the root. Agents that support this convention (Claude Code, Cursor, pi, etc.) will automatically pick up the instructions when working in the repo.

### Pi skill

Copy the `skill/` directory into your pi skills folder:

```bash
cp -r skill/ ~/.pi/agent/skills/deword
```

Or add to your project's `.pi/skills/` directory. Pi will auto-discover it and load the instructions when you work with Word files.

### Other agents

Point your agent at the `AGENTS.md` file, or simply tell it:

> Use `deword` to read and edit Word documents. Install with `npm install -g deword` if not available.

### npx (zero-install)

Any agent with shell access can use deword without pre-installing:

```bash
npx deword read document.docx
npx deword edit document.docx --old "foo" --new "bar"
```

## Why not just `python-docx`?

| | Raw XML | python-docx script | deword |
|---|---|---|---|
| Tokens (1.1MB docx) | 198,613 | 5,600 | 6,255 |
| Formatting | ❌ | ❌ | ✅ |
| Tables | ❌ | Bare text | ✅ Markdown |
| Images | ❌ | ❌ | ✅ Auto-extracted |
| Editing | Manual XML surgery | Fragile | ✅ One command |
| Round-trips | 0 | 2-3 | 1 |

## License

MIT
