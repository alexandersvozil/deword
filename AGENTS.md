# 🪱 deword — Agent Instructions

You have access to `deword`, a CLI tool that converts Word documents (.docx, .doc) into clean markdown that you can read and reason about.

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

## Unpacking for deep inspection

```bash
deword unpack <file> -o ./work        # full extraction to working dir
```

This creates:
- `CONTENT.md` — full markdown
- `content.json` — structured JSON
- `xml/` — prettified XML files (for debugging formatting)
- `media/` — extracted images
- `manifest.json` — stats and file listing

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

## Supported formats

| Format | Support |
|--------|---------|
| `.docx` (Word 2007+) | ✅ Full |
| `.doc` (MHTML/Web Page) | ✅ Full |
| `.doc` (Legacy binary) | ❌ Error with explanation |

Detection is by magic bytes, not file extension.

## Tips

- **Always use `deword read` first** — it gives you everything you need in one call
- Use `-f summary` when you only need to know what's in a document before reading it fully
- Image paths in the output are absolute — use your file/image viewing tool to inspect them
- For editing .docx files, read the document with deword first to understand its structure, then use XML-level editing (docx files are ZIP archives containing XML)
