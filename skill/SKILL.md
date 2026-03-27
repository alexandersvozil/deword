---
name: deword
description: Read and extract content from Word documents (.docx, .doc). Converts Word files to clean markdown with images, tables, and formatting preserved. Use when you need to read, inspect, or extract content from any .docx or .doc file.
---

# deword — Word Document Reader

Converts Word documents into clean markdown for AI agents. One command, full document content.

## Setup

Install once (requires Node.js 18+ or Homebrew):

```bash
# Option A: Homebrew (macOS/Linux)
brew install alexandersvozil/tap/deword

# Option B: npm global install
npm install -g deword
```

Verify:
```bash
deword --version
```

## Reading documents

```bash
deword read document.docx              # markdown to stdout
deword read document.docx -f json      # structured JSON output
deword read document.docx -f summary   # metadata + preview only
deword read document.docx -i ./imgs    # custom image output directory
```

Images are automatically extracted to a temp directory. The markdown output includes absolute paths to extracted images.

## Deep inspection

```bash
deword unpack document.docx -o ./work
```

Creates a working directory with:
- `CONTENT.md` — full markdown content
- `content.json` — structured JSON
- `xml/` — prettified source XML
- `media/` — extracted images
- `manifest.json` — document stats

## Supported formats

- `.docx` (Word 2007+) — full support
- `.doc` (MHTML/Web Page) — full support  
- `.doc` (legacy binary) — clear error with guidance

Format is detected by magic bytes, not file extension.

## Output features

- Headings with proper `#` levels
- **Bold** and *italic* formatting
- Markdown tables
- Image extraction with paths
- Hyperlinks preserved
- YAML frontmatter with metadata (title, author, dates)
- Fragmented spell-check runs merged automatically
