# 🪱 deword

**De-Words your documents for AI agents.**

Word docs are the worst format for LLMs — bloated XML, fragmented text, images trapped in ZIP archives. `deword` rips that away and gives your agent clean markdown + images.

## Quick start

```bash
npx deword read report.docx
```

That's it. Clean markdown to stdout. Images auto-extracted to a temp dir with paths in the output so your agent can view them too.

## What you get

- **Markdown** with headings, **bold**, *italic*, tables, links preserved
- **Images** auto-extracted to `$TMPDIR/deword/<file>/images/` — no flag needed
- **Tables** as proper markdown tables
- **Metadata** (title, author, dates) in YAML frontmatter
- Fragmented Word runs merged (spell-check splits `"Hello"` across 3 XML elements — we fix that)

## Formats

| Input | Status |
|---|---|
| `.docx` (Word 2007+) | ✅ Full support |
| `.doc` (MHTML/Web Page) | ✅ Full support |
| `.doc` (Legacy binary) | ❌ Clear error |

Detection is by magic bytes, not extension.

## Options

```bash
deword read report.docx              # markdown (default)
deword read report.docx -f json      # structured JSON
deword read report.docx -f summary   # metadata + preview
deword read report.docx -i ./imgs    # custom image output dir
```

## Why not just `python-docx`?

| | Raw XML | python-docx script | deword |
|---|---|---|---|
| Tokens (1.1MB docx) | 198,613 | 5,600 | 6,255 |
| Formatting | ❌ | ❌ | ✅ |
| Tables | ❌ | Bare text | ✅ Markdown |
| Images | ❌ | ❌ | ✅ Auto-extracted |
| Round-trips | 0 | 2-3 | 1 |

Similar token count, but you keep formatting, tables, and images — and it's one tool call instead of three.

## License

MIT
