# 🪱 deword

**De-Words your documents for AI agents.**

Word docs are the worst format for LLMs — bloated XML, fragmented text, images trapped in ZIP archives. `deword` rips that away and gives your agent clean markdown + images.

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
deword read report.docx
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

> Use `deword read <file>` to read Word documents. Install with `npm install -g deword` if not available.

### npx (zero-install)

Any agent with shell access can use deword without pre-installing:

```bash
npx deword read document.docx
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

## Setting up the Homebrew tap

If you want to host your own tap (for `brew install alexandersvozil/tap/deword`):

1. Create a GitHub repo called `homebrew-tap` under your account
2. Copy `homebrew/deword.rb` into `Formula/deword.rb` in that repo
3. After publishing to npm, update the `sha256` in the formula:
   ```bash
   curl -sL https://registry.npmjs.org/deword/-/deword-0.1.0.tgz | shasum -a 256
   ```
4. Users can then: `brew tap alexandersvozil/tap && brew install deword`

## Publishing to npm

```bash
npm login
npm publish
```

After publishing, `npx deword` and `npm install -g deword` will work for everyone.

## License

MIT
