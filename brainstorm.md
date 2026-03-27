# Brainstorm: Agent-Friendly Word Document Editing via deword

## The Problem

Agents (Pi, Claude, etc.) edit text files brilliantly. Pi's `edit` tool is dead simple:

```
edit(path, oldText="Hello world", newText="Hello universe")
```

Find exact text → replace it → done. No corruption, no format loss. It works because text files are what they are — plain text.

Word documents are the opposite. They're ZIP archives full of XML where a simple sentence like "Hello world" might be:

```xml
<w:r><w:rPr><w:b/></w:rPr><w:t>Hel</w:t></w:r>
<w:r><w:rPr><w:b/></w:rPr><w:t>lo wor</w:t></w:r>
<w:r><w:rPr><w:b/></w:rPr><w:t>ld</w:t></w:r>
```

Three runs. Same formatting. Split by spell-check tracking. An agent doing `sed` or string replace on this is doomed.

## What We Have Today

### deword (read-only)
- `deword read` → clean markdown
- `deword unpack` → markdown + JSON + prettified XML + images
- Already merges fragmented runs for clean reading (`mergeAdjacentRuns` in `utils/xml.ts`)
- Understands document structure: paragraphs, tables, SDTs, images, links

### docx-edit skill (current approach)
- 200+ line instruction manual teaching agents to write Python + lxml
- Full of "golden rules" because there are a dozen ways to corrupt a docx
- Agent must: unzip → parse XML → navigate namespaces → modify nodes → fix XML declarations → repack ZIP
- Works, but agents frequently mess up: inline/block SDT confusion, missing `showingPlcHdr` removal, wrong numbering IDs, broken XML declarations
- **Not satisfying because:** The agent is doing plumbing, not editing

## The Insight

Pi's `edit` tool doesn't ask the agent to understand UTF-8 encoding, file system inodes, or disk sectors. It just says: *"give me the old text and the new text."*

We need the same thing for Word docs. **deword should handle the plumbing.**

---

## Approach 1: `deword edit` — The Obvious One

```bash
deword edit document.docx --old "Hello world" --new "Hello universe"
deword edit document.docx --old "Hello world" --new "Hello universe" -o output.docx
```

### How it works internally
1. Parse `word/document.xml`
2. Merge adjacent runs (deword already does this!) to get contiguous text
3. Find the target text across the (now-merged) runs
4. Replace it, preserving the formatting of the first run
5. Repack the ZIP byte-for-byte (only `word/document.xml` changes)

### What deword already has for this
- `mergeAdjacentRuns()` solves the fragmentation problem
- `unpack()` reads all ZIP contents
- Just need a repacker + text-level find/replace on the merged XML

### Challenges
- **Ambiguity:** What if "Hello world" appears twice? → Error with context, like Pi's edit tool does
- **Cross-element text:** Text spanning paragraphs, table cells, SDTs → Scope replacement to paragraph level? Cell level?
- **Formatting preservation:** If replacement is longer/shorter, inherit formatting from the matched run(s)
- **Merged runs ≠ original runs:** We merge for searching but need to produce valid XML for the output. Do we write back merged runs (safe, simpler) or try to preserve original run boundaries (complex, fragile)?

### Verdict
🟢 **High value, moderate complexity.** This covers 80% of use cases (fixing typos, filling in names/dates, changing values). The run-merging is already done.

---

## Approach 2: `deword fill` — Form-Aware Editing

```bash
deword fill document.docx --field "Employee Name" --value "John Smith"
deword fill document.docx --field "I agree" --check
deword fill document.docx --fields fields.json
```

Where `fields.json` is:
```json
{
  "Employee Name": "John Smith",
  "Start Date": "2025-01-15",
  "I agree to terms": { "type": "checkbox", "checked": true },
  "Department": "Engineering"
}
```

### How it works
1. Find all SDTs (content controls) in the document
2. Match them by tag name, alias, or placeholder text
3. Fill them following all the golden rules (inline vs block, `showingPlcHdr`, placeholder styling)
4. Handle checkboxes with `w14:checkbox`

### What this replaces
The entire docx-edit skill's SDT section. All those rules about inline vs block, `showingPlcHdr`, placeholder styling — baked into deword, not taught to the agent.

### Bonus: `deword fields`
```bash
deword fields document.docx          # list all fillable fields
deword fields document.docx -f json  # structured output
```

Output:
```
Fields in document.docx:
  1. "Employee Name" (text, inline) — placeholder: "Click here to enter text"
  2. "Start Date" (date, inline) — placeholder: "Select date"
  3. "I agree to terms" (checkbox) — unchecked
  4. "Department" (dropdown) — options: Engineering, Sales, HR
```

Agent sees available fields → fills them by name. No XML navigation needed.

### Verdict
🟢 **Very high value for form-filling use cases.** The docx-edit skill exists precisely because this is hard. Encapsulating it in a command is the whole point.

---

## Approach 3: `deword replace` — Regex/Pattern Replace

```bash
deword replace document.docx --pattern "{{NAME}}" --value "John Smith"
deword replace document.docx --patterns replacements.json
```

Template-style replacement. Useful for mail-merge-like scenarios. Same as Approach 1 but optimized for multiple replacements and pattern matching.

### Verdict
🟡 **Nice to have.** This is Approach 1 with a batch mode. Could be a flag on `deword edit` instead of a separate command.

---

## Approach 4: `deword patch` — Markdown Round-Trip

The dream:
```bash
deword read doc.docx > content.md
# Agent edits content.md with Pi's normal edit tool
deword patch doc.docx content.md -o updated.docx
```

Agent works in markdown (its native habitat), deword maps changes back to XML.

### Why this is hard
- Markdown → XML mapping is lossy. You lose run-level formatting, styles, spacing, etc.
- A paragraph in markdown could map to many XML elements
- Adding new content means generating new XML nodes with correct styles
- Deletions require removing the right XML and fixing references

### Could work for a subset
- Text-only changes within existing paragraphs (no structural changes)
- Compute a diff between original markdown and edited markdown
- Map each diff hunk to the corresponding XML paragraph
- Apply text changes at the XML level using Approach 1's mechanism

### Verdict
🔴 **High complexity, fragile.** The lossy round-trip makes this unreliable. Maybe a future v2 thing if Approach 1 proves solid.

---

## Approach 5: `deword xml` — Safe XML Escape Hatch

```bash
deword xml document.docx                          # print document.xml (prettified)
deword xml document.docx --set document.xml < modified.xml   # write back
deword xml document.docx --set document.xml -i modified.xml  # same, from file
```

For when agents need to do something deword doesn't support (add a table, insert an image, etc.), give them a safe way to extract → modify → repack XML without the ZIP plumbing.

### What deword handles
- Byte-for-byte copy of all non-modified files
- XML declaration fix (`'` → `"`)
- Validation that the XML is well-formed before repacking
- Optional: basic OOXML structural validation (paragraphs inside body, runs inside paragraphs, etc.)

### Verdict
🟡 **Good escape hatch.** Low effort to build, covers edge cases. Pairs well with the docx-edit skill (which already teaches agents XML editing — this just removes the zip/repack footgun).

---

## Approach 6: Hybrid Skill — `deword` Commands + Minimal Guidance

Instead of a 200-line skill teaching XML surgery, a slim skill that says:

```markdown
# Editing Word Documents

## Replace text
deword edit doc.docx --old "old text" --new "new text"

## Fill form fields
deword fields doc.docx                    # see what's fillable
deword fill doc.docx --fields data.json   # fill it

## Check checkboxes
deword fill doc.docx --field "I agree" --check

## Advanced (XML-level editing)
deword xml doc.docx > doc.xml             # extract
# edit doc.xml with Pi's edit tool
deword xml doc.docx --set doc.xml         # repack safely

## Read back to verify
deword read doc.docx
```

That's a 20-line skill instead of 200 lines. The agent doesn't need to know about namespaces, run fragmentation, SDT contexts, or ZIP repacking.

### Verdict
🟢🟢 **This is the goal.** Everything above leads here.

---

## Priority Ranking

| # | Feature | Impact | Effort | Priority |
|---|---------|--------|--------|----------|
| 1 | `deword edit --old/--new` | Covers 80% of text edits | Medium | 🔥 Do first |
| 2 | `deword fields` (list fields) | Enables form filling | Low | 🔥 Do first |
| 3 | `deword fill` (fill fields/checkboxes) | Replaces most of docx-edit skill | Medium | 🔥 Do first |
| 4 | `deword xml` (safe extract/repack) | Escape hatch for everything else | Low | 👍 Do second |
| 5 | `deword replace` (batch/pattern) | Convenience over `edit` | Low | 👍 Do second |
| 6 | `deword patch` (markdown round-trip) | Dream feature | Very High | 🔮 Future |

---

## Key Design Decisions to Make

### 1. In-place or copy?
- Default: write to new file (`-o output.docx`)?
- Or default: edit in-place (like `sed -i`)?
- Pi's `edit` tool edits in-place. Probably should match that convention.
- Compromise: in-place by default, `-o` for explicit output.

### 2. Merged runs in output?
- deword merges fragmented runs for reading. Should edits produce merged runs?
- **Yes, for the replaced region.** If "Hel" + "lo wor" + "ld" → replace with "Hello universe", output one run: `<w:r><w:rPr>...</w:rPr><w:t>Hello universe</w:t></w:r>`
- Runs outside the replacement stay untouched (byte-for-byte preservation)

### 3. Scope of text matching?
- Per-paragraph? Per-cell? Per-document?
- Probably per-paragraph for safety. Cross-paragraph replacement is risky (structural changes).
- Flag for cross-paragraph? `--scope paragraph|cell|document`

### 4. What about headers/footers?
- `word/header1.xml`, `word/footer1.xml` etc. are separate files
- `deword edit` should search all text-bearing XML files, not just `document.xml`
- Or have `--scope` include header/footer

### 5. Error handling philosophy
- **Fail loudly.** If old text isn't found → error. If it's ambiguous → error with context.
- **Never produce a corrupt file.** Validate before writing.
- **Always keep original.** If in-place editing, maybe keep `.docx.bak`?

---

## What deword Already Has That Makes This Feasible

| Need | Already exists |
|------|---------------|
| Parse ZIP → XML | `unpack.ts` |
| Merge fragmented runs | `mergeAdjacentRuns()` in `utils/xml.ts` |
| Extract text from paragraphs | `extractContent()` in `extractors/content.ts` |
| Handle relationships/media | `unpack()` returns full structure |
| Pretty-print XML | `prettyPrintXml()` |
| Detect file format | `detectFileType()` |

**What's missing:**
- ZIP repacker (copy-all-replace-one pattern)
- Text → XML position mapping (know *where* in the XML a matched text lives)
- SDT field discovery and manipulation
- CLI commands for edit/fill/fields/xml

---

## Summary

The docx-edit skill is a band-aid: it teaches the agent to be a Word XML surgeon. That's like teaching an agent about file system inodes instead of giving it `edit()`.

**deword should be the `edit()` for Word documents.** Read is already solved. Now solve write — not full authoring, just *safe, targeted modifications*. The agent says what to change. deword handles the XML/ZIP plumbing.

Simplest possible interface, maximum format preservation, loud failures on ambiguity.
