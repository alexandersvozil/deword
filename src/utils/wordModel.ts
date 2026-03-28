import { loadDocx, getTextXmlPaths } from "./repack.js";
import { extractAllText, replaceTextInXml, xmlEncode, xmlDecode } from "./xml.js";
import { findSdtFields } from "./sdt.js";

export type RegionKind = "body" | "header" | "footer" | "footnote" | "endnote";

export interface ParagraphInfo {
  id: string;
  xmlPath: string;
  kind: RegionKind;
  sectionId: string;
  text: string;
  style: string | null;
  headingLevel: number | null;
  xml: string;
  start: number;
  end: number;
}

export interface ParagraphModelInfo {
  id: string;
  xmlPath: string;
  kind: RegionKind;
  sectionId: string;
  text: string;
  style: string | null;
  headingLevel: number | null;
}

export interface TableInfo {
  id: string;
  xmlPath: string;
  kind: RegionKind;
  sectionId: string;
  rows: string[][];
  xml: string;
  start: number;
  end: number;
}

export interface FootnoteInfo {
  id: string;
  noteId: string;
  text: string;
  xmlPath: string;
  xml: string;
  start: number;
  end: number;
}

export interface RegionInfo {
  id: string;
  xmlPath: string;
  kind: "header" | "footer";
  sectionId: string;
  text: string;
  paragraphIds: string[];
}

export interface FieldInfo {
  id: string;
  name: string;
  type: string;
  xmlPath: string;
  currentValue: string;
  isChecked?: boolean;
}

export interface WordModel {
  version: "1.0";
  document: {
    path: string;
    format: "docx";
    title: string | null;
    author: string | null;
    paragraphCount: number;
    tableCount: number;
    imageCount: number;
    sectionCount: number;
  };
  content: {
    text: string;
    markdown: string;
  };
  headings: Array<{ id: string; level: number; text: string; paragraphId: string }>;
  paragraphs: ParagraphModelInfo[];
  tables: Array<{ id: string; rows: string[][]; sectionId: string }>;
  images: Array<{ id: string; path: string; description: string; sizeBytes: number }>;
  footnotes: Array<{ id: string; noteId: string; text: string }>;
  headers: RegionInfo[];
  footers: RegionInfo[];
  fields: FieldInfo[];
}

interface BalancedTagRange {
  start: number;
  end: number;
  innerStart: number;
  innerEnd: number;
  openTag: string;
  closeTag: string;
  xml: string;
  innerXml: string;
}

interface BlockRange {
  tag: string;
  start: number;
  end: number;
  xml: string;
}

export async function buildWordModel(
  filePath: string,
  source: {
    markdown: string;
    images: Array<{ id: string; path: string; description: string; size: number }>;
    metadata: { title: string | null; author: string | null };
  }
): Promise<WordModel> {
  const { files } = await loadDocx(filePath);
  const paragraphs: ParagraphInfo[] = [];
  const tables: TableInfo[] = [];
  const footnotes: FootnoteInfo[] = [];
  const headers: RegionInfo[] = [];
  const footers: RegionInfo[] = [];
  const fields: FieldInfo[] = [];
  const headings: Array<{ id: string; level: number; text: string; paragraphId: string }> = [];

  let paragraphSeq = 0;
  let tableSeq = 0;
  let headerSeq = 0;
  let footerSeq = 0;
  let sectionSeq = 0;
  let fieldSeq = 0;

  const bodyXml = files.get("word/document.xml");
  if (bodyXml) {
    const body = getBalancedTagRange(bodyXml, "w:body");
    if (body) {
      sectionSeq++;
      for (const block of findTopLevelBlocks(body.innerXml, ["w:p", "w:tbl", "w:sdt"])) {
        if (block.tag === "w:p") {
          const p = makeParagraphInfo(
            `p${++paragraphSeq}`,
            "word/document.xml",
            "body",
            `s${sectionSeq}`,
            block.xml,
            body.innerStart + block.start,
            body.innerStart + block.end
          );
          paragraphs.push(p);
          if (p.headingLevel && p.text.trim()) {
            headings.push({ id: `h${headings.length + 1}`, level: p.headingLevel, text: p.text, paragraphId: p.id });
          }
        } else if (block.tag === "w:tbl") {
          tables.push({
            id: `t${++tableSeq}`,
            xmlPath: "word/document.xml",
            kind: "body",
            sectionId: `s${sectionSeq}`,
            rows: extractTableRows(block.xml),
            xml: block.xml,
            start: body.innerStart + block.start,
            end: body.innerStart + block.end,
          });
        } else if (block.tag === "w:sdt") {
          const innerBlocks = findTopLevelBlocks(block.xml, ["w:p", "w:tbl"]);
          for (const inner of innerBlocks) {
            if (inner.tag === "w:p") {
              const p = makeParagraphInfo(
                `p${++paragraphSeq}`,
                "word/document.xml",
                "body",
                `s${sectionSeq}`,
                inner.xml,
                body.innerStart + block.start + inner.start,
                body.innerStart + block.start + inner.end
              );
              paragraphs.push(p);
            } else if (inner.tag === "w:tbl") {
              tables.push({
                id: `t${++tableSeq}`,
                xmlPath: "word/document.xml",
                kind: "body",
                sectionId: `s${sectionSeq}`,
                rows: extractTableRows(inner.xml),
                xml: inner.xml,
                start: body.innerStart + block.start + inner.start,
                end: body.innerStart + block.start + inner.end,
              });
            }
          }
        }
      }
    }
  }

  for (const xmlPath of [...files.keys()].filter((p) => /^word\/header\d+\.xml$/.test(p)).sort()) {
    const xml = files.get(xmlPath)!;
    sectionSeq++;
    const regionParagraphs: string[] = [];
    const paragraphsInFile = collectRegionParagraphs(xml, xmlPath, "header", `s${sectionSeq}`, paragraphSeq);
    for (const p of paragraphsInFile) {
      paragraphSeq++;
      p.id = `p${paragraphSeq}`;
      regionParagraphs.push(p.id);
      paragraphs.push(p);
    }
    headers.push({
      id: `hdr${++headerSeq}`,
      xmlPath,
      kind: "header",
      sectionId: `s${sectionSeq}`,
      text: paragraphsInFile.map((p) => p.text).filter(Boolean).join("\n"),
      paragraphIds: regionParagraphs,
    });
  }

  for (const xmlPath of [...files.keys()].filter((p) => /^word\/footer\d+\.xml$/.test(p)).sort()) {
    const xml = files.get(xmlPath)!;
    sectionSeq++;
    const regionParagraphs: string[] = [];
    const paragraphsInFile = collectRegionParagraphs(xml, xmlPath, "footer", `s${sectionSeq}`, paragraphSeq);
    for (const p of paragraphsInFile) {
      paragraphSeq++;
      p.id = `p${paragraphSeq}`;
      regionParagraphs.push(p.id);
      paragraphs.push(p);
    }
    footers.push({
      id: `ftr${++footerSeq}`,
      xmlPath,
      kind: "footer",
      sectionId: `s${sectionSeq}`,
      text: paragraphsInFile.map((p) => p.text).filter(Boolean).join("\n"),
      paragraphIds: regionParagraphs,
    });
  }

  const footnotesXml = files.get("word/footnotes.xml");
  if (footnotesXml) {
    for (const range of findBalancedTags(footnotesXml, "w:footnote")) {
      const noteId = extractAttribute(range.openTag, "w:id");
      if (!noteId || noteId === "-1" || noteId === "0") continue;
      footnotes.push({
        id: `fn${noteId}`,
        noteId,
        text: extractAllText(range.xml).trim(),
        xmlPath: "word/footnotes.xml",
        xml: range.xml,
        start: range.start,
        end: range.end,
      });
    }
  }

  for (const xmlPath of getTextXmlPaths(files)) {
    const xml = files.get(xmlPath)!;
    for (const field of findSdtFields(xml, xmlPath)) {
      fields.push({
        id: `fld${++fieldSeq}`,
        name: field.tag ?? field.alias ?? field.placeholderText ?? `field-${fieldSeq}`,
        type: field.type,
        xmlPath,
        currentValue: field.currentValue,
        isChecked: field.type === "checkbox" ? field.isChecked : undefined,
      });
    }
  }

  return {
    version: "1.0",
    document: {
      path: filePath,
      format: "docx",
      title: source.metadata.title,
      author: source.metadata.author,
      paragraphCount: paragraphs.length,
      tableCount: tables.length,
      imageCount: source.images.length,
      sectionCount: sectionSeq,
    },
    content: {
      text: markdownToText(source.markdown),
      markdown: source.markdown,
    },
    headings,
    paragraphs: paragraphs.map(({ id, xmlPath, kind, sectionId, text, style, headingLevel }) => ({
      id,
      xmlPath,
      kind,
      sectionId,
      text,
      style,
      headingLevel,
    })),
    tables: tables.map((t) => ({ id: t.id, rows: t.rows, sectionId: t.sectionId })),
    images: source.images.map((img) => ({ id: img.id, path: img.path, description: img.description, sizeBytes: img.size })),
    footnotes: footnotes.map((f) => ({ id: f.id, noteId: f.noteId, text: f.text })),
    headers,
    footers,
    fields,
  };
}

export function buildMutableWordStateFromFiles(files: Map<string, string>): {
  files: Map<string, string>;
  paragraphs: ParagraphInfo[];
  tables: TableInfo[];
  footnotes: FootnoteInfo[];
  headers: RegionInfo[];
  footers: RegionInfo[];
} {
  const paragraphs: ParagraphInfo[] = [];
  const tables: TableInfo[] = [];
  const footnotes: FootnoteInfo[] = [];
  const headers: RegionInfo[] = [];
  const footers: RegionInfo[] = [];

  let paragraphSeq = 0;
  let tableSeq = 0;
  let headerSeq = 0;
  let footerSeq = 0;
  let sectionSeq = 0;

  const bodyXml = files.get("word/document.xml");
  if (bodyXml) {
    const body = getBalancedTagRange(bodyXml, "w:body");
    if (body) {
      sectionSeq++;
      for (const block of findTopLevelBlocks(body.innerXml, ["w:p", "w:tbl"])) {
        if (block.tag === "w:p") {
          paragraphs.push(
            makeParagraphInfo(
              `p${++paragraphSeq}`,
              "word/document.xml",
              "body",
              `s${sectionSeq}`,
              block.xml,
              body.innerStart + block.start,
              body.innerStart + block.end
            )
          );
        } else if (block.tag === "w:tbl") {
          tables.push({
            id: `t${++tableSeq}`,
            xmlPath: "word/document.xml",
            kind: "body",
            sectionId: `s${sectionSeq}`,
            rows: extractTableRows(block.xml),
            xml: block.xml,
            start: body.innerStart + block.start,
            end: body.innerStart + block.end,
          });
        }
      }
    }
  }

  for (const xmlPath of [...files.keys()].filter((p) => /^word\/header\d+\.xml$/.test(p)).sort()) {
    const xml = files.get(xmlPath)!;
    sectionSeq++;
    const regionParagraphs: string[] = [];
    const collected = collectRegionParagraphs(xml, xmlPath, "header", `s${sectionSeq}`, paragraphSeq);
    for (const p of collected) {
      paragraphSeq++;
      p.id = `p${paragraphSeq}`;
      regionParagraphs.push(p.id);
      paragraphs.push(p);
    }
    headers.push({
      id: `hdr${++headerSeq}`,
      xmlPath,
      kind: "header",
      sectionId: `s${sectionSeq}`,
      text: collected.map((p) => p.text).filter(Boolean).join("\n"),
      paragraphIds: regionParagraphs,
    });
  }

  for (const xmlPath of [...files.keys()].filter((p) => /^word\/footer\d+\.xml$/.test(p)).sort()) {
    const xml = files.get(xmlPath)!;
    sectionSeq++;
    const regionParagraphs: string[] = [];
    const collected = collectRegionParagraphs(xml, xmlPath, "footer", `s${sectionSeq}`, paragraphSeq);
    for (const p of collected) {
      paragraphSeq++;
      p.id = `p${paragraphSeq}`;
      regionParagraphs.push(p.id);
      paragraphs.push(p);
    }
    footers.push({
      id: `ftr${++footerSeq}`,
      xmlPath,
      kind: "footer",
      sectionId: `s${sectionSeq}`,
      text: collected.map((p) => p.text).filter(Boolean).join("\n"),
      paragraphIds: regionParagraphs,
    });
  }

  const footnotesXml = files.get("word/footnotes.xml");
  if (footnotesXml) {
    for (const range of findBalancedTags(footnotesXml, "w:footnote")) {
      const noteId = extractAttribute(range.openTag, "w:id");
      if (!noteId || noteId === "-1" || noteId === "0") continue;
      footnotes.push({
        id: `fn${noteId}`,
        noteId,
        text: extractAllText(range.xml).trim(),
        xmlPath: "word/footnotes.xml",
        xml: range.xml,
        start: range.start,
        end: range.end,
      });
    }
  }

  return { files, paragraphs, tables, footnotes, headers, footers };
}

export async function loadMutableWordState(filePath: string): Promise<{
  files: Map<string, string>;
  paragraphs: ParagraphInfo[];
  tables: TableInfo[];
  footnotes: FootnoteInfo[];
  headers: RegionInfo[];
  footers: RegionInfo[];
}> {
  const { files } = await loadDocx(filePath);
  return buildMutableWordStateFromFiles(files);
}

export function extractTableRows(tableXml: string): string[][] {
  const rows: string[][] = [];
  for (const row of findBalancedTags(tableXml, "w:tr")) {
    const cells: string[] = [];
    for (const cell of findBalancedTags(row.xml, "w:tc")) {
      cells.push(extractAllText(cell.xml).trim());
    }
    if (cells.length > 0) rows.push(cells);
  }
  return rows;
}

export function replaceParagraphText(paragraphXml: string, newText: string): string {
  const current = extractAllText(paragraphXml);
  if (current.length === 0) {
    const pPr = paragraphXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? "";
    return `<w:p>${pPr}${buildRunFromParagraph(paragraphXml, newText)}</w:p>`;
  }
  const { xml, count } = replaceTextInXml(paragraphXml, current, newText, false);
  if (count === 1) return xml;
  return rebuildParagraphText(paragraphXml, newText);
}

export function buildParagraphLike(referenceParagraphXml: string, text: string, style?: string): string {
  const pPr = referenceParagraphXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? "";
  let nextPPr = pPr;
  if (style) {
    if (/<w:pStyle\b/.test(nextPPr)) {
      nextPPr = nextPPr.replace(/<w:pStyle[^>]*w:val="[^"]*"[^>]*\/?>(?:<\/w:pStyle>)?/g, `<w:pStyle w:val="${style}"></w:pStyle>`);
    } else if (nextPPr) {
      nextPPr = nextPPr.replace("</w:pPr>", `<w:pStyle w:val="${style}"></w:pStyle></w:pPr>`);
    } else {
      nextPPr = `<w:pPr><w:pStyle w:val="${style}"></w:pStyle></w:pPr>`;
    }
  }
  return `<w:p>${nextPPr}${buildRunFromParagraph(referenceParagraphXml, text)}</w:p>`;
}

export function rebuildParagraphText(paragraphXml: string, newText: string): string {
  const pPr = paragraphXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? "";
  return `<w:p>${pPr}${buildRunFromParagraph(paragraphXml, newText)}</w:p>`;
}

export function buildRunFromParagraph(paragraphXml: string, text: string): string {
  const runStart = paragraphXml.match(/<w:r(?:\s|>)[\s\S]*?<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0];
  const rPr = paragraphXml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? "";
  const attrs = runStart?.match(/^<w:r([^>]*)>/)?.[1] ?? "";
  const xmlSpace = text.startsWith(" ") || text.endsWith(" ") || text.includes("  ") ? ' xml:space="preserve"' : "";
  return `<w:r${attrs}>${rPr}<w:t${xmlSpace}>${xmlEncode(text)}</w:t></w:r>`;
}

function markdownToText(markdown: string): string {
  return markdown
    .replace(/^---[\s\S]*?---\n?/m, "")
    .replace(/!\[[^\]]*\]\([^)]*\)/g, "")
    .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1")
    .replace(/^#{1,6}\s+/gm, "")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/\*([^*]+)\*/g, "$1")
    .trim();
}

function collectRegionParagraphs(
  xml: string,
  xmlPath: string,
  kind: "header" | "footer",
  sectionId: string,
  paragraphOffset: number
): ParagraphInfo[] {
  const root = getBalancedTagRange(xml, kind === "header" ? "w:hdr" : "w:ftr");
  if (!root) return [];
  const out: ParagraphInfo[] = [];
  let localIndex = 0;
  for (const block of findTopLevelBlocks(root.innerXml, ["w:p"])) {
    out.push(
      makeParagraphInfo(
        `p${paragraphOffset + ++localIndex}`,
        xmlPath,
        kind,
        sectionId,
        block.xml,
        root.innerStart + block.start,
        root.innerStart + block.end
      )
    );
  }
  return out;
}

function makeParagraphInfo(
  id: string,
  xmlPath: string,
  kind: RegionKind,
  sectionId: string,
  xml: string,
  start: number,
  end: number
): ParagraphInfo {
  const style = xml.match(/<w:pStyle[^>]*w:val="([^"]+)"/)?.[1] ?? null;
  const headingMatch = style?.match(/^Heading(\d)$/i) ?? style?.match(/^heading\s*(\d)$/i);
  const outline = xml.match(/<w:outlineLvl[^>]*w:val="(\d+)"/)?.[1];
  const headingLevel = headingMatch ? parseInt(headingMatch[1], 10) : outline !== undefined ? parseInt(outline, 10) + 1 : null;
  return {
    id,
    xmlPath,
    kind,
    sectionId,
    text: extractAllText(xml),
    style,
    headingLevel,
    xml,
    start,
    end,
  };
}

export function getBalancedTagRange(xml: string, tagName: string): BalancedTagRange | null {
  const all = findBalancedTags(xml, tagName);
  return all[0] ?? null;
}

export function findBalancedTags(xml: string, tagName: string): BalancedTagRange[] {
  const results: BalancedTagRange[] = [];
  const openRe = new RegExp(`<${escapeRegex(tagName)}(?=[\\s>/])[^>]*?>`, "g");
  const closeRe = new RegExp(`</${escapeRegex(tagName)}>`, "g");
  const events: Array<{ pos: number; type: "open" | "close"; tag: string }> = [];

  let m: RegExpExecArray | null;
  while ((m = openRe.exec(xml)) !== null) {
    const openTag = m[0];
    if (openTag.endsWith("/>") || openTag.startsWith("<?") || openTag.startsWith("<!")) {
      results.push({
        start: m.index,
        end: m.index + openTag.length,
        innerStart: m.index + openTag.length,
        innerEnd: m.index + openTag.length,
        openTag,
        closeTag: "",
        xml: openTag,
        innerXml: "",
      });
    } else {
      events.push({ pos: m.index, type: "open", tag: openTag });
    }
  }
  while ((m = closeRe.exec(xml)) !== null) {
    events.push({ pos: m.index, type: "close", tag: m[0] });
  }

  events.sort((a, b) => a.pos - b.pos || (a.type === "close" ? 1 : -1));
  const stack: Array<{ start: number; openTag: string }> = [];
  for (const event of events) {
    if (event.type === "open") {
      stack.push({ start: event.pos, openTag: event.tag });
    } else if (stack.length > 0) {
      const open = stack.pop()!;
      const closeTag = event.tag;
      const end = event.pos + closeTag.length;
      const innerStart = open.start + open.openTag.length;
      const innerEnd = event.pos;
      results.push({
        start: open.start,
        end,
        innerStart,
        innerEnd,
        openTag: open.openTag,
        closeTag,
        xml: xml.slice(open.start, end),
        innerXml: xml.slice(innerStart, innerEnd),
      });
    }
  }

  return results.sort((a, b) => a.start - b.start);
}

export function findTopLevelBlocks(xml: string, tags: string[]): BlockRange[] {
  const blocks: BlockRange[] = [];
  const stack: Array<{ tag: string; start: number; openTag: string }> = [];
  const tagRegex = /<\/?([A-Za-z0-9:_-]+)(?=[\s>/])[^>]*?>/g;
  let m: RegExpExecArray | null;
  while ((m = tagRegex.exec(xml)) !== null) {
    const full = m[0];
    const tag = m[1];
    if (!tags.includes(tag) && stack.length === 0) continue;
    const isClose = full.startsWith("</");
    const isSelfClosing = full.endsWith("/>");
    if (!isClose) {
      if (tags.includes(tag) && stack.length === 0) {
        if (isSelfClosing) {
          blocks.push({ tag, start: m.index, end: m.index + full.length, xml: full });
        } else {
          stack.push({ tag, start: m.index, openTag: full });
        }
      } else if (stack.length > 0 && !isSelfClosing) {
        stack.push({ tag, start: m.index, openTag: full });
      }
    } else if (stack.length > 0) {
      const open = stack.pop()!;
      if (stack.length === 0 && tags.includes(open.tag) && open.tag === tag) {
        blocks.push({ tag, start: open.start, end: m.index + full.length, xml: xml.slice(open.start, m.index + full.length) });
      }
    }
  }
  return blocks.sort((a, b) => a.start - b.start);
}

export function extractAttribute(tagXml: string, name: string): string | null {
  return tagXml.match(new RegExp(`${escapeRegex(name)}="([^"]*)"`))?.[1] ?? null;
}

function escapeRegex(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
