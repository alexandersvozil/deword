import { XMLParser } from "fast-xml-parser";
import type { UnpackedDocx, Relationship } from "../unpack.js";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: false,
  trimValues: false,
});

interface ExtractedContent {
  markdown: string;
  images: ImageRef[];
  tables: TableData[];
  metadata: DocMetadata;
}

interface ImageRef {
  id: string;
  path: string;
  description: string;
  /** Size in bytes */
  size: number;
}

interface TableData {
  index: number;
  rows: string[][];
}

interface DocMetadata {
  title: string | null;
  author: string | null;
  created: string | null;
  modified: string | null;
  paragraphCount: number;
  imageCount: number;
  tableCount: number;
}

export function extractContent(doc: UnpackedDocx): ExtractedContent {
  const images: ImageRef[] = [];
  const tables: TableData[] = [];

  // Parse core.xml for metadata
  const metadata = extractMetadata(doc);

  // Parse document body
  const parsed = parser.parse(doc.documentXml);
  const body = parsed?.["w:document"]?.["w:body"];
  if (!body) {
    return { markdown: "*Empty document*", images, tables, metadata };
  }

  const lines: string[] = [];
  let tableIndex = 0;

  // Body can contain: w:p (paragraphs), w:tbl (tables), w:sdt (structured doc tags)
  const elements = normalizeToArray(body);

  for (const [key, value] of iterateBodyElements(body)) {
    if (key === "w:p") {
      const text = extractParagraph(value, doc.relationships, images, doc.media);
      if (text !== null) {
        lines.push(text);
      }
    } else if (key === "w:tbl") {
      const table = extractTable(value, tableIndex, doc.relationships, images, doc.media);
      tables.push(table);
      lines.push(formatTableAsMarkdown(table));
      tableIndex++;
    } else if (key === "w:sdt") {
      // Structured document tag - extract inner content
      const sdtContent = value?.["w:sdtContent"];
      if (sdtContent) {
        for (const [innerKey, innerVal] of iterateBodyElements(sdtContent)) {
          if (innerKey === "w:p") {
            const text = extractParagraph(innerVal, doc.relationships, images, doc.media);
            if (text !== null) lines.push(text);
          }
        }
      }
    }
  }

  // Collect image refs from media map
  for (const [path, buf] of doc.media) {
    const existing = images.find((i) => i.path === path);
    if (!existing) {
      images.push({
        id: path,
        path,
        description: "",
        size: buf.length,
      });
    }
  }

  metadata.paragraphCount = lines.length;
  metadata.imageCount = images.length;
  metadata.tableCount = tables.length;

  return {
    markdown: lines.join("\n\n"),
    images,
    tables,
    metadata,
  };
}

/** Iterate over body-level elements preserving document order */
function* iterateBodyElements(
  body: any
): Generator<[string, any]> {
  // fast-xml-parser without preserveOrder groups by tag name,
  // so we lose document order for mixed p/tbl sequences.
  // We handle this by checking if body is object with arrays.
  const keys = ["w:p", "w:tbl", "w:sdt", "w:sectPr"];

  // If the body has arrays, we need to interleave them.
  // Without preserveOrder, we yield paragraphs first, then tables.
  // This is a known limitation - for perfect ordering we'd need
  // preserveOrder:true, but that makes the rest of parsing much harder.
  // For agent consumption, grouping is acceptable.
  for (const key of keys) {
    if (key === "w:sectPr") continue; // skip section properties
    const items = body[key];
    if (!items) continue;
    const arr = Array.isArray(items) ? items : [items];
    for (const item of arr) {
      yield [key, item];
    }
  }
}

function extractParagraph(
  p: any,
  rels: Map<string, Relationship>,
  images: ImageRef[],
  media: Map<string, Buffer>
): string | null {
  if (!p) return null;

  // Check paragraph style for heading level
  const pStyle = p?.["w:pPr"]?.["w:pStyle"]?.["@_w:val"] ?? "";
  const numPr = p?.["w:pPr"]?.["w:numPr"];
  const outlineLvl = p?.["w:pPr"]?.["w:outlineLvl"]?.["@_w:val"];

  let prefix = "";

  // Detect headings
  if (pStyle.match(/^Heading(\d)$/i) || pStyle.match(/^heading\s*(\d)$/i)) {
    const level = parseInt(pStyle.replace(/\D/g, ""), 10);
    prefix = "#".repeat(Math.min(level, 6)) + " ";
  } else if (outlineLvl !== undefined) {
    const level = parseInt(outlineLvl, 10) + 1;
    prefix = "#".repeat(Math.min(level, 6)) + " ";
  }

  // Detect list items
  if (numPr) {
    const ilvl = parseInt(numPr?.["w:ilvl"]?.["@_w:val"] ?? "0", 10);
    const indent = "  ".repeat(ilvl);
    prefix = indent + "- ";
  }

  // Extract text from runs
  const textParts: string[] = [];
  const runs = normalizeToArray(p?.["w:r"]);

  for (const run of runs) {
    if (!run) continue;

    // Check for images (drawings / inline shapes)
    const drawing = run?.["w:drawing"];
    if (drawing) {
      const imgRef = extractImageFromDrawing(drawing, rels, media);
      if (imgRef) {
        images.push(imgRef);
        textParts.push(`![${imgRef.description || "image"}](${imgRef.path})`);
        continue;
      }
    }

    // Check for embedded objects
    const obj = run?.["w:object"];
    if (obj) {
      textParts.push("[embedded object]");
      continue;
    }

    // Regular text
    const text = run?.["w:t"];
    if (text !== undefined) {
      const textStr = typeof text === "object" ? text["#text"] ?? "" : String(text);
      
      // Check for formatting — only apply markers to non-whitespace text
      // (Word splits runs at formatting boundaries, producing whitespace-only
      // bold/italic runs that become garbled markdown like "*** ***")
      const rPr = run?.["w:rPr"];
      let formatted = textStr;
      const hasContent = textStr.trim().length > 0;
      if (hasContent && rPr?.["w:b"] !== undefined) formatted = `**${formatted}**`;
      if (hasContent && rPr?.["w:i"] !== undefined) formatted = `*${formatted}*`;
      if (hasContent && rPr?.["w:strike"] !== undefined) formatted = `~~${formatted}~~`;
      if (rPr?.["w:u"] !== undefined) {
        // Underline doesn't have markdown, note it
      }

      textParts.push(formatted);
    }

    // Tab
    if (run?.["w:tab"] !== undefined) {
      textParts.push("\t");
    }

    // Line break
    if (run?.["w:br"] !== undefined) {
      textParts.push("\n");
    }
  }

  // Also check for hyperlinks
  const hyperlinks = normalizeToArray(p?.["w:hyperlink"]);
  for (const hl of hyperlinks) {
    if (!hl) continue;
    const rId = hl?.["@_r:id"];
    const innerRuns = normalizeToArray(hl?.["w:r"]);
    const linkText = innerRuns
      .map((r: any) => {
        const t = r?.["w:t"];
        return typeof t === "object" ? t["#text"] ?? "" : String(t ?? "");
      })
      .join("");

    if (rId && rels.has(rId)) {
      const rel = rels.get(rId)!;
      textParts.push(`[${linkText}](${rel.target})`);
    } else {
      textParts.push(linkText);
    }
  }

  const mathText = extractMathText(p?.["m:oMath"]) || extractMathText(p?.["m:oMathPara"]);
  if (mathText) {
    textParts.push(mathText);
  }

  const fullText = prefix + textParts.join("");
  // Return null for truly empty paragraphs (not even whitespace)
  if (fullText.trim() === "") return "";
  return fullText;
}

function extractMathText(node: any): string {
  if (!node) return "";
  if (typeof node === "string") return node;
  if (Array.isArray(node)) return node.map(extractMathText).join("");

  return Object.keys(node)
    .filter((key) => key.startsWith("m:"))
    .map((key) => renderMathKey(key, node[key]))
    .join("");
}

function renderMathKey(key: string, value: any): string {
  if (Array.isArray(value)) {
    return value.map((item) => renderMathKey(key, item)).join("");
  }
  if (key === "m:t") {
    return typeof value === "object" ? String(value["#text"] ?? "") : String(value ?? "");
  }
  if (key === "m:f") {
    return `(${extractMathText(value["m:num"])})/(${extractMathText(value["m:den"])})`;
  }
  if (key === "m:sSup") {
    return `${extractMathText(value["m:e"])}^(${extractMathText(value["m:sup"])})`;
  }
  if (key === "m:sSub") {
    return `${extractMathText(value["m:e"])}_(${extractMathText(value["m:sub"])})`;
  }
  if (key === "m:sSubSup") {
    return `${extractMathText(value["m:e"])}_(${extractMathText(value["m:sub"])})^(${extractMathText(value["m:sup"])})`;
  }
  if (key === "m:rad") {
    const degree = extractMathText(value["m:deg"]);
    const expr = extractMathText(value["m:e"]);
    return degree ? `root(${degree})(${expr})` : `sqrt(${expr})`;
  }
  if (key === "m:d") {
    return `(${extractMathText(value["m:e"] ?? value)})`;
  }
  return extractMathText(value);
}

function extractTable(
  tbl: any,
  index: number,
  rels: Map<string, Relationship>,
  images: ImageRef[],
  media: Map<string, Buffer>
): TableData {
  const rows: string[][] = [];
  const tblRows = normalizeToArray(tbl?.["w:tr"]);

  for (const row of tblRows) {
    if (!row) continue;
    const cells: string[] = [];
    const tblCells = normalizeToArray(row?.["w:tc"]);

    for (const cell of tblCells) {
      if (!cell) continue;
      const paragraphs = normalizeToArray(cell?.["w:p"]);
      const cellText = paragraphs
        .map((p: any) => extractParagraph(p, rels, images, media))
        .filter((t: string | null) => t !== null)
        .join(" | ");
      cells.push(cellText.trim());
    }
    rows.push(cells);
  }

  return { index, rows };
}

function formatTableAsMarkdown(table: TableData): string {
  if (table.rows.length === 0) return "";

  // Determine column widths
  const colCount = Math.max(...table.rows.map((r) => r.length));
  const colWidths = new Array(colCount).fill(3);
  for (const row of table.rows) {
    for (let i = 0; i < row.length; i++) {
      colWidths[i] = Math.max(colWidths[i], row[i].length);
    }
  }

  const lines: string[] = [];
  for (let rowIdx = 0; rowIdx < table.rows.length; rowIdx++) {
    const row = table.rows[rowIdx];
    const cells = [];
    for (let i = 0; i < colCount; i++) {
      cells.push((row[i] ?? "").padEnd(colWidths[i]));
    }
    lines.push("| " + cells.join(" | ") + " |");

    // Add header separator after first row
    if (rowIdx === 0) {
      const sep = colWidths.map((w) => "-".repeat(w));
      lines.push("| " + sep.join(" | ") + " |");
    }
  }

  return lines.join("\n");
}

function extractImageFromDrawing(
  drawing: any,
  rels: Map<string, Relationship>,
  media: Map<string, Buffer>
): ImageRef | null {
  // Navigate the drawing hierarchy to find the image reference
  // wp:inline or wp:anchor > a:graphic > a:graphicData > pic:pic > pic:blipFill > a:blip
  const inline = drawing?.["wp:inline"] ?? drawing?.["wp:anchor"];
  if (!inline) return null;

  const desc =
    inline?.["wp:docPr"]?.["@_descr"] ??
    inline?.["wp:docPr"]?.["@_name"] ??
    "";

  const graphic = inline?.["a:graphic"];
  const graphicData = graphic?.["a:graphicData"];
  const pic = graphicData?.["pic:pic"];
  const blipFill = pic?.["pic:blipFill"];
  const blip = blipFill?.["a:blip"];

  if (!blip) return null;

  const rId = blip?.["@_r:embed"];
  if (!rId || !rels.has(rId)) return null;

  const rel = rels.get(rId)!;
  const imgPath = "word/" + rel.target.replace(/^\//, "");
  const size = media.get(imgPath)?.length ?? 0;

  return {
    id: rId,
    path: imgPath,
    description: desc,
    size,
  };
}

/**
 * fast-xml-parser may return a plain string, an object with #text, or undefined.
 * Normalise to string | null.
 */
function extractTextValue(val: any): string | null {
  if (val == null) return null;
  if (typeof val === "string") return val || null;
  if (typeof val === "object" && "#text" in val) return String(val["#text"]) || null;
  return String(val) || null;
}

function extractMetadata(doc: UnpackedDocx): DocMetadata {
  const meta: DocMetadata = {
    title: null,
    author: null,
    created: null,
    modified: null,
    paragraphCount: 0,
    imageCount: 0,
    tableCount: 0,
  };

  const coreXml = doc.files.get("docProps/core.xml");
  if (!coreXml) return meta;

  try {
    const parsed = parser.parse(coreXml);
    const props = parsed?.["cp:coreProperties"];
    if (props) {
      meta.title = extractTextValue(props?.["dc:title"]);
      meta.author = extractTextValue(props?.["dc:creator"]);
      meta.created = extractTextValue(props?.["dcterms:created"]);
      meta.modified = extractTextValue(props?.["dcterms:modified"]);
    }
  } catch {
    // Metadata extraction is best-effort
  }

  return meta;
}

function normalizeToArray(val: any): any[] {
  if (!val) return [];
  return Array.isArray(val) ? val : [val];
}

export type { ExtractedContent, ImageRef, TableData, DocMetadata };
