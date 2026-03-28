import { readFile } from "fs/promises";
import { basename, extname } from "path";
import { detectFileType } from "../utils/detect.js";
import { buildMutableWordStateFromFiles, buildParagraphLike, extractAttribute, extractTableRows, findBalancedTags, getBalancedTagRange, loadMutableWordState, rebuildParagraphText, replaceParagraphText, type FootnoteInfo, type ParagraphInfo, type RegionInfo, type TableInfo } from "../utils/wordModel.js";
import { extractAllText, mergeAdjacentRuns, replaceTextInXml, xmlEncode } from "../utils/xml.js";
import { repackDocxEntries } from "../utils/zipEdit.js";
import { checkSdtField, fieldDisplayName, fillSdtField, findSdtFields, matchField } from "../utils/sdt.js";

interface PatchRequest {
  version?: string;
  document_path?: string;
  output_path?: string | null;
  options?: {
    dry_run?: boolean;
    validate?: boolean;
    track_changes?: boolean;
    fail_fast?: boolean;
    preserve_formatting?: boolean;
  };
  operations: PatchOperation[];
}

type MatchMode = "exact" | "contains" | "regex";

interface TextSelector {
  text: string;
  match?: MatchMode;
  occurrence?: number;
  within?: {
    paragraph_id?: string | null;
    heading_text?: string | null;
  };
}

interface Selector {
  by_id?: string;
  by_text?: TextSelector;
  by_heading?: { text: string; level?: number; occurrence?: number };
  by_table?: { table_id: string };
  by_region?: { kind: "header" | "footer"; section?: "first" | "last" | "all" | number | string; slot?: string };
}

type PatchOperation =
  | { op: "replace_text"; target: Selector; new_text: string }
  | { op: "replace_all"; find: { text: string; match?: MatchMode }; new_text: string; scope?: "body" | "headers" | "footers" | "footnotes" | "all" }
  | { op: "edit_paragraph"; target: Selector; new_text: string }
  | { op: "insert_paragraph"; location: { before?: Selector; after?: Selector }; text: string; style?: string }
  | { op: "insert_footnote"; anchor: Selector; footnote_text: string }
  | { op: "edit_footnote"; target: Selector; new_text: string }
  | { op: "insert_table"; location: { before?: Selector; after?: Selector }; data: string[][]; table_style?: string; header_row?: boolean; autofit?: boolean; caption?: string | null; clone_style_from_table?: string; clone_caption_from_paragraph?: string }
  | { op: "edit_table_cell"; target: { table_id: string; row: number; col: number }; new_text: string }
  | { op: "append_table_row"; target: { table_id: string }; row: string[] }
  | { op: "insert_image"; location: { before?: Selector; after?: Selector }; file_path: string; alt_text?: string | null; width_px?: number | null; height_px?: number | null; preserve_aspect_ratio?: boolean; alignment?: "left" | "center" | "right"; caption?: string | null; clone_layout_from_paragraph?: string; clone_caption_from_paragraph?: string }
  | { op: "set_region_text"; target: Selector; text: string }
  | { op: "fill_field"; field: { name: string }; value: string }
  | { op: "set_checkbox"; field: { name: string }; checked: boolean };

interface PatchOptions {
  input?: string;
  patchFile?: string;
  output?: string;
}

interface MutableState {
  files: Map<string, string>;
  paragraphs: ParagraphInfo[];
  tables: TableInfo[];
  footnotes: FootnoteInfo[];
  headers: RegionInfo[];
  footers: RegionInfo[];
}

interface PatchResultItem {
  index: number;
  op: string;
  status: "applied" | "failed";
  details?: Record<string, unknown>;
  error?: string;
}

export async function runPatch(filePath: string, options: PatchOptions): Promise<void> {
  const { type } = await detectFileType(filePath);
  if (type !== "docx") {
    throw new Error(`Only .docx files are supported. Detected format: ${type}`);
  }

  const request = await readPatchRequest(filePath, options);
  const state = await loadMutableWordState(filePath);
  const binaryEntries = new Map<string, Buffer>();
  const results: PatchResultItem[] = [];
  const failFast = request.options?.fail_fast !== false;

  for (let i = 0; i < request.operations.length; i++) {
    const op = request.operations[i];
    try {
      const details = await applyOperation(state, op, binaryEntries);
      results.push({ index: i, op: op.op, status: "applied", details });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      results.push({ index: i, op: op.op, status: "failed", error: message });
      if (failFast) {
        throw new Error(formatPatchFailure(message, results));
      }
    }
  }

  if (request.options?.dry_run) {
    process.stdout.write(
      JSON.stringify(
        {
          ok: results.every((r) => r.status === "applied"),
          dry_run: true,
          results,
        },
        null,
        2
      ) + "\n"
    );
    return;
  }

  const textEntries = new Map<string, string>();
  for (const [path, xml] of state.files) {
    textEntries.set(path, xml);
  }
  await repackDocxEntries(filePath, textEntries, binaryEntries, request.output_path ?? options.output ?? undefined);

  process.stdout.write(
    JSON.stringify(
      {
        ok: results.every((r) => r.status === "applied"),
        document_path: filePath,
        output_path: request.output_path ?? options.output ?? filePath,
        summary: {
          requested_ops: request.operations.length,
          applied_ops: results.filter((r) => r.status === "applied").length,
          failed_ops: results.filter((r) => r.status === "failed").length,
        },
        results,
      },
      null,
      2
    ) + "\n"
  );
}

async function readPatchRequest(filePath: string, options: PatchOptions): Promise<PatchRequest> {
  let raw: string;
  if (options.patchFile) {
    raw = await readFile(options.patchFile, "utf-8");
  } else if (options.input) {
    raw = await readFile(options.input, "utf-8");
  } else {
    const chunks: Buffer[] = [];
    for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
    raw = Buffer.concat(chunks).toString("utf-8");
  }
  const request = JSON.parse(raw) as PatchRequest;
  if (!Array.isArray(request.operations) || request.operations.length === 0) {
    throw new Error("Patch JSON must contain a non-empty operations array.");
  }
  request.document_path = request.document_path ?? filePath;
  request.options = request.options ?? {};
  return request;
}

async function applyOperation(state: MutableState, op: PatchOperation, binaryEntries: Map<string, Buffer>): Promise<Record<string, unknown>> {
  switch (op.op) {
    case "replace_text": {
      const target = resolveReplaceTarget(state, op.target);
      const currentXml = state.files.get(target.xmlPath)!;
      const sourceText = target.oldText;
      const { xml, count } = replaceTextInXml(currentXml, sourceText, op.new_text, false);
      if (count !== 1) throw new Error(`Expected exactly one match for replace_text in ${target.xmlPath}, got ${count}`);
      state.files.set(target.xmlPath, xml);
      rebuildState(state);
      return { matches: 1, xml_path: target.xmlPath };
    }
    case "replace_all": {
      const paths = scopePaths(state, op.scope ?? "all");
      let total = 0;
      for (const path of paths) {
        const currentXml = state.files.get(path)!;
        const { xml, count } = replaceTextInXml(currentXml, op.find.text, op.new_text, true);
        if (count > 0) {
          state.files.set(path, xml);
          total += count;
        }
      }
      if (total === 0) throw new Error(`No matches found for replace_all: ${op.find.text}`);
      rebuildState(state);
      return { matches: total, scope: op.scope ?? "all" };
    }
    case "edit_paragraph": {
      const paragraph = resolveParagraph(state, op.target);
      const xml = state.files.get(paragraph.xmlPath)!;
      const updatedParagraph = replaceParagraphText(paragraph.xml, op.new_text);
      state.files.set(paragraph.xmlPath, replaceRange(xml, paragraph.start, paragraph.end, updatedParagraph));
      rebuildState(state);
      return { paragraph_id: paragraph.id, xml_path: paragraph.xmlPath };
    }
    case "insert_paragraph": {
      const { anchor, mode } = resolveLocationParagraph(state, op.location);
      if (anchor.xmlPath !== "word/document.xml") throw new Error("insert_paragraph currently supports body locations only.");
      const documentXml = state.files.get("word/document.xml")!;
      const newParagraph = buildParagraphLike(anchor.xml, op.text, op.style);
      const insertAt = mode === "after" ? anchor.end : anchor.start;
      state.files.set("word/document.xml", documentXml.slice(0, insertAt) + newParagraph + documentXml.slice(insertAt));
      rebuildState(state);
      return { inserted_after: mode === "after" ? anchor.id : undefined, inserted_before: mode === "before" ? anchor.id : undefined };
    }
    case "insert_footnote": {
      const details = insertFootnote(state, op.anchor, op.footnote_text);
      rebuildState(state);
      return details;
    }
    case "edit_footnote": {
      const footnote = resolveFootnote(state, op.target);
      const xml = state.files.get(footnote.xmlPath)!;
      const replacement = replaceTextInXml(footnote.xml, footnote.text, op.new_text, false);
      const updatedNoteXml = replacement.count === 1 ? replacement.xml : rebuildFootnote(footnote.xml, op.new_text);
      state.files.set(footnote.xmlPath, replaceRange(xml, footnote.start, footnote.end, updatedNoteXml));
      rebuildState(state);
      return { footnote_id: footnote.id };
    }
    case "insert_table": {
      const { anchor, mode } = resolveLocationParagraph(state, op.location);
      if (anchor.xmlPath !== "word/document.xml") throw new Error("insert_table currently supports body locations only.");
      const tableXml = op.clone_style_from_table
        ? buildClonedTableXml(resolveTable(state, op.clone_style_from_table), op.data)
        : buildTableXml(op.data, op.table_style ?? "professional", op.header_row !== false);
      const captionSource = op.clone_caption_from_paragraph ? resolveParagraphById(state, op.clone_caption_from_paragraph) : undefined;
      const captionXml = op.caption ? buildCaptionParagraphXml(op.caption, captionSource) : "";
      const rawBlock = mode === "after" ? tableXml + captionXml : captionXml + tableXml;
      const xml = state.files.get("word/document.xml")!;
      const insertAt = mode === "after" ? anchor.end : anchor.start;
      const block = ensureTableBlockSeparated(rawBlock, xml, insertAt, Boolean(op.caption));
      state.files.set("word/document.xml", xml.slice(0, insertAt) + block + xml.slice(insertAt));
      rebuildState(state);
      return { rows: op.data.length, cols: Math.max(...op.data.map((r) => r.length)), cloned_from_table: op.clone_style_from_table ?? null };
    }
    case "edit_table_cell": {
      const table = resolveTable(state, op.target.table_id);
      const rows = findBalancedTags(table.xml, "w:tr");
      if (op.target.row < 0 || op.target.row >= rows.length) throw new Error(`Row ${op.target.row} out of range for ${table.id}`);
      const cells = findBalancedTags(rows[op.target.row].xml, "w:tc");
      if (op.target.col < 0 || op.target.col >= cells.length) throw new Error(`Column ${op.target.col} out of range for ${table.id}`);
      const cell = cells[op.target.col];
      const currentCellText = extractAllText(cell.xml);
      const updatedCell = currentCellText
        ? replaceTextInXml(cell.xml, currentCellText, op.new_text, false).xml
        : rebuildCell(cell.xml, op.new_text);
      const updatedRow = replaceRange(rows[op.target.row].xml, cell.start, cell.end, updatedCell);
      const updatedTable = replaceRange(table.xml, rows[op.target.row].start, rows[op.target.row].end, updatedRow);
      const documentXml = state.files.get(table.xmlPath)!;
      state.files.set(table.xmlPath, replaceRange(documentXml, table.start, table.end, updatedTable));
      rebuildState(state);
      return { table_id: table.id, row: op.target.row, col: op.target.col };
    }
    case "append_table_row": {
      const table = resolveTable(state, op.target.table_id);
      const newRow = buildTableRowXml(table.xml, op.row);
      const updatedTable = table.xml.replace(/<\/w:tbl>\s*$/, `${newRow}</w:tbl>`);
      const documentXml = state.files.get(table.xmlPath)!;
      state.files.set(table.xmlPath, replaceRange(documentXml, table.start, table.end, updatedTable));
      rebuildState(state);
      return { table_id: table.id, appended_cells: op.row.length };
    }
    case "set_region_text": {
      const regions = resolveRegions(state, op.target);
      for (const region of regions) {
        const paragraphs = state.paragraphs.filter((p) => region.paragraphIds.includes(p.id) && p.text.trim().length > 0);
        if (paragraphs.length !== 1) {
          throw new Error(`set_region_text expects exactly one non-empty paragraph in ${region.id}, found ${paragraphs.length}`);
        }
        const p = paragraphs[0];
        const regionXml = state.files.get(p.xmlPath)!;
        state.files.set(p.xmlPath, replaceRange(regionXml, p.start, p.end, replaceParagraphText(p.xml, op.text)));
        rebuildState(state);
      }
      return { regions: regions.map((r) => r.id) };
    }
    case "fill_field": {
      const details = fillFieldInState(state, op.field.name, op.value);
      rebuildState(state);
      return details;
    }
    case "set_checkbox": {
      const details = setCheckboxInState(state, op.field.name, op.checked);
      rebuildState(state);
      return details;
    }
    case "insert_image": {
      const details = await insertImage(state, op, binaryEntries);
      rebuildState(state);
      return details;
    }
  }
}

function formatPatchFailure(message: string, results: PatchResultItem[]): string {
  return `${message}\n\nPatch results so far:\n${results
    .map((r) => `  [${r.index}] ${r.op}: ${r.status}${r.error ? ` — ${r.error}` : ""}`)
    .join("\n")}`;
}

function rebuildState(state: MutableState): void {
  const next = buildMutableWordStateFromFiles(state.files);
  state.paragraphs = next.paragraphs;
  state.tables = next.tables;
  state.footnotes = next.footnotes;
  state.headers = next.headers;
  state.footers = next.footers;
}

function scopePaths(state: MutableState, scope: string): string[] {
  const all = [...state.files.keys()];
  switch (scope) {
    case "body":
      return ["word/document.xml"];
    case "headers":
      return all.filter((p) => /^word\/header\d+\.xml$/.test(p));
    case "footers":
      return all.filter((p) => /^word\/footer\d+\.xml$/.test(p));
    case "footnotes":
      return all.filter((p) => p === "word/footnotes.xml");
    default:
      return ["word/document.xml", ...all.filter((p) => /^word\/(header\d+|footer\d+)\.xml$/.test(p)), ...all.filter((p) => p === "word/footnotes.xml")];
  }
}

function resolveReplaceTarget(state: MutableState, selector: Selector): { xmlPath: string; oldText: string } {
  if (selector.by_text) {
    const target = selector.by_text;
    const occurrence = target.occurrence ?? 1;
    const paths = scopedPathsForTextSelector(state, target);
    let seen = 0;
    for (const path of paths) {
      const xml = state.files.get(path)!;
      const count = countMatches(xml, target.text, target.match ?? "exact");
      if (seen + count >= occurrence) {
        return { xmlPath: path, oldText: target.text };
      }
      seen += count;
    }
    throw new Error(`replace_text target not found: ${target.text}`);
  }
  throw new Error("replace_text currently requires a by_text selector.");
}

function countMatches(xml: string, text: string, mode: MatchMode): number {
  const haystack = extractAllText(xml);
  if (mode === "exact" || mode === "contains") {
    let count = 0;
    let idx = 0;
    while ((idx = haystack.indexOf(text, idx)) !== -1) {
      count++;
      idx += Math.max(text.length, 1);
    }
    return count;
  }
  const re = new RegExp(text, "g");
  return [...haystack.matchAll(re)].length;
}

function scopedPathsForTextSelector(state: MutableState, selector: TextSelector): string[] {
  if (selector.within?.paragraph_id) {
    const p = state.paragraphs.find((paragraph) => paragraph.id === selector.within!.paragraph_id);
    if (!p) throw new Error(`Paragraph not found: ${selector.within!.paragraph_id}`);
    return [p.xmlPath];
  }
  return scopePaths(state, "all");
}

function resolveParagraphById(state: MutableState, paragraphId: string): ParagraphInfo {
  const p = state.paragraphs.find((paragraph) => paragraph.id === paragraphId);
  if (!p) throw new Error(`Paragraph not found: ${paragraphId}`);
  return p;
}

function resolveParagraph(state: MutableState, selector: Selector): ParagraphInfo {
  if (selector.by_id) {
    return resolveParagraphById(state, selector.by_id);
  }
  if (selector.by_heading) {
    const occurrence = selector.by_heading.occurrence ?? 1;
    const candidates = state.paragraphs.filter(
      (p) => p.kind === "body" && p.headingLevel !== null && p.text === selector.by_heading!.text && (selector.by_heading!.level == null || p.headingLevel === selector.by_heading!.level)
    );
    if (candidates.length < occurrence) throw new Error(`Heading not found: ${selector.by_heading.text}`);
    return candidates[occurrence - 1];
  }
  if (selector.by_text) {
    const candidates = filterParagraphsByWithin(state, selector.by_text.within).filter((p) => matchParagraphText(p, selector.by_text!));
    const occurrence = selector.by_text.occurrence ?? 1;
    if (candidates.length < occurrence) throw new Error(`Paragraph text selector not found: ${selector.by_text.text}`);
    return candidates[occurrence - 1];
  }
  throw new Error("Unsupported paragraph selector.");
}

function filterParagraphsByWithin(state: MutableState, within?: { paragraph_id?: string | null; heading_text?: string | null }): ParagraphInfo[] {
  if (!within) return state.paragraphs;
  if (within.paragraph_id) {
    const p = state.paragraphs.find((paragraph) => paragraph.id === within.paragraph_id);
    return p ? [p] : [];
  }
  if (within.heading_text) {
    const headings = state.paragraphs.filter((p) => p.kind === "body" && p.headingLevel !== null);
    const heading = headings.find((p) => p.text === within.heading_text);
    if (!heading) return [];
    const bodyParagraphs = state.paragraphs.filter((p) => p.kind === "body");
    const startIndex = bodyParagraphs.findIndex((p) => p.id === heading.id);
    if (startIndex === -1) return [];
    const out: ParagraphInfo[] = [];
    for (let i = startIndex + 1; i < bodyParagraphs.length; i++) {
      const p = bodyParagraphs[i];
      if (p.headingLevel !== null && (heading.headingLevel == null || p.headingLevel <= heading.headingLevel)) break;
      out.push(p);
    }
    return out;
  }
  return state.paragraphs;
}

function matchParagraphText(paragraph: ParagraphInfo, selector: TextSelector): boolean {
  const mode = selector.match ?? "exact";
  if (mode === "exact") return paragraph.text === selector.text || paragraph.text.includes(selector.text);
  if (mode === "contains") return paragraph.text.includes(selector.text);
  return new RegExp(selector.text).test(paragraph.text);
}

function resolveLocationParagraph(state: MutableState, location: { before?: Selector; after?: Selector }): { anchor: ParagraphInfo; mode: "before" | "after" } {
  if (location.after) return { anchor: resolveParagraph(state, location.after), mode: "after" };
  if (location.before) return { anchor: resolveParagraph(state, location.before), mode: "before" };
  throw new Error("Location must include before or after.");
}

function resolveFootnote(state: MutableState, selector: Selector): FootnoteInfo {
  if (selector.by_id) {
    const footnote = state.footnotes.find((f) => f.id === selector.by_id || f.noteId === selector.by_id);
    if (!footnote) throw new Error(`Footnote not found: ${selector.by_id}`);
    return footnote;
  }
  if (selector.by_text) {
    const occurrence = selector.by_text.occurrence ?? 1;
    const matches = state.footnotes.filter((f) => f.text.includes(selector.by_text!.text));
    if (matches.length < occurrence) throw new Error(`Footnote text selector not found: ${selector.by_text.text}`);
    return matches[occurrence - 1];
  }
  throw new Error("Unsupported footnote selector.");
}

function resolveTable(state: MutableState, tableId: string): TableInfo {
  const table = state.tables.find((t) => t.id === tableId);
  if (!table) throw new Error(`Table not found: ${tableId}`);
  return table;
}

function resolveRegions(state: MutableState, selector: Selector): RegionInfo[] {
  if (!selector.by_region) throw new Error("set_region_text requires by_region selector.");
  const regions = selector.by_region.kind === "header" ? state.headers : state.footers;
  const section = selector.by_region.section ?? "all";
  if (section === "all") return regions;
  if (section === "first") return regions.length > 0 ? [regions[0]] : [];
  if (section === "last") return regions.length > 0 ? [regions[regions.length - 1]] : [];
  const numeric = typeof section === "number" ? section : parseInt(String(section), 10);
  if (!Number.isNaN(numeric)) {
    return regions[numeric - 1] ? [regions[numeric - 1]] : [];
  }
  throw new Error(`Unsupported region section selector: ${section}`);
}

function replaceRange(xml: string, start: number, end: number, replacement: string): string {
  return xml.slice(0, start) + replacement + xml.slice(end);
}

function rebuildFootnote(footnoteXml: string, text: string): string {
  const id = footnoteXml.match(/<w:footnote[^>]*w:id="([^"]+)"/)?.[1] ?? "1";
  return `<w:footnote w:id="${id}"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"></w:rStyle></w:rPr><w:footnoteRef></w:footnoteRef></w:r><w:r><w:t xml:space="preserve"> ${xmlEncode(text)}</w:t></w:r></w:p></w:footnote>`;
}

function rebuildCell(cellXml: string, text: string): string {
  const tcPr = cellXml.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/)?.[0] ?? "";
  const refParagraph = findBalancedTags(cellXml, "w:p")[0]?.xml ?? "<w:p></w:p>";
  return `<w:tc>${tcPr}${buildParagraphLike(refParagraph, text)}</w:tc>`;
}

function buildTableXml(data: string[][], tableStyle: string, headerRow: boolean): string {
  if (data.length === 0) throw new Error("insert_table requires at least one row.");
  const cols = Math.max(...data.map((row) => row.length));
  const width = Math.floor(9000 / Math.max(cols, 1));
  const grid = Array.from({ length: cols }, () => `<w:gridCol w:w="${width}"></w:gridCol>`).join("");
  const styleName = tableStyle === "professional" ? "TableGrid" : tableStyle === "minimal" ? "TableGrid" : tableStyle === "grid" ? "TableGrid" : "TableGrid";
  const rows = data.map((row, rowIndex) => {
    const cells = Array.from({ length: cols }, (_, i) => buildTableCellXml(row[i] ?? "", rowIndex === 0 && headerRow));
    const rowPr = rowIndex === 0 && headerRow ? `<w:trPr><w:tblHeader></w:tblHeader></w:trPr>` : "";
    return `<w:tr>${rowPr}${cells.join("")}</w:tr>`;
  }).join("");
  return `<w:tbl><w:tblPr><w:tblStyle w:val="${styleName}"></w:tblStyle><w:tblW w:w="0" w:type="auto"></w:tblW><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"></w:tblLook></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows}</w:tbl>`;
}

function ensureTableBlockSeparated(block: string, documentXml: string, insertAt: number, hasCaption: boolean): string {
  if (hasCaption) return block;
  const before = documentXml.slice(0, insertAt);
  const after = documentXml.slice(insertAt);
  const prevEndsWithTable = /<\/w:tbl>\s*$/.test(before);
  const nextStartsWithTable = /^\s*<w:tbl[\s>]/.test(after);
  let out = block;
  if (prevEndsWithTable) out = `<w:p></w:p>${out}`;
  if (nextStartsWithTable) out = `${out}<w:p></w:p>`;
  return out;
}

function buildClonedTableXml(sourceTable: TableInfo, data: string[][]): string {
  if (data.length === 0) throw new Error("insert_table requires at least one row.");
  const sourceRows = findBalancedTags(sourceTable.xml, "w:tr");
  if (sourceRows.length === 0) throw new Error(`Source table ${sourceTable.id} has no rows to clone.`);
  const templateColumnCount = Math.max(...sourceRows.map((row) => findBalancedTags(row.xml, "w:tc").length));
  const requestedColumnCount = Math.max(...data.map((row) => row.length));
  if (templateColumnCount !== requestedColumnCount) {
    throw new Error(
      `Source table ${sourceTable.id} has ${templateColumnCount} column(s), but new data has ${requestedColumnCount}. ` +
      `For now, clone_style_from_table requires matching column counts.`
    );
  }

  const tblPr = sourceTable.xml.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/)?.[0] ?? "";
  const tblGrid = sourceTable.xml.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/)?.[0] ?? "";
  const templateRows = sourceRows.map((row) => row.xml);
  const bodyTemplates = templateRows.slice(1);
  const rowsXml = data.map((values, rowIndex) => {
    const templateRow = rowIndex === 0 ? templateRows[0] : bodyTemplates[(rowIndex - 1) % Math.max(bodyTemplates.length, 1)] ?? templateRows[templateRows.length - 1];
    return rebuildTableRowFromTemplate(templateRow, values);
  }).join("");

  return `<w:tbl>${tblPr}${tblGrid}${rowsXml}</w:tbl>`;
}

function rebuildTableRowFromTemplate(templateRowXml: string, values: string[]): string {
  const rowPr = templateRowXml.match(/<w:trPr>[\s\S]*?<\/w:trPr>/)?.[0] ?? "";
  const templateCells = findBalancedTags(templateRowXml, "w:tc").map((cell) => cell.xml);
  const cellsXml = values.map((value, index) => {
    const templateCell = templateCells[index] ?? templateCells[templateCells.length - 1];
    return replaceCellTextKeepFormatting(templateCell, value);
  }).join("");
  return `<w:tr>${rowPr}${cellsXml}</w:tr>`;
}

function replaceCellTextKeepFormatting(cellXml: string, text: string): string {
  const currentCellText = extractAllText(cellXml);
  if (currentCellText.length === 0) return rebuildCell(cellXml, text);
  const replacement = replaceTextInXml(cellXml, currentCellText, text, false);
  return replacement.count === 1 ? replacement.xml : rebuildCell(cellXml, text);
}

function buildTableCellXml(text: string, isHeader: boolean): string {
  const escaped = xmlEncode(text);
  const rPr = isHeader ? `<w:rPr><w:b></w:b></w:rPr>` : "";
  const tcPr = `<w:tcPr><w:tcW w:w="0" w:type="auto"></w:tcW></w:tcPr>`;
  return `<w:tc>${tcPr}<w:p><w:r>${rPr}<w:t>${escaped}</w:t></w:r></w:p></w:tc>`;
}

function buildTableRowXml(tableXml: string, values: string[]): string {
  const rows = findBalancedTags(tableXml, "w:tr");
  const templateRow = rows[rows.length - 1]?.xml;
  if (!templateRow) {
    return `<w:tr>${values.map((v) => buildTableCellXml(v, false)).join("")}</w:tr>`;
  }
  const cells = findBalancedTags(templateRow, "w:tc");
  const newCells = Array.from({ length: Math.max(values.length, cells.length) }, (_, i) => {
    const templateCell = cells[i]?.xml ?? cells[0]?.xml ?? "<w:tc><w:p></w:p></w:tc>";
    return rebuildCell(templateCell, values[i] ?? "");
  }).join("");
  const rowPr = templateRow.match(/<w:trPr>[\s\S]*?<\/w:trPr>/)?.[0] ?? "";
  return `<w:tr>${rowPr}${newCells}</w:tr>`;
}

function fillFieldInState(state: MutableState, name: string, value: string): Record<string, unknown> {
  const matches: Array<{ xmlPath: string; fullName: string }> = [];
  for (const [xmlPath, xml] of state.files) {
    for (const field of findSdtFields(xml, xmlPath)) {
      if (matchField(field, name)) {
        const updated = fillSdtField(xml, field, value);
        state.files.set(xmlPath, updated);
        matches.push({ xmlPath, fullName: fieldDisplayName(field) });
      }
    }
  }
  if (matches.length === 0) throw new Error(`Field not found: ${name}`);
  if (matches.length > 1) throw new Error(`Field name matched multiple fields: ${name}`);
  return { field: matches[0].fullName, xml_path: matches[0].xmlPath };
}

function setCheckboxInState(state: MutableState, name: string, checked: boolean): Record<string, unknown> {
  const matches: Array<{ xmlPath: string; fullName: string }> = [];
  for (const [xmlPath, xml] of state.files) {
    for (const field of findSdtFields(xml, xmlPath)) {
      if (matchField(field, name)) {
        const updated = checkSdtField(xml, field, checked);
        state.files.set(xmlPath, updated);
        matches.push({ xmlPath, fullName: fieldDisplayName(field) });
      }
    }
  }
  if (matches.length === 0) throw new Error(`Checkbox field not found: ${name}`);
  if (matches.length > 1) throw new Error(`Field name matched multiple fields: ${name}`);
  return { field: matches[0].fullName, xml_path: matches[0].xmlPath, checked };
}

function insertFootnote(state: MutableState, selector: Selector, footnoteText: string): Record<string, unknown> {
  const paragraph = resolveParagraph(state, selector);
  if (paragraph.xmlPath !== "word/document.xml") throw new Error("insert_footnote currently supports body paragraphs only.");
  if (!selector.by_text) throw new Error("insert_footnote currently requires a by_text selector.");
  const anchorText = selector.by_text.text;
  const documentXml = state.files.get("word/document.xml")!;
  const nextId = ensureFootnoteInfrastructure(state);
  const updatedParagraph = insertFootnoteIntoParagraphXml(documentXml, paragraph, anchorText, nextId);
  state.files.set("word/document.xml", replaceRange(documentXml, paragraph.start, paragraph.end, updatedParagraph));
  const footnotesXml = state.files.get("word/footnotes.xml")!;
  state.files.set("word/footnotes.xml", appendFootnoteXml(footnotesXml, nextId, footnoteText));
  return { footnote_id: `fn${nextId}`, anchor_paragraph_id: paragraph.id };
}

function ensureFootnoteInfrastructure(state: MutableState): number {
  let footnotesXml = state.files.get("word/footnotes.xml");
  if (!footnotesXml) {
    footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>`;
    state.files.set("word/footnotes.xml", footnotesXml);
    const relsPath = "word/_rels/document.xml.rels";
    const rels = state.files.get(relsPath) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    if (!/relationships\/footnotes/.test(rels)) {
      const nextRelId = nextRelationshipId(rels);
      state.files.set(
        relsPath,
        rels.replace(
          /<\/Relationships>/,
          `<Relationship Id="rId${nextRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>`
        )
      );
    }
    const contentTypes = state.files.get("[Content_Types].xml");
    if (contentTypes && !/PartName="\/word\/footnotes.xml"/.test(contentTypes)) {
      state.files.set(
        "[Content_Types].xml",
        contentTypes.replace(
          /<\/Types>/,
          `<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>`
        )
      );
    }
  }
  const ids = [...(state.files.get("word/footnotes.xml") ?? "").matchAll(/<w:footnote[^>]*w:id="(-?\d+)"/g)].map((m) => parseInt(m[1], 10)).filter((n) => n > 0);
  return (ids.length ? Math.max(...ids) : 0) + 1;
}

function insertFootnoteIntoParagraphXml(documentXml: string, paragraph: ParagraphInfo, anchorText: string, footnoteId: number): string {
  const paragraphXml = mergeAdjacentRuns(paragraph.xml);
  const textRegex = /<w:t([^>]*)>([^<]*)<\/w:t>/g;
  let match: RegExpExecArray | null;
  const tElements: Array<{ attrs: string; rawText: string; decoded: string; start: number; end: number }> = [];
  while ((match = textRegex.exec(paragraphXml)) !== null) {
    tElements.push({ attrs: match[1], rawText: match[2], decoded: xmlDecodeSafe(match[2]), start: match.index, end: match.index + match[0].length });
  }
  const target = tElements.find((t) => t.decoded.includes(anchorText));
  if (!target) throw new Error(`Anchor text not found in paragraph: ${anchorText}`);
  const localIndex = target.decoded.indexOf(anchorText) + anchorText.length;
  const runStarts = [...paragraphXml.matchAll(/<w:r(?:\s|>)/g)].map((m) => m.index ?? -1).filter((idx) => idx >= 0 && idx < target.start);
  const runStart = runStarts.length > 0 ? runStarts[runStarts.length - 1] : -1;
  const runEnd = paragraphXml.indexOf("</w:r>", target.end);
  if (runStart === -1 || runEnd === -1) throw new Error("Could not isolate footnote anchor run.");
  const runXml = paragraphXml.slice(runStart, runEnd + 6);
  const attrs = runXml.match(/^<w:r([^>]*)>/)?.[1] ?? "";
  const rPr = runXml.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? "";
  const beforeText = target.decoded.slice(0, localIndex);
  const afterText = target.decoded.slice(localIndex);
  const beforeRun = `<w:r${attrs}>${rPr}<w:t${needsXmlSpace(beforeText)}>${xmlEncode(beforeText)}</w:t></w:r>`;
  const footnoteRun = buildFootnoteReferenceRun(documentXml, footnoteId);
  const afterRun = afterText.length > 0 ? `<w:r${attrs}>${rPr}<w:t${needsXmlSpace(afterText)}>${xmlEncode(afterText)}</w:t></w:r>` : "";
  const updatedRun = replaceRange(runXml, target.start - runStart, target.end - runStart, `<w:t${target.attrs}>${xmlEncode(beforeText)}</w:t>`);
  const rebuiltRun = beforeRun + footnoteRun + afterRun;
  return paragraphXml.slice(0, runStart) + rebuiltRun + paragraphXml.slice(runEnd + 6);
}

function buildFootnoteReferenceRun(documentXml: string, footnoteId: number): string {
  const existing = documentXml.match(/<w:r[^>]*>[\s\S]*?<w:rPr>[\s\S]*?<w:rStyle w:val="FootnoteReference"><\/w:rStyle>[\s\S]*?<\/w:rPr>[\s\S]*?<w:footnoteReference w:id="\d+"><\/w:footnoteReference>[\s\S]*?<\/w:r>/);
  if (existing) {
    return existing[0].replace(/w:id="\d+"/, `w:id="${footnoteId}"`);
  }
  return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"></w:rStyle></w:rPr><w:footnoteReference w:id="${footnoteId}"></w:footnoteReference></w:r>`;
}

function appendFootnoteXml(footnotesXml: string, footnoteId: number, text: string): string {
  const template = footnotesXml.match(/<w:footnote w:id="\d+">[\s\S]*?<\/w:footnote>/);
  if (template) {
    const cloned = template[0]
      .replace(/w:id="\d+"/, `w:id="${footnoteId}"`)
      .replace(/<w:t[^>]*>[\s\S]*?<\/w:t>/, `<w:t xml:space="preserve"> ${xmlEncode(text)}</w:t>`);
    return footnotesXml.replace(/<\/w:footnotes>\s*$/, `${cloned}</w:footnotes>`);
  }
  const note = `<w:footnote w:id="${footnoteId}"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"></w:rStyle></w:rPr><w:footnoteRef></w:footnoteRef></w:r><w:r><w:t xml:space="preserve"> ${xmlEncode(text)}</w:t></w:r></w:p></w:footnote>`;
  return footnotesXml.replace(/<\/w:footnotes>\s*$/, `${note}</w:footnotes>`);
}

function nextRelationshipId(relsXml: string): number {
  const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
  return (ids.length ? Math.max(...ids) : 0) + 1;
}

async function insertImage(
  state: MutableState,
  op: Extract<PatchOperation, { op: "insert_image" }>,
  binaryEntries: Map<string, Buffer>
): Promise<Record<string, unknown>> {
  let { anchor, mode } = resolveLocationParagraph(state, op.location);
  if (anchor.xmlPath !== "word/document.xml") throw new Error("insert_image currently supports body locations only.");
  const imageBuffer = await readFile(op.file_path);
  const ext = extname(op.file_path).replace(/^\./, "").toLowerCase();
  if (!ext) throw new Error(`Could not determine image extension for ${op.file_path}`);
  const imageName = nextMediaName(state, basename(op.file_path));
  const mediaPath = `word/media/${imageName}`;
  binaryEntries.set(mediaPath, imageBuffer);

  ensureImageContentType(state, ext);
  ensureDrawingNamespaces(state);
  rebuildState(state);
  ({ anchor, mode } = resolveLocationParagraph(state, op.location));
  if (anchor.xmlPath !== "word/document.xml") throw new Error("insert_image currently supports body locations only.");

  const relsPath = "word/_rels/document.xml.rels";
  const relsXml = state.files.get(relsPath) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const relId = `rId${nextRelationshipId(relsXml)}`;
  state.files.set(relsPath, relsXml.replace(/<\/Relationships>/, `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${imageName}"/></Relationships>`));

  const size = getImageDimensions(imageBuffer, ext);
  const widthPx = op.width_px ?? size.width;
  const heightPx = op.height_px ?? (op.preserve_aspect_ratio === false ? size.height : Math.round(widthPx * (size.height / size.width)));
  const drawingId = nextDocPrId(state.files.get("word/document.xml")!);
  const layoutSource = op.clone_layout_from_paragraph ? resolveParagraphById(state, op.clone_layout_from_paragraph) : undefined;
  const imageParagraph = buildImageParagraphXml({
    relId,
    name: imageName,
    alt: op.alt_text ?? imageName,
    widthPx,
    heightPx,
    drawingId,
    alignment: op.alignment ?? "center",
    referenceParagraph: layoutSource,
  });
  const captionSource = op.clone_caption_from_paragraph ? resolveParagraphById(state, op.clone_caption_from_paragraph) : undefined;
  const captionXml = op.caption ? buildCaptionParagraphXml(op.caption, captionSource) : "";
  const insertBlock = mode === "after" ? imageParagraph + captionXml : captionXml + imageParagraph;
  const xml = state.files.get("word/document.xml")!;
  const insertAt = mode === "after" ? anchor.end : anchor.start;
  state.files.set("word/document.xml", xml.slice(0, insertAt) + insertBlock + xml.slice(insertAt));
  return {
    image_path: mediaPath,
    rel_id: relId,
    width_px: widthPx,
    height_px: heightPx,
    cloned_layout_from_paragraph: op.clone_layout_from_paragraph ?? null,
    cloned_caption_from_paragraph: op.clone_caption_from_paragraph ?? null,
  };
}

function nextMediaName(state: MutableState, sourceName: string): string {
  const ext = extname(sourceName).toLowerCase();
  const base = basename(sourceName, ext).replace(/[^a-zA-Z0-9_-]/g, "-") || "image";
  const existing = [...state.files.keys()].filter((p) => p.startsWith("word/media/")).map((p) => basename(p));
  let i = 1;
  let candidate = `${base}${ext}`;
  while (existing.includes(candidate)) {
    candidate = `${base}-${i++}${ext}`;
  }
  return candidate;
}

function ensureImageContentType(state: MutableState, ext: string): void {
  const contentTypes = state.files.get("[Content_Types].xml");
  if (!contentTypes) return;
  if (new RegExp(`<Default Extension="${ext}"`).test(contentTypes)) return;
  const mime = ext === "png" ? "image/png" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : ext === "webp" ? "image/webp" : `image/${ext}`;
  state.files.set("[Content_Types].xml", contentTypes.replace(/<\/Types>/, `<Default Extension="${ext}" ContentType="${mime}"/></Types>`));
}

function ensureDrawingNamespaces(state: MutableState): void {
  let xml = state.files.get("word/document.xml")!;
  const rootMatch = xml.match(/<w:document[^>]*>/);
  if (!rootMatch) return;
  let root = rootMatch[0];
  if (!root.includes('xmlns:a=')) root = root.replace('<w:document ', '<w:document xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ');
  if (!root.includes('xmlns:pic=')) root = root.replace('<w:document ', '<w:document xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ');
  state.files.set("word/document.xml", xml.replace(rootMatch[0], root));
}

function nextDocPrId(documentXml: string): number {
  const ids = [...documentXml.matchAll(/wp:docPr[^>]* id="(\d+)"/g)].map((m) => parseInt(m[1], 10));
  return (ids.length ? Math.max(...ids) : 0) + 1;
}

function buildImageParagraphXml(args: {
  relId: string;
  name: string;
  alt: string;
  widthPx: number;
  heightPx: number;
  drawingId: number;
  alignment: "left" | "center" | "right";
  referenceParagraph?: ParagraphInfo;
}): string {
  const cx = Math.round(args.widthPx * 9525);
  const cy = Math.round(args.heightPx * 9525);
  const jc = args.alignment === "left" ? "left" : args.alignment === "right" ? "right" : "center";
  const pPr = buildImageParagraphProperties(args.referenceParagraph?.xml, jc);
  return `<w:p>${pPr}<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"></wp:extent><wp:docPr id="${args.drawingId}" name="Picture ${args.drawingId}" descr="${xmlEncode(args.alt)}"></wp:docPr><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"></a:graphicFrameLocks></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${xmlEncode(args.name)}" descr="${xmlEncode(args.alt)}"></pic:cNvPr><pic:cNvPicPr></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="${args.relId}"></a:blip><a:stretch><a:fillRect></a:fillRect></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"></a:off><a:ext cx="${cx}" cy="${cy}"></a:ext></a:xfrm><a:prstGeom prst="rect"><a:avLst></a:avLst></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

function buildImageParagraphProperties(referenceParagraphXml: string | undefined, alignment: "left" | "center" | "right"): string {
  const jc = `<w:jc w:val="${alignment}"></w:jc>`;
  if (!referenceParagraphXml) {
    return `<w:pPr>${jc}</w:pPr>`;
  }
  const pPr = referenceParagraphXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0];
  if (!pPr) return `<w:pPr>${jc}</w:pPr>`;
  let next = pPr.replace(/<w:pStyle[^>]*>[\s\S]*?<\/w:pStyle>/g, "");
  next = next.replace(/<w:jc[^>]*>[\s\S]*?<\/w:jc>/g, "");
  next = next.replace(/<w:jc[^>]*/g, "");
  return next.replace("</w:pPr>", `${jc}</w:pPr>`);
}

function buildCaptionParagraphXml(caption: string, sourceParagraph?: ParagraphInfo): string {
  if (sourceParagraph) {
    return buildParagraphLike(sourceParagraph.xml, caption);
  }
  return `<w:p><w:pPr><w:jc w:val="center"></w:jc></w:pPr><w:r><w:rPr><w:i></w:i><w:color w:val="000000"></w:color><w:sz w:val="20"></w:sz><w:szCs w:val="20"></w:szCs></w:rPr><w:t>${xmlEncode(caption)}</w:t></w:r></w:p>`;
}

function getImageDimensions(buffer: Buffer, ext: string): { width: number; height: number } {
  if (ext === "png") {
    return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
  }
  if (ext === "gif") {
    return { width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8) };
  }
  if (ext === "jpg" || ext === "jpeg") {
    let offset = 2;
    while (offset < buffer.length) {
      if (buffer[offset] !== 0xff) {
        offset++;
        continue;
      }
      const marker = buffer[offset + 1];
      const length = buffer.readUInt16BE(offset + 2);
      if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
        return { height: buffer.readUInt16BE(offset + 5), width: buffer.readUInt16BE(offset + 7) };
      }
      offset += 2 + length;
    }
  }
  throw new Error(`Unsupported image format for size detection: .${ext}`);
}

function needsXmlSpace(text: string): string {
  return text.startsWith(" ") || text.endsWith(" ") || text.includes("  ") ? ' xml:space="preserve"' : "";
}

function xmlDecodeSafe(text: string): string {
  return text
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}
