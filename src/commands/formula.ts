import { detectFileType } from "../utils/detect.js";
import { repackDocx } from "../utils/repack.js";
import { loadMutableWordState, getBalancedTagRange, type ParagraphInfo } from "../utils/wordModel.js";
import { buildFormulaParagraphXml, ensureMathNamespace, type FormulaAlignment } from "../utils/math.js";

export interface FormulaOptions {
  latex: string;
  after?: string;
  before?: string;
  replace?: string;
  append?: boolean;
  output?: string;
  align?: FormulaAlignment;
  match?: "exact" | "contains" | "regex";
  occurrence?: number;
}

export async function runFormula(filePath: string, options: FormulaOptions): Promise<void> {
  const { type } = await detectFileType(filePath);
  if (type !== "docx") {
    throw new Error(`Only .docx files are supported. Detected format: ${type}`);
  }

  const modeCount = [options.after, options.before, options.replace, options.append].filter(Boolean).length;
  if (modeCount !== 1) {
    throw new Error("Choose exactly one location mode: --after, --before, --replace, or --append.");
  }

  if (options.align && !["left", "center", "right"].includes(options.align)) {
    throw new Error(`Unsupported alignment: ${options.align}. Use left, center, or right.`);
  }
  if (options.match && !["exact", "contains", "regex"].includes(options.match)) {
    throw new Error(`Unsupported match mode: ${options.match}. Use exact, contains, or regex.`);
  }
  if (!Number.isInteger(options.occurrence ?? 1) || (options.occurrence ?? 1) < 1) {
    throw new Error("--occurrence must be a positive integer.");
  }

  const state = await loadMutableWordState(filePath);
  const originalDocumentXml = state.files.get("word/document.xml");
  if (!originalDocumentXml) {
    throw new Error("Document is missing word/document.xml.");
  }

  const formulaParagraphXml = await buildFormulaParagraphXml(options.latex, options.align ?? "center");
  let documentXml = originalDocumentXml;

  if (options.append) {
    documentXml = insertAtBodyEnd(documentXml, formulaParagraphXml);
  } else if (options.after) {
    const paragraph = resolveBodyParagraph(state.paragraphs, options.after, options.match ?? "contains", options.occurrence ?? 1);
    documentXml = replaceRange(documentXml, paragraph.end, paragraph.end, formulaParagraphXml);
  } else if (options.before) {
    const paragraph = resolveBodyParagraph(state.paragraphs, options.before, options.match ?? "contains", options.occurrence ?? 1);
    documentXml = replaceRange(documentXml, paragraph.start, paragraph.start, formulaParagraphXml);
  } else if (options.replace) {
    const paragraph = resolveBodyParagraph(state.paragraphs, options.replace, options.match ?? "exact", options.occurrence ?? 1);
    documentXml = replaceRange(documentXml, paragraph.start, paragraph.end, formulaParagraphXml);
  }

  documentXml = ensureMathNamespace(documentXml);
  await repackDocx(filePath, new Map([["word/document.xml", documentXml]]), options.output);

  const target = options.output ?? filePath;
  console.error(`✓ Added formula → ${target}`);
}

function resolveBodyParagraph(
  paragraphs: ParagraphInfo[],
  query: string,
  match: "exact" | "contains" | "regex",
  occurrence: number
): ParagraphInfo {
  const bodyParagraphs = paragraphs.filter((p) => p.kind === "body");
  const matches = bodyParagraphs.filter((p) => matchesText(p.text, query, match));
  if (matches.length === 0) {
    throw new Error(`No body paragraph matched: ${query}`);
  }
  if (matches.length < occurrence) {
    throw new Error(`Matched ${matches.length} paragraph(s), but occurrence ${occurrence} was requested.`);
  }
  if (occurrence === 1 && matches.length > 1) {
    const preview = matches.slice(0, 10).map((p) => `  ${p.id}: ${shorten(p.text)}`).join("\n");
    throw new Error(
      `Paragraph selector matched ${matches.length} paragraphs. Add more context or pass --occurrence.\nMatches:\n${preview}`
    );
  }
  return matches[occurrence - 1];
}

function matchesText(text: string, query: string, match: "exact" | "contains" | "regex"): boolean {
  if (match === "regex") return new RegExp(query).test(text);
  if (match === "contains") return text.includes(query);
  return text === query;
}

function insertAtBodyEnd(documentXml: string, blockXml: string): string {
  const body = getBalancedTagRange(documentXml, "w:body");
  if (!body) throw new Error("Document is missing <w:body>.");
  const sectPrIndex = body.innerXml.lastIndexOf("<w:sectPr");
  const insertAt = sectPrIndex === -1 ? body.innerEnd : body.innerStart + sectPrIndex;
  return replaceRange(documentXml, insertAt, insertAt, blockXml);
}

function replaceRange(xml: string, start: number, end: number, replacement: string): string {
  return xml.slice(0, start) + replacement + xml.slice(end);
}

function shorten(text: string, max = 80): string {
  return text.length > max ? `${text.slice(0, max - 3)}...` : text;
}
