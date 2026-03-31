import { XMLParser, XMLBuilder } from "fast-xml-parser";

// Parser configured to preserve the structure we need
const parserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  trimValues: false,
  // Keep CDATA, processing instructions
  cdataPropName: "__cdata",
  commentPropName: "__comment",
};

const builderOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  format: true,
  indentBy: "  ",
  suppressEmptyNode: false,
};

const parser = new XMLParser(parserOptions);
const builder = new XMLBuilder(builderOptions);

/**
 * Pretty-print XML with indentation for agent readability.
 * Uses fast-xml-parser round-trip to normalize formatting.
 */
export function prettyPrintXml(xml: string): string {
  try {
    const parsed = parser.parse(xml);
    return builder.build(parsed);
  } catch {
    // If parsing fails, return as-is
    return xml;
  }
}

/**
 * OOXML fragments text across multiple <w:r> (run) elements due to
 * spell-check, formatting changes, revision tracking, etc.
 * 
 * Example input:
 *   <w:r><w:t>Hel</w:t></w:r>
 *   <w:r><w:t>lo wo</w:t></w:r>
 *   <w:r><w:t>rld</w:t></w:r>
 * 
 * After merging (same formatting):
 *   <w:r><w:t>Hello world</w:t></w:r>
 * 
 * This is critical for agents: fragmented runs make text search/replace
 * nearly impossible since the target string spans multiple XML elements.
 */
export function mergeAdjacentRuns(xml: string): string {
  // We use regex-based merging here rather than DOM manipulation.
  // This is intentionally conservative: only merge runs with identical
  // formatting (rPr) to avoid destroying style boundaries.

  // Pattern: match consecutive <w:r>...</w:r> elements
  const runPattern = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;

  // Extract run properties and text from a single run
  function parseRun(runContent: string): { rPr: string; texts: string[] } | null {
    const rPrMatch = runContent.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
    const rPr = rPrMatch ? rPrMatch[0] : "";

    const texts: string[] = [];
    const textPattern = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
    let m;
    while ((m = textPattern.exec(runContent)) !== null) {
      texts.push(m[1]);
    }

    if (texts.length === 0) return null;
    return { rPr, texts };
  }

  // Find paragraphs and process each independently
  return xml.replace(
    /(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/g,
    (_match, pOpen: string, pContent: string, pClose: string) => {
      // Collect all runs in this paragraph
      const segments: Array<{ isRun: boolean; content: string; rPr?: string; texts?: string[] }> = [];
      let lastIndex = 0;

      const runRegex = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;
      let rm;
      while ((rm = runRegex.exec(pContent)) !== null) {
        // Capture non-run content before this run
        if (rm.index > lastIndex) {
          segments.push({ isRun: false, content: pContent.slice(lastIndex, rm.index) });
        }
        const parsed = parseRun(rm[1]);
        if (parsed) {
          segments.push({ isRun: true, content: rm[0], rPr: parsed.rPr, texts: parsed.texts });
        } else {
          // Run without text (e.g., images, breaks) - keep as-is
          segments.push({ isRun: false, content: rm[0] });
        }
        lastIndex = rm.index + rm[0].length;
      }
      if (lastIndex < pContent.length) {
        segments.push({ isRun: false, content: pContent.slice(lastIndex) });
      }

      // Merge consecutive runs with identical rPr
      const merged: typeof segments = [];
      for (const seg of segments) {
        const prev = merged[merged.length - 1];
        if (
          seg.isRun &&
          prev?.isRun &&
          seg.rPr === prev.rPr &&
          seg.texts &&
          prev.texts
        ) {
          // Merge texts
          prev.texts = [...prev.texts, ...seg.texts];
          const combinedText = prev.texts.join("");
          const xmlSpace = combinedText.startsWith(" ") || combinedText.endsWith(" ")
            ? ' xml:space="preserve"'
            : "";
          prev.content = `<w:r>${prev.rPr}<w:t${xmlSpace}>${combinedText}</w:t></w:r>`;
        } else {
          merged.push({ ...seg });
        }
      }

      return pOpen + merged.map((s) => s.content).join("") + pClose;
    }
  );
}

/**
 * Normalize smart quotes and special characters to XML-safe entities.
 * Agents sometimes struggle with curly quotes in XML context.
 */
export function normalizeQuotes(xml: string): string {
  return xml
    .replace(/\u201C/g, "&#x201C;") // left double quote
    .replace(/\u201D/g, "&#x201D;") // right double quote
    .replace(/\u2018/g, "&#x2018;") // left single quote
    .replace(/\u2019/g, "&#x2019;") // right single quote
    .replace(/\u2013/g, "&#x2013;") // en dash
    .replace(/\u2014/g, "&#x2014;") // em dash
    .replace(/\u2026/g, "&#x2026;"); // ellipsis
}

/**
 * Encode text for safe inclusion in XML text nodes.
 * Only encodes the three characters that MUST be escaped in XML text content.
 */
export function xmlEncode(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Decode XML entities back to plain text.
 * Handles named entities, decimal, and hex character references.
 */
export function xmlDecode(text: string): string {
  return text
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}

/**
 * Extract all plain text from an XML fragment by reading <w:t> and <m:t> elements.
 */
export function extractAllText(xml: string): string {
  const texts: string[] = [];
  const tRegex = /<(?:w|m):t[^>]*>([^<]*)<\/(?:w|m):t>/g;
  let m;
  while ((m = tRegex.exec(xml)) !== null) {
    texts.push(xmlDecode(m[1]));
  }
  return texts.join("");
}

/**
 * Normalize smart/curly quotes and typographic characters to their ASCII equivalents.
 * This allows agents to match text using regular quotes even when the document
 * contains smart quotes (which is extremely common in Word docs).
 */
export function normalizeForSearch(text: string): string {
  return text
    .replace(/[\u2018\u2019\u201A\u201B]/g, "'") // smart single quotes → '
    .replace(/[\u201C\u201D\u201E\u201F]/g, '"') // smart double quotes → "
    .replace(/[\u2013\u2014]/g, "-")              // en/em dash → -
    .replace(/\u2026/g, "...")                     // ellipsis → ...
    .replace(/\u00A0/g, " ");                      // non-breaking space → space
}

// ── Text element info within a paragraph ───────────────────────────────────────

interface TElement {
  /** Full match string <w:t...>...</w:t> */
  match: string;
  /** Attributes string (e.g. ' xml:space="preserve"') */
  attrs: string;
  /** Raw XML-encoded content between tags */
  rawContent: string;
  /** Decoded text content */
  decoded: string;
  /** Start position of this <w:t> element in the parent string */
  startInXml: number;
  /** End position of this <w:t> element in the parent string */
  endInXml: number;
  /** Start position in concatenated paragraph text */
  textStart: number;
  /** End position in concatenated paragraph text */
  textEnd: number;
}

/**
 * Find all <w:t> elements in an XML string and return their info.
 */
function findTElements(xml: string): TElement[] {
  const elements: TElement[] = [];
  const tRegex = /<w:t([^>]*)>([^<]*)<\/w:t>/g;
  let m;
  let textPos = 0;

  while ((m = tRegex.exec(xml)) !== null) {
    const decoded = xmlDecode(m[2]);
    elements.push({
      match: m[0],
      attrs: m[1],
      rawContent: m[2],
      decoded,
      startInXml: m.index,
      endInXml: m.index + m[0].length,
      textStart: textPos,
      textEnd: textPos + decoded.length,
    });
    textPos += decoded.length;
  }

  return elements;
}

/**
 * Build a new <w:t> element with proper xml:space handling.
 */
function buildTElement(text: string, originalAttrs: string): string {
  const encoded = xmlEncode(text);
  let attrs = originalAttrs;
  const needsSpace =
    text.startsWith(" ") || text.endsWith(" ") || text.includes("  ");
  if (needsSpace && !attrs.includes("xml:space")) {
    attrs = ' xml:space="preserve"';
  }
  return `<w:t${attrs}>${encoded}</w:t>`;
}

/**
 * Replace text in OOXML content, handling:
 * - Run fragmentation (merges adjacent same-format runs)
 * - Cross-run text (text split by proofErr, bookmarks, comments, etc.)
 * - Smart quote normalization (regular quotes match smart quotes)
 * - XML entity encoding/decoding
 *
 * Strategy:
 * 1. Merge adjacent same-format runs
 * 2. Process each paragraph: concatenate all <w:t> text and search
 * 3. Replace across run boundaries when needed
 *
 * @returns The modified XML and number of replacements made.
 */
export function replaceTextInXml(
  xml: string,
  oldText: string,
  newText: string,
  replaceAll: boolean = false
): { xml: string; count: number } {
  // Step 1: merge adjacent same-format runs so text is contiguous where possible
  const merged = mergeAdjacentRuns(xml);

  // Normalize oldText for smart-quote-insensitive matching
  const oldNorm = normalizeForSearch(oldText);

  let totalCount = 0;

  // Step 2: process paragraph by paragraph
  // We match <w:p>...</w:p> and process each independently
  const result = merged.replace(
    /(<w:p\b[^>]*>)([\s\S]*?)(<\/w:p>)/g,
    (fullMatch, pOpen: string, pContent: string, pClose: string) => {
      if (!replaceAll && totalCount >= 1) return fullMatch;

      // Find and replace within this paragraph (may loop for replaceAll)
      const { content: newContent, count } = replaceInParagraph(
        pContent,
        oldText,
        oldNorm,
        newText,
        replaceAll ? Infinity : 1 - totalCount
      );

      if (count > 0) {
        totalCount += count;
        return pOpen + newContent + pClose;
      }

      return fullMatch;
    }
  );

  return { xml: result, count: totalCount };
}

/**
 * Find and replace text within a single paragraph's content.
 * Handles both single-run and cross-run replacements.
 * Re-scans after each replacement to keep positions correct.
 */
function replaceInParagraph(
  pContent: string,
  oldText: string,
  oldNorm: string,
  newText: string,
  maxReplacements: number
): { content: string; count: number } {
  let content = pContent;
  let count = 0;

  while (count < maxReplacements) {
    // Re-scan <w:t> elements (positions change after each replacement)
    const tElements = findTElements(content);
    if (tElements.length === 0) break;

    // Concatenate paragraph text and normalize for search
    const paraText = tElements.map((t) => t.decoded).join("");
    const paraNorm = normalizeForSearch(paraText);

    // Find the first occurrence
    const idx = paraNorm.indexOf(oldNorm);
    if (idx === -1) break;

    const matchEnd = idx + oldNorm.length;

    // Find affected <w:t> elements
    const affected = tElements.filter(
      (t) => t.textEnd > idx && t.textStart < matchEnd
    );
    if (affected.length === 0) break;

    // Apply replacement from last affected element to first (to preserve positions)
    for (let j = affected.length - 1; j >= 0; j--) {
      const t = affected[j];
      let newDecoded: string;

      if (affected.length === 1) {
        // Entire match within a single <w:t>
        const localStart = idx - t.textStart;
        const localEnd = matchEnd - t.textStart;
        newDecoded =
          t.decoded.substring(0, localStart) +
          newText +
          t.decoded.substring(localEnd);
      } else if (j === 0) {
        // First element: keep text before match, append replacement
        newDecoded =
          t.decoded.substring(0, idx - t.textStart) + newText;
      } else if (j === affected.length - 1) {
        // Last element: keep text after match
        newDecoded = t.decoded.substring(matchEnd - t.textStart);
      } else {
        // Middle element: clear text
        newDecoded = "";
      }

      const newT = buildTElement(newDecoded, t.attrs);
      content =
        content.substring(0, t.startInXml) +
        newT +
        content.substring(t.endInXml);
    }

    count++;
  }

  return { content, count };
}
