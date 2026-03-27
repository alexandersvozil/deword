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
 * Extract all plain text from an XML fragment by reading <w:t> elements.
 */
export function extractAllText(xml: string): string {
  const texts: string[] = [];
  const tRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let m;
  while ((m = tRegex.exec(xml)) !== null) {
    texts.push(xmlDecode(m[1]));
  }
  return texts.join("");
}

/**
 * Replace text in OOXML content, handling run fragmentation and XML encoding.
 *
 * Strategy:
 * 1. Merge adjacent runs with identical formatting (so fragmented text is contiguous)
 * 2. For each <w:t> element, decode content and search for oldText
 * 3. Replace and re-encode, preserving xml:space attributes
 *
 * @returns The modified XML and number of replacements made.
 */
export function replaceTextInXml(
  xml: string,
  oldText: string,
  newText: string,
  replaceAll: boolean = false
): { xml: string; count: number } {
  // Step 1: merge adjacent same-format runs so text is contiguous
  const merged = mergeAdjacentRuns(xml);

  // Step 2: count total occurrences across all <w:t> elements
  let totalCount = 0;
  const tCountRegex = /<w:t([^>]*)>([^<]*)<\/w:t>/g;
  let cm;
  while ((cm = tCountRegex.exec(merged)) !== null) {
    const decoded = xmlDecode(cm[2]);
    let pos = 0;
    while ((pos = decoded.indexOf(oldText, pos)) !== -1) {
      totalCount++;
      pos += oldText.length;
    }
  }

  if (totalCount === 0) {
    return { xml: merged, count: 0 };
  }

  // Step 3: do the replacement
  let replaced = 0;
  const result = merged.replace(
    /<w:t([^>]*)>([^<]*)<\/w:t>/g,
    (match, attrs: string, content: string) => {
      if (!replaceAll && replaced >= 1) return match;

      const decoded = xmlDecode(content);
      if (!decoded.includes(oldText)) return match;

      let newDecoded: string;
      if (replaceAll) {
        const parts = decoded.split(oldText);
        replaced += parts.length - 1;
        newDecoded = parts.join(newText);
      } else {
        // Replace only the first occurrence
        const idx = decoded.indexOf(oldText);
        newDecoded =
          decoded.substring(0, idx) + newText + decoded.substring(idx + oldText.length);
        replaced++;
      }

      const encoded = xmlEncode(newDecoded);

      // Preserve or add xml:space="preserve" when text has significant whitespace
      let newAttrs = attrs;
      const needsSpace =
        newDecoded.startsWith(" ") || newDecoded.endsWith(" ") || newDecoded.includes("  ");
      if (needsSpace && !attrs.includes("xml:space")) {
        newAttrs = ' xml:space="preserve"';
      }

      return `<w:t${newAttrs}>${encoded}</w:t>`;
    }
  );

  return { xml: result, count: replaced };
}
