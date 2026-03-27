import { xmlDecode, xmlEncode } from "./xml.js";

// ── Types ──────────────────────────────────────────────────────────────────────

export interface SdtField {
  /** Programmatic tag name from <w:tag w:val="..."/> */
  tag: string | null;
  /** Display alias from <w:alias w:val="..."/> */
  alias: string | null;
  /** Field type */
  type: "text" | "date" | "dropdown" | "combobox" | "checkbox" | "richtext" | "unknown";
  /** Current text content (decoded) */
  currentValue: string;
  /** Whether checkbox is checked */
  isChecked: boolean;
  /** Whether placeholder text is showing */
  isPlaceholder: boolean;
  /** Placeholder text if showing */
  placeholderText: string;
  /** Dropdown/combobox options */
  options: Array<{ display: string; value: string }>;
  /** Date format string */
  dateFormat: string | null;
  /** True if SDT is inside a paragraph (inline), false if block-level */
  isInline: boolean;
  /** Start position in the XML string */
  xmlStart: number;
  /** End position in the XML string */
  xmlEnd: number;
  /** The <w:sdtPr>...</w:sdtPr> XML */
  sdtPrXml: string;
  /** The <w:sdtContent>...</w:sdtContent> XML */
  sdtContentXml: string;
  /** The full <w:sdt>...</w:sdt> XML */
  fullXml: string;
  /** Which XML file this field is in */
  xmlPath: string;
}

// ── Discovery ──────────────────────────────────────────────────────────────────

/**
 * Find and parse all SDT content controls in an XML string.
 * Returns outermost SDTs only (nested ones are accessible via their parent).
 */
export function findSdtFields(xml: string, xmlPath: string): SdtField[] {
  const positions = findBalancedTags(xml, "w:sdt");
  const fields: SdtField[] = [];

  for (const pos of positions) {
    const sdtXml = xml.substring(pos.start, pos.end);
    const field = parseSdt(sdtXml, pos.start, xml, xmlPath);
    if (field) fields.push(field);

    // Also find nested SDTs inside sdtContent
    const nestedPositions = findBalancedTags(sdtXml.substring(sdtXml.indexOf("</w:sdtPr>") + 10), "w:sdt");
    for (const np of nestedPositions) {
      const nestedXml = sdtXml.substring(
        sdtXml.indexOf("</w:sdtPr>") + 10 + np.start,
        sdtXml.indexOf("</w:sdtPr>") + 10 + np.end
      );
      const nestedField = parseSdt(
        nestedXml,
        pos.start + sdtXml.indexOf("</w:sdtPr>") + 10 + np.start,
        xml,
        xmlPath
      );
      if (nestedField) fields.push(nestedField);
    }
  }

  return fields;
}

/**
 * Find balanced XML tags handling nesting. Returns outermost matches only.
 */
function findBalancedTags(xml: string, tagName: string): Array<{ start: number; end: number }> {
  const results: Array<{ start: number; end: number }> = [];
  const openPattern = new RegExp(`<${tagName}(?=[\\s>/])`, "g");
  const closeStr = `</${tagName}>`;

  // Collect all open and close positions
  type Event = { pos: number; type: "open" | "close" };
  const events: Event[] = [];

  let m;
  while ((m = openPattern.exec(xml)) !== null) {
    events.push({ pos: m.index, type: "open" });
  }

  let idx = 0;
  while ((idx = xml.indexOf(closeStr, idx)) !== -1) {
    events.push({ pos: idx, type: "close" });
    idx += closeStr.length;
  }

  events.sort((a, b) => a.pos - b.pos);

  // Match using a stack — collect only outermost (depth 0 → depth 1 → back to 0)
  const stack: number[] = [];
  for (const event of events) {
    if (event.type === "open") {
      stack.push(event.pos);
    } else if (stack.length > 0) {
      const openPos = stack.pop()!;
      if (stack.length === 0) {
        results.push({ start: openPos, end: event.pos + closeStr.length });
      }
    }
  }

  return results;
}

// ── Parsing ────────────────────────────────────────────────────────────────────

function parseSdt(
  sdtXml: string,
  xmlStart: number,
  fullDocXml: string,
  xmlPath: string
): SdtField | null {
  // Extract sdtPr
  const prMatch = sdtXml.match(/<w:sdtPr\b[^>]*>([\s\S]*?)<\/w:sdtPr>/);
  if (!prMatch) return null;
  const prXml = prMatch[0];
  const prContent = prMatch[1];

  // Extract sdtContent using balanced matching (handles nested SDTs in content)
  const contentStartTag = "<w:sdtContent";
  const csIdx = sdtXml.indexOf(contentStartTag);
  if (csIdx === -1) return null;

  const closeTag = "</w:sdtContent>";
  let depth = 0;
  let searchPos = csIdx;
  let contentEndIdx = -1;

  while (searchPos < sdtXml.length) {
    const nextOpen = sdtXml.indexOf(contentStartTag, searchPos + 1);
    const nextClose = sdtXml.indexOf(closeTag, searchPos + 1);
    if (nextClose === -1) break;

    if (nextOpen !== -1 && nextOpen < nextClose) {
      depth++;
      searchPos = nextOpen;
    } else {
      if (depth === 0) {
        contentEndIdx = nextClose + closeTag.length;
        break;
      }
      depth--;
      searchPos = nextClose;
    }
  }

  if (contentEndIdx === -1) return null;
  const contentXml = sdtXml.substring(csIdx, contentEndIdx);

  // Parse properties
  const tag = extractAttr(prContent, "w:tag", "w:val");
  const alias = extractAttr(prContent, "w:alias", "w:val");
  const showingPlcHdr = /<w:showingPlcHdr/.test(prContent);

  // Type detection
  let type: SdtField["type"] = "unknown";
  if (/<w14:checkbox/.test(prContent) || /<w:checkbox/.test(prContent)) type = "checkbox";
  else if (/<w:date[\s>]/.test(prContent)) type = "date";
  else if (/<w:dropDownList/.test(prContent)) type = "dropdown";
  else if (/<w:comboBox/.test(prContent)) type = "combobox";
  else if (/<w:text[\s/>]/.test(prContent)) type = "text";
  else type = "richtext";

  // Current value — extract all text from sdtContent
  const currentValue = extractTextContent(contentXml);

  // Checkbox state
  let isChecked = false;
  if (type === "checkbox") {
    const checkedMatch = prContent.match(/<w14:checked[^>]*w14:val="(\d+)"/);
    if (!checkedMatch) {
      // Try alternate format: <w14:checked w14:val="1"/>
      const altMatch = prContent.match(/<w14:checked[^>]*val="(\d+)"/);
      isChecked = altMatch ? altMatch[1] === "1" : false;
    } else {
      isChecked = checkedMatch[1] === "1";
    }
  }

  // Dropdown/combobox options
  const options: Array<{ display: string; value: string }> = [];
  if (type === "dropdown" || type === "combobox") {
    const listItemRegex = /<w:listItem[^>]*?w:displayText="([^"]*)"[^>]*?w:value="([^"]*)"/g;
    let lm;
    while ((lm = listItemRegex.exec(prContent)) !== null) {
      options.push({ display: lm[1], value: lm[2] });
    }
    // Also try reversed attribute order
    const listItemRegex2 = /<w:listItem[^>]*?w:value="([^"]*)"[^>]*?w:displayText="([^"]*)"/g;
    let lm2: RegExpExecArray | null;
    while ((lm2 = listItemRegex2.exec(prContent)) !== null) {
      if (!options.some((o) => o.value === lm2![1])) {
        options.push({ display: lm2![2], value: lm2![1] });
      }
    }
  }

  // Date format
  let dateFormat: string | null = null;
  if (type === "date") {
    const dfMatch = prContent.match(/<w:dateFormat[^>]*w:val="([^"]*)"/);
    dateFormat = dfMatch ? dfMatch[1] : null;
  }

  // Determine if inline: block SDTs have <w:p> in content, inline have <w:r> directly
  const contentInner = contentXml.replace(/<\/?w:sdtContent[^>]*>/g, "").trim();
  const isInline = !contentInner.startsWith("<w:p") && !contentInner.startsWith("<w:tc");

  return {
    tag,
    alias,
    type,
    currentValue,
    isChecked,
    isPlaceholder: showingPlcHdr,
    placeholderText: showingPlcHdr ? currentValue : "",
    options,
    dateFormat,
    isInline,
    xmlStart,
    xmlEnd: xmlStart + sdtXml.length,
    sdtPrXml: prXml,
    sdtContentXml: contentXml,
    fullXml: sdtXml,
    xmlPath,
  };
}

function extractAttr(xml: string, tagName: string, attrName: string): string | null {
  // Handle both w:val and just val with namespace prefix
  const regex = new RegExp(`<${tagName}[^>]*?${attrName}="([^"]*)"`, "i");
  const match = xml.match(regex);
  return match ? match[1] : null;
}

function extractTextContent(xml: string): string {
  const texts: string[] = [];
  const tRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let m;
  while ((m = tRegex.exec(xml)) !== null) {
    texts.push(xmlDecode(m[1]));
  }
  return texts.join("");
}

// ── Manipulation ───────────────────────────────────────────────────────────────

/**
 * Build a display name for a field, using tag, alias, or placeholder text.
 */
export function fieldDisplayName(field: SdtField): string {
  if (field.tag) return field.tag;
  if (field.alias) return field.alias;
  if (field.placeholderText) return `[${field.placeholderText.substring(0, 40)}]`;
  return `[unnamed at ${field.xmlStart}]`;
}

/**
 * Match a field by name — checks tag, alias, and placeholder text.
 */
export function matchField(field: SdtField, name: string): boolean {
  const lower = name.toLowerCase();
  if (field.tag?.toLowerCase() === lower) return true;
  if (field.alias?.toLowerCase() === lower) return true;
  if (field.placeholderText?.toLowerCase() === lower) return true;
  // Also try partial match on placeholder
  if (field.placeholderText?.toLowerCase().includes(lower)) return true;
  return false;
}

/**
 * Fill an SDT text field with a new value.
 * Returns the modified full document XML.
 */
export function fillSdtField(xml: string, field: SdtField, value: string): string {
  // Extract run properties from existing content, strip placeholder styling
  const rPr = extractCleanRpr(field);

  // Extract paragraph properties from existing content (for block SDTs)
  const pPr = extractParagraphProperties(field);

  // Build new sdtContent
  let newContent: string;
  if (field.isInline) {
    newContent = `<w:sdtContent>${makeRun(value, rPr)}</w:sdtContent>`;
  } else {
    // Handle multi-line values: each line becomes a paragraph
    const lines = value.split("\n");
    const paragraphs = lines.map((line) => `<w:p>${pPr}${makeRun(line, rPr)}</w:p>`).join("");
    newContent = `<w:sdtContent>${paragraphs}</w:sdtContent>`;
  }

  // Remove showingPlcHdr from sdtPr
  const newPr = field.sdtPrXml.replace(/<w:showingPlcHdr[^>]*(?:\/>|><\/w:showingPlcHdr>)/g, "");

  // Build new SDT XML
  const newSdt = field.fullXml.replace(field.sdtPrXml, newPr).replace(field.sdtContentXml, newContent);

  // Replace using position (safer than string matching)
  return xml.substring(0, field.xmlStart) + newSdt + xml.substring(field.xmlEnd);
}

/**
 * Check or uncheck a checkbox SDT.
 * Returns the modified full document XML.
 */
export function checkSdtField(xml: string, field: SdtField, checked: boolean): string {
  let newSdt = field.fullXml;

  // Update w14:checked value
  newSdt = newSdt.replace(
    /(<w14:checked[^>]*?val=")(\d+)(")/,
    `$1${checked ? "1" : "0"}$3`
  );

  // Update display character in sdtContent
  const checkChar = checked ? "\u2612" : "\u2610"; // ☒ or ☐
  newSdt = newSdt.replace(
    /(<w:sdtContent>[\s\S]*?<w:t[^>]*>)[^<]*(<\/w:t>[\s\S]*?<\/w:sdtContent>)/,
    `$1${checkChar}$2`
  );

  // Remove showingPlcHdr if present
  newSdt = newSdt.replace(/<w:showingPlcHdr[^>]*(?:\/>|><\/w:showingPlcHdr>)/g, "");

  return xml.substring(0, field.xmlStart) + newSdt + xml.substring(field.xmlEnd);
}

// ── Helpers ────────────────────────────────────────────────────────────────────

function makeRun(text: string, rPr: string): string {
  const needsSpace = text.startsWith(" ") || text.endsWith(" ") || text.includes("  ");
  const xmlSpace = needsSpace ? ' xml:space="preserve"' : "";
  return `<w:r>${rPr}<w:t${xmlSpace}>${xmlEncode(text)}</w:t></w:r>`;
}

function extractCleanRpr(field: SdtField): string {
  // Get rPr from the first run in sdtContent
  const rPrMatch = field.sdtContentXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
  if (!rPrMatch) return "";

  let rPr = rPrMatch[0];

  // Remove placeholder-specific styling:
  // - <w:rStyle val="PlaceholderText"/> (placeholder style)
  // - <w:color .../> (typically gray for placeholders)
  rPr = rPr.replace(/<w:rStyle[^>]*(?:\/>|>[^<]*<\/w:rStyle>)/g, "");
  rPr = rPr.replace(/<w:color[^>]*(?:\/>|>[^<]*<\/w:color>)/g, "");

  // If rPr is now effectively empty, return empty string
  const inner = rPr.replace(/<\/?w:rPr>/g, "").trim();
  if (!inner) return "";

  return rPr;
}

function extractParagraphProperties(field: SdtField): string {
  const pPrMatch = field.sdtContentXml.match(/<w:pPr>([\s\S]*?)<\/w:pPr>/);
  return pPrMatch ? pPrMatch[0] : "";
}
