/**
 * Convert Word-exported HTML to Markdown.
 *
 * This is a lightweight, dependency-free converter that handles
 * the subset of HTML that Word / MHTML .doc files produce:
 *   headings, paragraphs, bold, italic, strikethrough,
 *   tables, lists, links, images, code/pre blocks.
 */

import type { ExtractedContent, ImageRef, TableData, DocMetadata } from "./content.js";

export function htmlToExtractedContent(html: string): ExtractedContent {
  const images: ImageRef[] = [];
  const tables: TableData[] = [];

  // Extract <title> for metadata
  const titleMatch = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  const title = titleMatch ? decodeEntities(titleMatch[1]).trim() : null;

  // Extract the <body> content
  const bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  const body = bodyMatch ? bodyMatch[1] : html;

  const markdown = convertNode(body, images, tables);

  // Clean up excessive blank lines
  const cleaned = markdown
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  const metadata: DocMetadata = {
    title,
    author: null,
    created: null,
    modified: null,
    paragraphCount: cleaned.split(/\n\n+/).length,
    imageCount: images.length,
    tableCount: tables.length,
  };

  return { markdown: cleaned, images, tables, metadata };
}

function convertNode(html: string, images: ImageRef[], tables: TableData[]): string {
  let result = html;

  // Process tables first (before we strip other tags)
  result = result.replace(/<table[^>]*>([\s\S]*?)<\/table>/gi, (_, tableHtml) => {
    const table = parseTable(tableHtml);
    tables.push(table);
    return "\n\n" + formatTableMd(table) + "\n\n";
  });

  // Headings
  result = result.replace(/<h([1-6])[^>]*>([\s\S]*?)<\/h\1>/gi, (_, level, content) => {
    const prefix = "#".repeat(parseInt(level, 10));
    return `\n\n${prefix} ${inlineToMd(content).trim()}\n\n`;
  });

  // Pre/code blocks
  result = result.replace(/<pre[^>]*>([\s\S]*?)<\/pre>/gi, (_, content) => {
    // Strip inner <code> tags if present
    let code = content.replace(/<\/?code[^>]*>/gi, "");
    code = decodeEntities(code).trim();
    return `\n\n\`\`\`\n${code}\n\`\`\`\n\n`;
  });

  // Lists — <ul> / <ol>
  result = result.replace(/<ul[^>]*>([\s\S]*?)<\/ul>/gi, (_, inner) => {
    return "\n\n" + parseList(inner, false) + "\n\n";
  });
  result = result.replace(/<ol[^>]*>([\s\S]*?)<\/ol>/gi, (_, inner) => {
    return "\n\n" + parseList(inner, true) + "\n\n";
  });

  // Paragraphs
  result = result.replace(/<p[^>]*>([\s\S]*?)<\/p>/gi, (_, content) => {
    const md = inlineToMd(content).trim();
    if (!md) return "\n";
    return `\n\n${md}\n\n`;
  });

  // Line breaks
  result = result.replace(/<br\s*\/?>/gi, "\n");

  // Horizontal rules
  result = result.replace(/<hr\s*\/?>/gi, "\n\n---\n\n");

  // Images
  result = result.replace(/<img[^>]*>/gi, (tag) => {
    const src = attr(tag, "src") ?? "";
    const alt = attr(tag, "alt") ?? "image";
    images.push({ id: src, path: src, description: alt, size: 0 });
    return `![${alt}](${src})`;
  });

  // Strip any remaining HTML tags
  result = result.replace(/<[^>]+>/g, "");

  // Decode HTML entities
  result = decodeEntities(result);

  return result;
}

/** Convert inline HTML (bold, italic, links, code, etc.) to Markdown */
function inlineToMd(html: string): string {
  let s = html;

  // Bold
  s = s.replace(/<(strong|b)\b[^>]*>([\s\S]*?)<\/\1>/gi, (_, _tag, content) => `**${content}**`);
  // Italic
  s = s.replace(/<(em|i)\b[^>]*>([\s\S]*?)<\/\1>/gi, (_, _tag, content) => `*${content}*`);
  // Strikethrough
  s = s.replace(/<(del|s|strike)\b[^>]*>([\s\S]*?)<\/\1>/gi, (_, _tag, content) => `~~${content}~~`);
  // Inline code
  s = s.replace(/<code[^>]*>([\s\S]*?)<\/code>/gi, (_, content) => `\`${decodeEntities(content)}\``);
  // Superscript
  s = s.replace(/<sup[^>]*>([\s\S]*?)<\/sup>/gi, (_, content) => `^${content}^`);
  // Subscript
  s = s.replace(/<sub[^>]*>([\s\S]*?)<\/sub>/gi, (_, content) => `~${content}~`);

  // Links
  s = s.replace(/<a\s[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, (_, href, text) => {
    const linkText = text.replace(/<[^>]+>/g, "").trim();
    return `[${linkText}](${href})`;
  });

  // Images
  s = s.replace(/<img[^>]*>/gi, (tag) => {
    const src = attr(tag, "src") ?? "";
    const alt = attr(tag, "alt") ?? "image";
    return `![${alt}](${src})`;
  });

  // Line breaks
  s = s.replace(/<br\s*\/?>/gi, "\n");

  // Strip remaining tags
  s = s.replace(/<[^>]+>/g, "");

  // Decode entities
  s = decodeEntities(s);

  return s;
}

function parseList(html: string, ordered: boolean): string {
  const items: string[] = [];
  const liRegex = /<li[^>]*>([\s\S]*?)<\/li>/gi;
  let match;
  let idx = 1;
  while ((match = liRegex.exec(html)) !== null) {
    const content = inlineToMd(match[1]).trim();
    const prefix = ordered ? `${idx}. ` : "- ";
    items.push(prefix + content);
    idx++;
  }
  return items.join("\n");
}

function parseTable(html: string): TableData {
  const rows: string[][] = [];
  const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let trMatch;

  while ((trMatch = trRegex.exec(html)) !== null) {
    const cells: string[] = [];
    const cellRegex = /<(td|th)[^>]*>([\s\S]*?)<\/\1>/gi;
    let cellMatch;
    while ((cellMatch = cellRegex.exec(trMatch[1])) !== null) {
      cells.push(inlineToMd(cellMatch[2]).trim());
    }
    if (cells.length > 0) rows.push(cells);
  }

  return { index: 0, rows };
}

function formatTableMd(table: TableData): string {
  if (table.rows.length === 0) return "";

  const colCount = Math.max(...table.rows.map((r) => r.length));
  const colWidths = new Array(colCount).fill(3);
  for (const row of table.rows) {
    for (let i = 0; i < row.length; i++) {
      colWidths[i] = Math.max(colWidths[i], (row[i] ?? "").length);
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
    if (rowIdx === 0) {
      lines.push("| " + colWidths.map((w) => "-".repeat(w)).join(" | ") + " |");
    }
  }

  return lines.join("\n");
}

/** Extract an attribute value from a raw HTML tag string */
function attr(tag: string, name: string): string | null {
  const re = new RegExp(`${name}\\s*=\\s*"([^"]*)"`, "i");
  const m = tag.match(re);
  return m ? m[1] : null;
}

/** Decode common HTML entities */
function decodeEntities(s: string): string {
  return s
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&nbsp;/g, " ")
    .replace(/&mdash;/g, "—")
    .replace(/&ndash;/g, "–")
    .replace(/&hellip;/g, "…")
    .replace(/&lsquo;/g, "'")
    .replace(/&rsquo;/g, "'")
    .replace(/&ldquo;/g, "\u201C")
    .replace(/&rdquo;/g, "\u201D")
    .replace(/&trade;/g, "™")
    .replace(/&copy;/g, "©")
    .replace(/&reg;/g, "®")
    .replace(/&deg;/g, "°")
    .replace(/&plusmn;/g, "±")
    .replace(/&times;/g, "×")
    .replace(/&divide;/g, "÷");
}
