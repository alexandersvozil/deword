import type { ExtractedContent } from "../extractors/content.js";

/**
 * Format extraction results for agent consumption.
 * 
 * Three output modes:
 * - markdown: Human/agent-readable markdown (default)
 * - json: Structured JSON for programmatic agent use
 * - xml: Prettified raw XML for agents that need full fidelity
 */

export function formatMarkdown(content: ExtractedContent): string {
  const sections: string[] = [];

  // Metadata header
  sections.push("---");
  sections.push(`title: ${content.metadata.title ?? "Untitled"}`);
  if (content.metadata.author) sections.push(`author: ${content.metadata.author}`);
  if (content.metadata.created) sections.push(`created: ${content.metadata.created}`);
  if (content.metadata.modified) sections.push(`modified: ${content.metadata.modified}`);
  sections.push(`paragraphs: ${content.metadata.paragraphCount}`);
  sections.push(`images: ${content.metadata.imageCount}`);
  sections.push(`tables: ${content.metadata.tableCount}`);
  sections.push("---");
  sections.push("");

  // Main content
  sections.push(content.markdown);

  // Image inventory (if any)
  if (content.images.length > 0) {
    sections.push("");
    sections.push("## Image Inventory");
    sections.push("");
    for (const img of content.images) {
      const sizeKb = (img.size / 1024).toFixed(1);
      sections.push(
        `- **${img.path}** (${sizeKb} KB)${img.description ? `: ${img.description}` : ""}`
      );
    }
  }

  return sections.join("\n");
}

export function formatJson(content: ExtractedContent): string {
  return JSON.stringify(
    {
      metadata: content.metadata,
      content: content.markdown,
      images: content.images.map((img) => ({
        id: img.id,
        path: img.path,
        description: img.description,
        sizeBytes: img.size,
      })),
      tables: content.tables.map((tbl) => ({
        index: tbl.index,
        headers: tbl.rows[0] ?? [],
        rows: tbl.rows.slice(1),
      })),
    },
    null,
    2
  );
}

export function formatSummary(content: ExtractedContent): string {
  const lines: string[] = [];

  lines.push(`Document: ${content.metadata.title ?? "Untitled"}`);
  lines.push(`Author: ${content.metadata.author ?? "Unknown"}`);
  lines.push(`Paragraphs: ${content.metadata.paragraphCount}`);
  lines.push(`Tables: ${content.metadata.tableCount}`);
  lines.push(`Images: ${content.metadata.imageCount}`);
  lines.push("");

  // First ~500 chars as preview
  const preview = content.markdown.slice(0, 500);
  lines.push("Preview:");
  lines.push(preview);
  if (content.markdown.length > 500) {
    lines.push("...");
  }

  return lines.join("\n");
}
