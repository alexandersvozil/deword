import { unpack } from "../unpack.js";
import { extractContent, type ExtractedContent } from "../extractors/content.js";
import { parseMhtml } from "../extractors/mhtml.js";
import { htmlToExtractedContent } from "../extractors/html.js";
import { mergeAdjacentRuns } from "../utils/xml.js";
import { detectFileType } from "../utils/detect.js";
import { formatMarkdown, formatJson, formatSummary } from "../utils/format.js";
import { mkdir, writeFile } from "fs/promises";
import { join, basename, resolve, parse as parsePath } from "path";
import { tmpdir } from "os";

export interface ReadOptions {
  format: "markdown" | "json" | "summary";
  imageDir?: string;
}

/**
 * Determine where to extract images.
 * Explicit --images flag wins, otherwise auto-create a temp dir
 * based on the source filename for stable, predictable paths.
 */
function resolveImageDir(filePath: string, explicit?: string): string {
  if (explicit) return resolve(explicit);
  const stem = parsePath(filePath).name.replace(/[^a-zA-Z0-9_-]/g, "_");
  return join(tmpdir(), "deword", stem, "images");
}

/**
 * Read a .docx or MHTML .doc file and output content to stdout.
 * Images are automatically extracted to a temp directory so agents
 * can view them — just like a human opening the document.
 */
export async function runRead(filePath: string, options: ReadOptions): Promise<void> {
  const { type, data } = await detectFileType(filePath);

  let content: ExtractedContent;
  let mediaMap: Map<string, Buffer> | undefined;

  switch (type) {
    case "docx": {
      const doc = await unpack(filePath);
      // Merge fragmented runs for cleaner text; skip normalizeQuotes
      // since it encodes Unicode → XML entities that fast-xml-parser won't decode back.
      const mergedXml = mergeAdjacentRuns(doc.documentXml);
      const docWithMerged = { ...doc, documentXml: mergedXml };
      content = extractContent(docWithMerged);
      mediaMap = doc.media;
      break;
    }
    case "mhtml": {
      const { html } = parseMhtml(data);
      content = htmlToExtractedContent(html);
      break;
    }
    default:
      throw new Error(
        `Unsupported file format. Expected a .docx (ZIP/XML) or MHTML .doc file.\n` +
        `Hint: use \`file ${filePath}\` to check the actual format.`
      );
  }

  // Always extract images when present — agents should see what humans see
  if (mediaMap && mediaMap.size > 0) {
    const imageDir = resolveImageDir(filePath, options.imageDir);
    await mkdir(imageDir, { recursive: true });

    for (const [zipPath, buf] of mediaMap) {
      const filename = basename(zipPath);
      const outPath = join(imageDir, filename);
      await writeFile(outPath, buf);
    }

    // Rewrite image references in markdown to point to extracted files
    for (const img of content.images) {
      const filename = basename(img.path);
      const newPath = join(imageDir, filename);
      content.markdown = content.markdown.replaceAll(
        `](${img.path})`,
        `](${newPath})`
      );
      img.path = newPath;
    }
  }

  switch (options.format) {
    case "markdown":
      process.stdout.write(formatMarkdown(content));
      break;
    case "json":
      process.stdout.write(formatJson(content));
      break;
    case "summary":
      process.stdout.write(formatSummary(content));
      break;
  }
}
