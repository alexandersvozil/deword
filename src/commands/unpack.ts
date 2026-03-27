import { mkdir, writeFile } from "fs/promises";
import { join } from "path";
import { unpack } from "../unpack.js";
import { extractContent } from "../extractors/content.js";
import { mergeAdjacentRuns, normalizeQuotes, prettyPrintXml } from "../utils/xml.js";
import { formatMarkdown, formatJson } from "../utils/format.js";

export interface UnpackOptions {
  output: string;
  format: "markdown" | "json" | "both";
  includeXml: boolean;
  includeMedia: boolean;
}

export async function runUnpack(filePath: string, options: UnpackOptions): Promise<void> {
  const outDir = options.output;
  await mkdir(outDir, { recursive: true });

  console.log(`Unpacking: ${filePath}`);
  const doc = await unpack(filePath);

  // 1. Write prettified + run-merged XML files
  if (options.includeXml) {
    const xmlDir = join(outDir, "xml");
    await mkdir(xmlDir, { recursive: true });

    for (const [path, content] of doc.files) {
      const outPath = join(xmlDir, path);
      await mkdir(join(outPath, ".."), { recursive: true });

      let processed = content;
      if (path.endsWith(".xml")) {
        processed = prettyPrintXml(processed);
        if (path === "word/document.xml") {
          processed = mergeAdjacentRuns(processed);
          processed = normalizeQuotes(processed);
        }
      }
      await writeFile(outPath, processed, "utf-8");
    }
    console.log(`  XML files written to: ${xmlDir}`);
  }

  // 2. Extract media files
  if (options.includeMedia && doc.media.size > 0) {
    const mediaDir = join(outDir, "media");
    await mkdir(mediaDir, { recursive: true });

    for (const [path, buf] of doc.media) {
      const filename = path.split("/").pop()!;
      await writeFile(join(mediaDir, filename), buf);
    }
    console.log(`  Media files written to: ${join(outDir, "media")} (${doc.media.size} files)`);
  }

  // 3. Run-merge the document XML before extraction
  const mergedXml = mergeAdjacentRuns(normalizeQuotes(doc.documentXml));
  const docWithMerged = { ...doc, documentXml: mergedXml };
  const content = extractContent(docWithMerged);

  // 4. Write agent-friendly output
  if (options.format === "markdown" || options.format === "both") {
    const md = formatMarkdown(content);
    await writeFile(join(outDir, "CONTENT.md"), md, "utf-8");
    console.log(`  Markdown written to: ${join(outDir, "CONTENT.md")}`);
  }

  if (options.format === "json" || options.format === "both") {
    const json = formatJson(content);
    await writeFile(join(outDir, "content.json"), json, "utf-8");
    console.log(`  JSON written to: ${join(outDir, "content.json")}`);
  }

  // 5. Write a manifest for the agent
  const manifest = {
    source: filePath,
    extractedAt: new Date().toISOString(),
    files: {
      markdown: options.format !== "json" ? "CONTENT.md" : null,
      json: options.format !== "markdown" ? "content.json" : null,
      xml: options.includeXml ? "xml/" : null,
      media: options.includeMedia && doc.media.size > 0 ? "media/" : null,
    },
    stats: {
      paragraphs: content.metadata.paragraphCount,
      tables: content.metadata.tableCount,
      images: content.metadata.imageCount,
      xmlFiles: doc.files.size,
      mediaFiles: doc.media.size,
    },
  };
  await writeFile(join(outDir, "manifest.json"), JSON.stringify(manifest, null, 2), "utf-8");
  console.log(`  Manifest written to: ${join(outDir, "manifest.json")}`);

  console.log("\nDone. Agent can now read CONTENT.md or content.json for document contents.");
}
