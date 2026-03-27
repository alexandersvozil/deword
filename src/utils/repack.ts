import JSZip from "jszip";
import { readFile, writeFile } from "fs/promises";

/**
 * Load a .docx file and return the JSZip instance plus all XML files as text.
 */
export async function loadDocx(filePath: string): Promise<{
  zip: JSZip;
  files: Map<string, string>;
}> {
  const data = await readFile(filePath);
  const zip = await JSZip.loadAsync(data);

  const files = new Map<string, string>();
  for (const [path, entry] of Object.entries(zip.files)) {
    if (!entry.dir) {
      // Read XML and rels files as text
      if (path.endsWith(".xml") || path.endsWith(".rels")) {
        files.set(path, await entry.async("text"));
      }
    }
  }

  return { zip, files };
}

/**
 * Repack a .docx file, replacing specified files while preserving everything else.
 * Edits in-place by default (same as Pi's edit tool).
 */
export async function repackDocx(
  filePath: string,
  replacements: Map<string, string>,
  outputPath?: string
): Promise<void> {
  const data = await readFile(filePath);
  const zip = await JSZip.loadAsync(data);

  for (const [path, content] of replacements) {
    zip.file(path, content);
  }

  const output = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  await writeFile(outputPath ?? filePath, output);
}

/**
 * Get paths of XML files that contain document text
 * (body, headers, footers, footnotes, endnotes).
 */
export function getTextXmlPaths(files: Map<string, string>): string[] {
  const paths: string[] = [];
  for (const path of files.keys()) {
    if (
      path === "word/document.xml" ||
      /^word\/header\d+\.xml$/.test(path) ||
      /^word\/footer\d+\.xml$/.test(path) ||
      path === "word/footnotes.xml" ||
      path === "word/endnotes.xml"
    ) {
      paths.push(path);
    }
  }
  return paths;
}
