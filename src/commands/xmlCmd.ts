import { readFile } from "fs/promises";
import { loadDocx, repackDocx } from "../utils/repack.js";
import { prettyPrintXml } from "../utils/xml.js";
import { detectFileType } from "../utils/detect.js";
import { XMLValidator } from "fast-xml-parser";

export interface XmlOptions {
  /** Specific file path inside the ZIP to show (default: word/document.xml) */
  path?: string;
  /** List all files in the ZIP */
  list?: boolean;
  /** Replace a file inside the ZIP */
  set?: string;
  /** Input file for --set (if not provided, reads from stdin) */
  input?: string;
  /** Output path (default: in-place) */
  output?: string;
}

/**
 * Safe XML escape hatch: extract, list, or replace XML files in a .docx.
 * Handles the ZIP plumbing so agents can focus on the XML content.
 */
export async function runXml(filePath: string, options: XmlOptions): Promise<void> {
  const { type } = await detectFileType(filePath);
  if (type !== "docx") {
    throw new Error(`Only .docx files are supported. Detected format: ${type}`);
  }

  // Mode: list all files
  if (options.list) {
    await listFiles(filePath);
    return;
  }

  // Mode: replace a file
  if (options.set) {
    await setFile(filePath, options.set, options.input, options.output);
    return;
  }

  // Mode: show a file (default)
  await showFile(filePath, options.path ?? "word/document.xml");
}

async function listFiles(filePath: string): Promise<void> {
  const { zip } = await loadDocx(filePath);
  const entries: Array<{ path: string; size: number }> = [];

  for (const [path, entry] of Object.entries(zip.files)) {
    if (!entry.dir) {
      const data = await entry.async("nodebuffer");
      entries.push({ path, size: data.length });
    }
  }

  entries.sort((a, b) => a.path.localeCompare(b.path));

  console.log(`Files in ${filePath}:\n`);
  for (const entry of entries) {
    const sizeStr = entry.size > 1024
      ? `${(entry.size / 1024).toFixed(1)} KB`
      : `${entry.size} B`;
    console.log(`  ${entry.path.padEnd(45)} ${sizeStr.padStart(10)}`);
  }
  console.log(`\n${entries.length} file(s)`);
}

async function showFile(filePath: string, xmlPath: string): Promise<void> {
  const { files } = await loadDocx(filePath);
  const content = files.get(xmlPath);

  if (!content) {
    const available = [...files.keys()].sort().join("\n  ");
    throw new Error(
      `File "${xmlPath}" not found in archive.\n` +
        `Available files:\n  ${available}`
    );
  }

  // Pretty-print XML files
  if (xmlPath.endsWith(".xml") || xmlPath.endsWith(".rels")) {
    process.stdout.write(prettyPrintXml(content));
  } else {
    process.stdout.write(content);
  }
}

async function setFile(
  filePath: string,
  targetPath: string,
  inputFile: string | undefined,
  outputPath: string | undefined
): Promise<void> {
  // Read new content
  let newContent: string;
  if (inputFile) {
    newContent = await readFile(inputFile, "utf-8");
  } else {
    // Read from stdin
    const chunks: Buffer[] = [];
    for await (const chunk of process.stdin) {
      chunks.push(chunk);
    }
    newContent = Buffer.concat(chunks).toString("utf-8");
  }

  if (!newContent.trim()) {
    throw new Error("Input is empty. Refusing to write an empty file.");
  }

  // Validate XML
  if (targetPath.endsWith(".xml") || targetPath.endsWith(".rels")) {
    const validation = XMLValidator.validate(newContent);
    if (validation !== true) {
      const err = validation as { err: { msg: string; line: number; col: number } };
      throw new Error(
        `Invalid XML: ${err.err.msg} at line ${err.err.line}, col ${err.err.col}\n` +
          `Fix the XML and try again.`
      );
    }
  }

  // Verify the target path exists in the archive
  const { files } = await loadDocx(filePath);
  if (!files.has(targetPath)) {
    const available = [...files.keys()]
      .filter((p) => p.endsWith(".xml") || p.endsWith(".rels"))
      .sort()
      .join("\n  ");
    console.error(
      `Warning: "${targetPath}" does not exist in archive. It will be added.\n` +
        `Existing XML files:\n  ${available}`
    );
  }

  const replacements = new Map<string, string>();
  replacements.set(targetPath, newContent);

  await repackDocx(filePath, replacements, outputPath);

  const target = outputPath ?? filePath;
  console.error(`✓ Set ${targetPath} → ${target}`);
}
