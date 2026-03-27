import { readFile } from "fs/promises";
import { loadDocx, repackDocx, getTextXmlPaths } from "../utils/repack.js";
import { replaceTextInXml } from "../utils/xml.js";
import { detectFileType } from "../utils/detect.js";

export interface ReplaceOptions {
  /** Single old text */
  old?: string;
  /** Single new text */
  new?: string;
  /** Path to JSON file with replacement mappings */
  map?: string;
  /** Output path (default: in-place) */
  output?: string;
}

/**
 * Replace text in a .docx file. Unlike 'edit', replaces ALL occurrences
 * (like template/mail-merge filling). Supports batch replacements via JSON.
 */
export async function runReplace(filePath: string, options: ReplaceOptions): Promise<void> {
  const { type } = await detectFileType(filePath);
  if (type !== "docx") {
    throw new Error(`Only .docx files are supported. Detected format: ${type}`);
  }

  // Build replacement map
  const replacementMap = new Map<string, string>();

  if (options.map) {
    const jsonContent = await readFile(options.map, "utf-8");
    const data = JSON.parse(jsonContent);
    if (typeof data !== "object" || Array.isArray(data)) {
      throw new Error("JSON file must contain an object mapping old text → new text.");
    }
    for (const [oldText, newText] of Object.entries(data)) {
      replacementMap.set(oldText, String(newText));
    }
  } else if (options.old !== undefined && options.new !== undefined) {
    replacementMap.set(options.old, options.new);
  } else {
    throw new Error("Must provide --old/--new or --map <json-file>.");
  }

  if (replacementMap.size === 0) {
    throw new Error("No replacements specified.");
  }

  const { files } = await loadDocx(filePath);
  const textPaths = getTextXmlPaths(files);

  // Apply all replacements to all text-bearing XML files
  const modifications = new Map<string, string>();
  let totalReplacements = 0;
  const details: string[] = [];

  for (const xmlPath of textPaths) {
    let xml = files.get(xmlPath)!;
    let fileChanged = false;

    for (const [oldText, newText] of replacementMap) {
      const { xml: newXml, count } = replaceTextInXml(xml, oldText, newText, true);
      if (count > 0) {
        xml = newXml;
        fileChanged = true;
        totalReplacements += count;
        details.push(
          `  "${oldText.substring(0, 40)}${oldText.length > 40 ? "..." : ""}" → ` +
            `"${newText.substring(0, 40)}${newText.length > 40 ? "..." : ""}" ` +
            `(${count}× in ${xmlPath})`
        );
      }
    }

    if (fileChanged) {
      modifications.set(xmlPath, xml);
    }
  }

  if (totalReplacements === 0) {
    const searched = [...replacementMap.keys()]
      .map((k) => `"${k.substring(0, 50)}"`)
      .join(", ");
    throw new Error(
      `No matches found for: ${searched}\n` +
        `Hint: use 'deword read ${filePath}' to see the document content.`
    );
  }

  await repackDocx(filePath, modifications, options.output);

  const targetFile = options.output ?? filePath;
  console.error(`✓ ${totalReplacements} replacement(s) made → ${targetFile}`);
  for (const detail of details) {
    console.error(detail);
  }
}
