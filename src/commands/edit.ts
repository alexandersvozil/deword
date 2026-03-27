import { loadDocx, repackDocx, getTextXmlPaths } from "../utils/repack.js";
import { replaceTextInXml, extractAllText } from "../utils/xml.js";
import { detectFileType } from "../utils/detect.js";

export interface EditOptions {
  old: string;
  new: string;
  output?: string;
}

/**
 * Replace a unique text occurrence in a .docx file (in-place by default).
 * Like Pi's edit tool: oldText must match exactly once across the entire document.
 */
export async function runEdit(filePath: string, options: EditOptions): Promise<void> {
  // Verify it's a docx
  const { type } = await detectFileType(filePath);
  if (type !== "docx") {
    throw new Error(
      `Only .docx files can be edited. Detected format: ${type}.\n` +
        `Hint: .doc files must be converted to .docx first.`
    );
  }

  const { files } = await loadDocx(filePath);
  const textPaths = getTextXmlPaths(files);

  if (textPaths.length === 0) {
    throw new Error("No text-bearing XML files found in document.");
  }

  // Count total occurrences across all text-bearing XML files (use replaceAll to get full count)
  let totalCount = 0;
  const matches: Array<{ path: string; count: number }> = [];

  for (const xmlPath of textPaths) {
    const xml = files.get(xmlPath)!;
    const { count } = replaceTextInXml(xml, options.old, options.new, true);
    if (count > 0) {
      totalCount += count;
      matches.push({ path: xmlPath, count });
    }
  }

  if (totalCount === 0) {
    // Check if text exists anywhere (maybe in non-text XML files or across paragraphs)
    let foundAnywhere = false;
    for (const xmlPath of textPaths) {
      const xml = files.get(xmlPath)!;
      const allText = extractAllText(xml);
      if (allText.includes(options.old)) {
        foundAnywhere = true;
        break;
      }
    }

    if (foundAnywhere) {
      throw new Error(
        `Text found but spans different formatting runs that couldn't be merged.\n` +
          `Try replacing a smaller segment, or use 'deword xml' for direct XML editing.`
      );
    }

    throw new Error(
      `Text not found in document: "${options.old.substring(0, 80)}${options.old.length > 80 ? "..." : ""}"\n` +
        `Hint: use 'deword read ${filePath}' to see the document content.`
    );
  }

  if (totalCount > 1) {
    const details = matches.map((m) => `  ${m.path}: ${m.count} occurrence(s)`).join("\n");
    throw new Error(
      `Text found ${totalCount} times. Must be unique (like Pi's edit tool).\n` +
        `Add more surrounding context to make it unique.\n` +
        `Occurrences:\n${details}`
    );
  }

  // Exactly one match — do the replacement (replaceAll=false for single match)
  const replacements = new Map<string, string>();
  for (const { path } of matches) {
    const xml = files.get(path)!;
    const { xml: newXml } = replaceTextInXml(xml, options.old, options.new, false);
    replacements.set(path, newXml);
  }

  await repackDocx(filePath, replacements, options.output);

  const targetFile = options.output ?? filePath;
  console.error(`✓ Replaced 1 occurrence in ${matches[0].path} → ${targetFile}`);
}
