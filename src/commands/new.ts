import { access, copyFile, mkdir } from "fs/promises";
import { dirname, resolve } from "path";
import { fileURLToPath } from "url";
import { loadDocx, repackDocx } from "../utils/repack.js";
import { replaceTextInXml } from "../utils/xml.js";

export interface NewOptions {
  force?: boolean;
  text?: string;
}

function getTemplatePath(): string {
  return fileURLToPath(new URL("../../assets/new_document.docx", import.meta.url));
}

/**
 * Create a new .docx file from the bundled template.
 *
 * The template contains a single <empty> placeholder so agents can either:
 *  - replace it later with `deword edit` / `deword replace`, or
 *  - set initial content immediately with --text.
 */
export async function runNew(filePath: string, options: NewOptions): Promise<void> {
  const outputPath = resolve(filePath);
  const templatePath = getTemplatePath();

  try {
    await access(templatePath);
  } catch {
    throw new Error(`Bundled template not found: ${templatePath}`);
  }

  if (!options.force) {
    try {
      await access(outputPath);
      throw new Error(`File already exists: ${outputPath}\nUse --force to overwrite it.`);
    } catch (err) {
      if (err instanceof Error && err.message.startsWith("File already exists:")) {
        throw err;
      }
    }
  }

  await mkdir(dirname(outputPath), { recursive: true });
  await copyFile(templatePath, outputPath);

  if (options.text !== undefined) {
    const { files } = await loadDocx(outputPath);
    const xml = files.get("word/document.xml");
    if (!xml) {
      throw new Error("Template is missing word/document.xml.");
    }

    const { xml: newXml, count } = replaceTextInXml(xml, "<empty>", options.text, false);
    if (count !== 1) {
      throw new Error(`Expected exactly one <empty> placeholder in template, found ${count}.`);
    }

    await repackDocx(outputPath, new Map([["word/document.xml", newXml]]));
  }

  console.error(`✓ Created new document → ${outputPath}`);
  if (options.text === undefined) {
    console.error(`  Template placeholder: <empty>`);
  }
}
