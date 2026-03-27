import { readFile } from "fs/promises";
import { loadDocx, repackDocx, getTextXmlPaths } from "../utils/repack.js";
import {
  findSdtFields,
  matchField,
  fieldDisplayName,
  fillSdtField,
  checkSdtField,
  type SdtField,
} from "../utils/sdt.js";
import { detectFileType } from "../utils/detect.js";

export interface FillOptions {
  /** Single field name to fill */
  field?: string;
  /** Value for single field */
  value?: string;
  /** Check a checkbox */
  check?: boolean;
  /** Uncheck a checkbox */
  uncheck?: boolean;
  /** Path to JSON file with field→value mappings */
  json?: string;
  /** Output path (default: in-place) */
  output?: string;
}

interface FillEntry {
  name: string;
  value?: string;
  checked?: boolean;
}

/**
 * Fill form fields (SDT content controls) in a .docx file.
 */
export async function runFill(filePath: string, options: FillOptions): Promise<void> {
  const { type } = await detectFileType(filePath);
  if (type !== "docx") {
    throw new Error(`Only .docx files are supported. Detected format: ${type}`);
  }

  // Build the list of fields to fill
  const entries: FillEntry[] = [];

  if (options.json) {
    const jsonContent = await readFile(options.json, "utf-8");
    const data = JSON.parse(jsonContent);
    for (const [name, val] of Object.entries(data)) {
      if (typeof val === "boolean") {
        entries.push({ name, checked: val });
      } else if (typeof val === "object" && val !== null && "checked" in val) {
        entries.push({ name, checked: (val as { checked: boolean }).checked });
      } else {
        entries.push({ name, value: String(val) });
      }
    }
  } else if (options.field) {
    if (options.check !== undefined || options.uncheck !== undefined) {
      entries.push({ name: options.field, checked: options.check === true });
    } else if (options.value !== undefined) {
      entries.push({ name: options.field, value: options.value });
    } else {
      throw new Error("Must provide --value, --check, or --uncheck with --field.");
    }
  } else {
    throw new Error("Must provide --field or --json.");
  }

  // Load document and find all fields
  const { files } = await loadDocx(filePath);
  const textPaths = getTextXmlPaths(files);

  const allFields: SdtField[] = [];
  for (const xmlPath of textPaths) {
    const xml = files.get(xmlPath)!;
    const fields = findSdtFields(xml, xmlPath);
    allFields.push(...fields);
  }

  if (allFields.length === 0) {
    throw new Error(
      "No content controls (SDT fields) found in document.\n" +
        "Hint: use 'deword fields' to inspect available fields."
    );
  }

  // Process each fill entry
  // Group by XML file, process entries from end-to-start to maintain positions
  const modifications = new Map<string, string>(); // xmlPath → modified xml

  // Get current XML state for each file
  for (const xmlPath of textPaths) {
    if (files.has(xmlPath)) {
      modifications.set(xmlPath, files.get(xmlPath)!);
    }
  }

  const results: string[] = [];

  for (const entry of entries) {
    // Find matching field
    const matching = allFields.filter((f) => matchField(f, entry.name));

    if (matching.length === 0) {
      const available = allFields.map((f) => fieldDisplayName(f)).join(", ");
      throw new Error(
        `Field "${entry.name}" not found.\n` + `Available fields: ${available}`
      );
    }

    if (matching.length > 1) {
      const names = matching.map((f) => `"${fieldDisplayName(f)}" in ${f.xmlPath}`).join(", ");
      throw new Error(
        `Field "${entry.name}" matched ${matching.length} fields: ${names}\n` +
          `Use a more specific name (tag or alias).`
      );
    }

    const field = matching[0];
    let currentXml = modifications.get(field.xmlPath)!;

    // Re-find the field in the (possibly modified) XML
    const currentFields = findSdtFields(currentXml, field.xmlPath);
    const currentField = currentFields.find((f) => matchField(f, entry.name));
    if (!currentField) {
      throw new Error(`Field "${entry.name}" could not be re-located after prior modifications.`);
    }

    if (entry.checked !== undefined) {
      if (currentField.type !== "checkbox") {
        throw new Error(
          `Field "${entry.name}" is a ${currentField.type}, not a checkbox. Use --value instead.`
        );
      }
      currentXml = checkSdtField(currentXml, currentField, entry.checked);
      results.push(
        `✓ ${entry.checked ? "Checked" : "Unchecked"} "${fieldDisplayName(currentField)}" in ${currentField.xmlPath}`
      );
    } else if (entry.value !== undefined) {
      currentXml = fillSdtField(currentXml, currentField, entry.value);
      results.push(
        `✓ Filled "${fieldDisplayName(currentField)}" with "${entry.value.substring(0, 50)}" in ${currentField.xmlPath}`
      );
    }

    modifications.set(field.xmlPath, currentXml);
  }

  // Write modifications
  const replacements = new Map<string, string>();
  for (const [xmlPath, xml] of modifications) {
    if (xml !== files.get(xmlPath)) {
      replacements.set(xmlPath, xml);
    }
  }

  if (replacements.size === 0) {
    console.error("No changes made.");
    return;
  }

  await repackDocx(filePath, replacements, options.output);

  const targetFile = options.output ?? filePath;
  for (const result of results) {
    console.error(result);
  }
  console.error(`→ ${targetFile}`);
}
