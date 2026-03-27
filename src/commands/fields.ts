import { loadDocx, getTextXmlPaths } from "../utils/repack.js";
import { findSdtFields, fieldDisplayName, type SdtField } from "../utils/sdt.js";
import { detectFileType } from "../utils/detect.js";

export interface FieldsOptions {
  format: "text" | "json";
}

/**
 * List all SDT content controls (form fields) in a .docx file.
 */
export async function runFields(filePath: string, options: FieldsOptions): Promise<void> {
  const { type } = await detectFileType(filePath);
  if (type !== "docx") {
    throw new Error(`Only .docx files are supported. Detected format: ${type}`);
  }

  const { files } = await loadDocx(filePath);
  const textPaths = getTextXmlPaths(files);

  const allFields: SdtField[] = [];
  for (const xmlPath of textPaths) {
    const xml = files.get(xmlPath)!;
    const fields = findSdtFields(xml, xmlPath);
    allFields.push(...fields);
  }

  if (allFields.length === 0) {
    if (options.format === "json") {
      process.stdout.write(JSON.stringify({ fields: [] }, null, 2) + "\n");
    } else {
      console.log("No content controls (SDT fields) found in document.");
    }
    return;
  }

  if (options.format === "json") {
    const jsonFields = allFields.map((f, i) => ({
      index: i + 1,
      tag: f.tag,
      alias: f.alias,
      type: f.type,
      currentValue: f.currentValue,
      isChecked: f.type === "checkbox" ? f.isChecked : undefined,
      isPlaceholder: f.isPlaceholder,
      placeholderText: f.isPlaceholder ? f.placeholderText : undefined,
      options: f.options.length > 0 ? f.options : undefined,
      dateFormat: f.dateFormat,
      isInline: f.isInline,
      xmlFile: f.xmlPath,
    }));
    process.stdout.write(JSON.stringify({ fields: jsonFields }, null, 2) + "\n");
    return;
  }

  // Text format
  console.log(`Fields in ${filePath}:\n`);
  for (let i = 0; i < allFields.length; i++) {
    const f = allFields[i];
    const num = String(i + 1).padStart(3);
    const name = fieldDisplayName(f);
    const typeStr = f.type;
    const context = f.isInline ? "inline" : "block";

    let details = `${num}. "${name}" (${typeStr}, ${context})`;

    if (f.type === "checkbox") {
      details += ` — ${f.isChecked ? "☑ checked" : "☐ unchecked"}`;
    } else if (f.isPlaceholder) {
      details += ` — placeholder: "${f.placeholderText.substring(0, 50)}"`;
    } else if (f.currentValue) {
      details += ` — value: "${f.currentValue.substring(0, 50)}"`;
    }

    if (f.options.length > 0) {
      const opts = f.options.map((o) => o.display).join(", ");
      details += `\n     options: ${opts}`;
    }

    if (f.xmlPath !== "word/document.xml") {
      details += `\n     in: ${f.xmlPath}`;
    }

    console.log(details);
  }

  console.log(`\n${allFields.length} field(s) found.`);
}
