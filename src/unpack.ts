import JSZip from "jszip";
import { readFile } from "fs/promises";
import { XMLParser } from "fast-xml-parser";

export interface UnpackedDocx {
  /** Raw XML files keyed by path inside the zip */
  files: Map<string, string>;
  /** Binary files (images, etc.) keyed by path */
  media: Map<string, Buffer>;
  /** The main document.xml parsed */
  documentXml: string;
  /** Styles XML if present */
  stylesXml: string | null;
  /** Numbering XML (lists) if present */
  numberingXml: string | null;
  /** Relationships */
  relationships: Map<string, Relationship>;
  /** Content types */
  contentTypes: string;
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
}

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: false,
  trimValues: false,
});

export async function unpack(filePath: string): Promise<UnpackedDocx> {
  const data = await readFile(filePath);
  const zip = await JSZip.loadAsync(data);

  const files = new Map<string, string>();
  const media = new Map<string, Buffer>();
  let documentXml = "";
  let stylesXml: string | null = null;
  let numberingXml: string | null = null;
  let contentTypes = "";

  for (const [path, zipEntry] of Object.entries(zip.files)) {
    if (zipEntry.dir) continue;

    if (path.startsWith("word/media/") || path.match(/\.(png|jpg|jpeg|gif|bmp|tiff|emf|wmf)$/i)) {
      const buf = await zipEntry.async("nodebuffer");
      media.set(path, buf);
    } else {
      const text = await zipEntry.async("text");
      files.set(path, text);

      if (path === "word/document.xml") documentXml = text;
      if (path === "word/styles.xml") stylesXml = text;
      if (path === "word/numbering.xml") numberingXml = text;
      if (path === "[Content_Types].xml") contentTypes = text;
    }
  }

  // Parse relationships
  const relationships = new Map<string, Relationship>();
  const relsXml = files.get("word/_rels/document.xml.rels");
  if (relsXml) {
    const parsed = xmlParser.parse(relsXml);
    const rels = parsed?.Relationships?.Relationship;
    if (rels) {
      const relArray = Array.isArray(rels) ? rels : [rels];
      for (const rel of relArray) {
        relationships.set(rel["@_Id"], {
          id: rel["@_Id"],
          type: rel["@_Type"],
          target: rel["@_Target"],
        });
      }
    }
  }

  return {
    files,
    media,
    documentXml,
    stylesXml,
    numberingXml,
    relationships,
    contentTypes,
  };
}
