import { readFile } from "fs/promises";

export type FileType = "docx" | "mhtml" | "unknown";

/**
 * Detect the actual file type by inspecting magic bytes / content,
 * regardless of the file extension.
 *
 *  - ZIP (docx):  starts with PK\x03\x04
 *  - MHTML:       starts with "MIME-Version:"
 *  - OLE2 (.doc): starts with \xD0\xCF\x11\xE0  (not supported yet)
 */
export async function detectFileType(filePath: string): Promise<{ type: FileType; data: Buffer }> {
  const data = await readFile(filePath);

  // ZIP magic bytes → .docx
  if (data[0] === 0x50 && data[1] === 0x4b && data[2] === 0x03 && data[3] === 0x04) {
    return { type: "docx", data };
  }

  // MIME header → MHTML (Word "Single File Web Page")
  const head = data.subarray(0, 256).toString("utf-8");
  if (head.startsWith("MIME-Version:") || head.includes("Content-Type: multipart/related")) {
    return { type: "mhtml", data };
  }

  // OLE2 compound document (legacy .doc) — not yet supported
  if (data[0] === 0xd0 && data[1] === 0xcf && data[2] === 0x11 && data[3] === 0xe0) {
    throw new Error(
      "This is a legacy binary .doc file (OLE2 format). Only .docx (ZIP/XML) and MHTML .doc files are supported."
    );
  }

  return { type: "unknown", data };
}
