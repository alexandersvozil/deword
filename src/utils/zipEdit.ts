import JSZip from "jszip";
import { readFile, writeFile } from "fs/promises";

export async function repackDocxEntries(
  filePath: string,
  textEntries: Map<string, string>,
  binaryEntries: Map<string, Buffer> = new Map(),
  outputPath?: string
): Promise<void> {
  const data = await readFile(filePath);
  const zip = await JSZip.loadAsync(data);

  for (const [path, content] of textEntries) {
    zip.file(path, content);
  }
  for (const [path, content] of binaryEntries) {
    zip.file(path, content);
  }

  const output = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  await writeFile(outputPath ?? filePath, output);
}
