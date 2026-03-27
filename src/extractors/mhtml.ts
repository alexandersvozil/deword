/**
 * Parse an MHTML (MIME multipart) buffer and extract the HTML body part.
 * Also extracts any inline images encoded as base64 MIME parts.
 */

export interface MhtmlPart {
  contentType: string;
  encoding: string;
  body: string | Buffer;
}

export interface ParsedMhtml {
  html: string;
  parts: MhtmlPart[];
}

export function parseMhtml(data: Buffer): ParsedMhtml {
  const text = data.toString("utf-8");

  // Extract the MIME boundary from the top-level Content-Type header
  const boundaryMatch = text.match(/boundary="([^"]+)"/);
  if (!boundaryMatch) {
    // No boundary — maybe it's just plain HTML with MIME headers
    const htmlStart = text.indexOf("<html");
    if (htmlStart !== -1) {
      return { html: text.slice(htmlStart), parts: [] };
    }
    throw new Error("Cannot parse MHTML: no boundary found and no HTML content detected.");
  }

  const boundary = boundaryMatch[1];
  const delimiter = `--${boundary}`;
  const segments = text.split(delimiter);

  const parts: MhtmlPart[] = [];
  let html = "";

  for (const segment of segments) {
    // Skip the preamble and closing delimiter
    if (!segment.trim() || segment.trim() === "--") continue;

    // Split headers from body (double newline)
    const headerEnd = segment.indexOf("\n\n") !== -1
      ? segment.indexOf("\n\n")
      : segment.indexOf("\r\n\r\n");

    if (headerEnd === -1) continue;

    const headerBlock = segment.slice(0, headerEnd);
    const body = segment.slice(headerEnd).replace(/^(\r?\n){1,2}/, "");

    // Parse headers
    const contentTypeMatch = headerBlock.match(/Content-Type:\s*([^\s;]+)/i);
    const encodingMatch = headerBlock.match(/Content-Transfer-Encoding:\s*(\S+)/i);

    const contentType = contentTypeMatch?.[1] ?? "application/octet-stream";
    const encoding = encodingMatch?.[1]?.toLowerCase() ?? "8bit";

    if (contentType.startsWith("text/html")) {
      // The HTML part — strip any trailing boundary artifacts
      html = body.replace(/\r?\n$/, "");
      parts.push({ contentType, encoding, body: html });
    } else if (contentType.startsWith("image/")) {
      // Base64-encoded image
      const imgBuffer = encoding === "base64"
        ? Buffer.from(body.replace(/\s/g, ""), "base64")
        : Buffer.from(body);
      parts.push({ contentType, encoding, body: imgBuffer });
    } else {
      parts.push({ contentType, encoding, body });
    }
  }

  if (!html) {
    // Fallback: try to find HTML anywhere in the raw text
    const htmlStart = text.indexOf("<html");
    const htmlEnd = text.lastIndexOf("</html>");
    if (htmlStart !== -1 && htmlEnd !== -1) {
      html = text.slice(htmlStart, htmlEnd + "</html>".length);
    } else {
      throw new Error("No HTML content found in MHTML file.");
    }
  }

  return { html, parts };
}
