#!/usr/bin/env node

import { Command } from "commander";
import { runUnpack } from "./commands/unpack.js";
import { runRead } from "./commands/read.js";

const program = new Command();

program
  .name("deword")
  .description("🪱 De-Words your documents for AI agents")
  .version("0.1.0");

program
  .command("read")
  .description("Read a .docx or .doc (MHTML) file and output contents to stdout")
  .argument("<file>", "Path to .docx or .doc file")
  .option("-f, --format <format>", "Output format: markdown, json, summary", "markdown")
  .option("-i, --images <dir>", "Override image extraction directory (default: auto temp dir)")
  .action(async (file: string, opts: { format: string; images?: string }) => {
    try {
      await runRead(file, {
        format: opts.format as "markdown" | "json" | "summary",
        imageDir: opts.images,
      });
    } catch (err) {
      console.error(`Error: ${err instanceof Error ? err.message : err}`);
      process.exit(1);
    }
  });

program
  .command("unpack")
  .description("Unpack a .docx to a working directory with agent-friendly files")
  .argument("<file>", "Path to .docx file")
  .option("-o, --output <dir>", "Output directory", "./docx-work")
  .option("-f, --format <format>", "Content format: markdown, json, both", "both")
  .option("--no-xml", "Skip writing prettified XML files")
  .option("--no-media", "Skip extracting media files")
  .action(async (file: string, opts: { output: string; format: string; xml: boolean; media: boolean }) => {
    try {
      await runUnpack(file, {
        output: opts.output,
        format: opts.format as "markdown" | "json" | "both",
        includeXml: opts.xml,
        includeMedia: opts.media,
      });
    } catch (err) {
      console.error(`Error: ${err instanceof Error ? err.message : err}`);
      process.exit(1);
    }
  });

program.parse();
