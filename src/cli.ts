#!/usr/bin/env node

import { Command } from "commander";
import { runUnpack } from "./commands/unpack.js";
import { runRead } from "./commands/read.js";
import { runEdit } from "./commands/edit.js";
import { runFields } from "./commands/fields.js";
import { runFill } from "./commands/fill.js";
import { runXml } from "./commands/xmlCmd.js";
import { runReplace } from "./commands/replace.js";

const program = new Command();

program
  .name("deword")
  .description("🪱 De-Words your documents for AI agents")
  .version("0.2.0");

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

// ── Feature 1: edit ────────────────────────────────────────────────────────────

program
  .command("edit")
  .description("Replace a unique text occurrence in a .docx file (in-place)")
  .argument("<file>", "Path to .docx file")
  .requiredOption("--old <text>", "Exact text to find (must be unique)")
  .requiredOption("--new <text>", "Replacement text")
  .option("-o, --output <file>", "Write to a different file instead of editing in-place")
  .action(async (file: string, opts: { old: string; new: string; output?: string }) => {
    try {
      await runEdit(file, { old: opts.old, new: opts.new, output: opts.output });
    } catch (err) {
      console.error(`Error: ${err instanceof Error ? err.message : err}`);
      process.exit(1);
    }
  });

// ── Feature 2: fields ─────────────────────────────────────────────────────────

program
  .command("fields")
  .description("List all fillable content controls (SDT fields) in a .docx file")
  .argument("<file>", "Path to .docx file")
  .option("-f, --format <format>", "Output format: text, json", "text")
  .action(async (file: string, opts: { format: string }) => {
    try {
      await runFields(file, { format: opts.format as "text" | "json" });
    } catch (err) {
      console.error(`Error: ${err instanceof Error ? err.message : err}`);
      process.exit(1);
    }
  });

// ── Feature 3: fill ───────────────────────────────────────────────────────────

program
  .command("fill")
  .description("Fill form fields (SDT content controls) in a .docx file (in-place)")
  .argument("<file>", "Path to .docx file")
  .option("--field <name>", "Field name (tag, alias, or placeholder text)")
  .option("--value <text>", "Value to set")
  .option("--check", "Check a checkbox field")
  .option("--uncheck", "Uncheck a checkbox field")
  .option("--json <file>", "JSON file with field→value mappings")
  .option("-o, --output <file>", "Write to a different file instead of editing in-place")
  .action(
    async (
      file: string,
      opts: {
        field?: string;
        value?: string;
        check?: boolean;
        uncheck?: boolean;
        json?: string;
        output?: string;
      }
    ) => {
      try {
        await runFill(file, {
          field: opts.field,
          value: opts.value,
          check: opts.check,
          uncheck: opts.uncheck,
          json: opts.json,
          output: opts.output,
        });
      } catch (err) {
        console.error(`Error: ${err instanceof Error ? err.message : err}`);
        process.exit(1);
      }
    }
  );

// ── Feature 4: xml ────────────────────────────────────────────────────────────

program
  .command("xml")
  .description("Extract, list, or replace XML files inside a .docx (safe escape hatch)")
  .argument("<file>", "Path to .docx file")
  .option("-p, --path <path>", "XML file path inside the ZIP (default: word/document.xml)")
  .option("-l, --list", "List all files in the ZIP archive")
  .option("-s, --set <path>", "Replace a file inside the ZIP")
  .option("-i, --input <file>", "Input file for --set (reads stdin if omitted)")
  .option("-o, --output <file>", "Write to a different file instead of editing in-place")
  .action(
    async (
      file: string,
      opts: { path?: string; list?: boolean; set?: string; input?: string; output?: string }
    ) => {
      try {
        await runXml(file, {
          path: opts.path,
          list: opts.list,
          set: opts.set,
          input: opts.input,
          output: opts.output,
        });
      } catch (err) {
        console.error(`Error: ${err instanceof Error ? err.message : err}`);
        process.exit(1);
      }
    }
  );

// ── Feature 5: replace ────────────────────────────────────────────────────────

program
  .command("replace")
  .description("Replace all occurrences of text in a .docx file (batch/template mode, in-place)")
  .argument("<file>", "Path to .docx file")
  .option("--old <text>", "Text to find (replaces ALL occurrences)")
  .option("--new <text>", "Replacement text")
  .option("--map <file>", "JSON file mapping old text → new text")
  .option("-o, --output <file>", "Write to a different file instead of editing in-place")
  .action(
    async (file: string, opts: { old?: string; new?: string; map?: string; output?: string }) => {
      try {
        await runReplace(file, {
          old: opts.old,
          new: opts.new,
          map: opts.map,
          output: opts.output,
        });
      } catch (err) {
        console.error(`Error: ${err instanceof Error ? err.message : err}`);
        process.exit(1);
      }
    }
  );

program.parse();
