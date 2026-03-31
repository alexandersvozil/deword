#!/usr/bin/env node

import { Command } from "commander";
import { runUnpack } from "./commands/unpack.js";
import { runRead } from "./commands/read.js";
import { runEdit } from "./commands/edit.js";
import { runFields } from "./commands/fields.js";
import { runFill } from "./commands/fill.js";
import { runXml } from "./commands/xmlCmd.js";
import { runReplace } from "./commands/replace.js";
import { runPatch } from "./commands/patch.js";
import { runNew } from "./commands/new.js";
import { runFormula } from "./commands/formula.js";

const program = new Command();

program.showHelpAfterError();
program.showSuggestionAfterError();

program
  .name("deword")
  .description("🪱 Create, read, edit, fill, patch, and add real equations to Word documents")
  .version("0.3.1")
  .addHelpText(
    "after",
    `
First-time agent quickstart:
  deword help
  deword new draft.docx
  deword read draft.docx -f summary
  deword formula draft.docx --replace "<empty>" --latex "E = mc^2"
  deword read draft.docx

Pick the right command:
  new      create a fresh .docx from the bundled template
  read     inspect content as markdown, json, summary, or model
  formula  insert a real Word equation (OMML / Math Mode)
  edit     replace exactly one unique text occurrence
  replace  replace all occurrences / fill placeholders in bulk
  fields   list form fields / checkboxes
  fill     fill form fields or checkboxes
  patch    do multi-step structural edits from JSON
  unpack   export markdown/json/xml/media for deep inspection
  xml      raw ZIP/XML escape hatch when nothing else fits

Examples:
  deword new report.docx
  deword formula report.docx --replace "<empty>" --latex "E = mc^2"
  deword read report.docx
  deword read report.docx -f model
  deword edit report.docx --old "Draft" --new "Final"
  deword replace report.docx --map replacements.json
  deword fields form.docx
  deword fill form.docx --field "I agree" --check
  deword patch report.docx -p patch.json -o final.docx
  deword xml report.docx --list

Recommended workflow:
  1. existing doc? deword read <file> -f summary
  2. structural work? deword read <file> -f model
  3. choose edit / replace / fill / formula / patch
  4. deword read <file>            # verify semantics
  5. open in Word for final visual check when layout matters

Notes:
  - .docx: full support for reading and editing
  - .doc (MHTML/Web Page): read only
  - legacy binary .doc: not supported
  - formulas are inserted as real Word equations, not plain text
  - use 'deword help <command>' for detailed command help
`
  );

program
  .command("new")
  .alias("create")
  .description("Create a new .docx file from the bundled template")
  .argument("<file>", "Path to the new .docx file")
  .option("--text <text>", "Replace the default <empty> placeholder with initial text")
  .option("-f, --force", "Overwrite the output file if it already exists")
  .addHelpText(
    "after",
    `
Examples:
  deword new note.docx
  deword new note.docx --text "Project kickoff notes"
  deword create drafts/memo.docx
`
  )
  .action(async (file: string, opts: { text?: string; force?: boolean }) => {
    try {
      await runNew(file, { text: opts.text, force: opts.force });
    } catch (err) {
      console.error(`Error: ${err instanceof Error ? err.message : err}`);
      process.exit(1);
    }
  });

program
  .command("formula")
  .alias("math")
  .description("Insert a math formula as a real Word equation (OMML)")
  .argument("<file>", "Path to .docx file")
  .requiredOption("--latex <formula>", 'LaTeX-style formula, e.g. "\\frac{a}{b} = c^2"')
  .option("--after <text>", "Insert after a uniquely matched body paragraph")
  .option("--before <text>", "Insert before a uniquely matched body paragraph")
  .option("--replace <text>", "Replace a uniquely matched body paragraph with the formula")
  .option("--append", "Append the formula at the end of the document body")
  .option("--match <mode>", "Paragraph match mode: exact, contains, regex")
  .option("--occurrence <n>", "Use the nth matching paragraph", (value) => parseInt(value, 10), 1)
  .option("--align <alignment>", "Equation alignment: left, center, right", "center")
  .option("-o, --output <file>", "Write to a different file instead of editing in-place")
  .addHelpText(
    "after",
    `
Examples:
  deword formula note.docx --replace "<empty>" --latex "E = mc^2"
  deword formula report.docx --after "Financial Model" --latex "\\frac{Revenue - Cost}{Revenue}"
  deword math report.docx --append --latex "x_1 + x_2 = y"

Supported syntax (common subset):
  superscripts/subscripts: x^2, x_1, x_1^2
  fractions: \\frac{a+b}{c}
  radicals: \\sqrt{x}, \\sqrt[n]{x}
  Greek/symbols: \\alpha, \\beta, \\pi, \\leq, \\geq, \\neq, \\infty
`
  )
  .action(
    async (
      file: string,
      opts: {
        latex: string;
        after?: string;
        before?: string;
        replace?: string;
        append?: boolean;
        match?: string;
        occurrence: number;
        align: string;
        output?: string;
      }
    ) => {
      try {
        await runFormula(file, {
          latex: opts.latex,
          after: opts.after,
          before: opts.before,
          replace: opts.replace,
          append: opts.append,
          output: opts.output,
          match: opts.match as "exact" | "contains" | "regex",
          occurrence: opts.occurrence,
          align: opts.align as "left" | "center" | "right",
        });
      } catch (err) {
        console.error(`Error: ${err instanceof Error ? err.message : err}`);
        process.exit(1);
      }
    }
  );

program
  .command("read")
  .description("Read a .docx or .doc (MHTML) file and output contents to stdout")
  .argument("<file>", "Path to .docx or .doc file")
  .option("-f, --format <format>", "Output format: markdown, json, summary, model", "markdown")
  .option("-i, --images <dir>", "Override image extraction directory (default: auto temp dir)")
  .addHelpText(
    "after",
    `
Formats:
  markdown  Human/agent-readable content with extracted image paths
  json      Structured content + metadata
  summary   Short metadata + preview
  model     Agent-friendly structure with paragraph/table/footnote IDs

Examples:
  deword read report.docx
  deword read report.docx -f model
  deword read report.docx -f json -i ./images
`
  )
  .action(async (file: string, opts: { format: string; images?: string }) => {
    try {
      await runRead(file, {
        format: opts.format as "markdown" | "json" | "summary" | "model",
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
  .addHelpText(
    "after",
    `
Use edit for one surgical replacement.
If the text appears more than once, deword fails and tells you to add context.

Example:
  deword edit report.docx --old "Draft Report" --new "Final Report"
`
  )
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
  .addHelpText(
    "after",
    `
Use replace for template filling or global replacements.

Examples:
  deword replace report.docx --old "{{NAME}}" --new "John Smith"
  deword replace report.docx --map replacements.json
`
  )
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

program
  .command("patch")
  .description("Apply a high-level JSON patch plan to a .docx file")
  .argument("<file>", "Path to .docx file")
  .option("-p, --patch <file>", "Path to patch JSON file (reads stdin if omitted)")
  .option("-i, --input <file>", "Alias for --patch")
  .option("-o, --output <file>", "Write to a different file instead of editing in-place")
  .addHelpText(
    "after",
    `
Patch is the main multi-step editing interface for agents.
Typical flow:
  deword read report.docx -f model
  deword patch report.docx -p patch.json

Common patch ops:
  replace_text, replace_all, edit_paragraph, insert_paragraph,
  insert_footnote, edit_footnote, insert_table, edit_table_cell,
  append_table_row, insert_image, set_region_text, fill_field, set_checkbox

Minimal example patch:
  {
    "version": "1.0",
    "operations": [
      {
        "op": "replace_text",
        "target": { "by_text": { "text": "Draft Report", "match": "exact" } },
        "new_text": "Final Report"
      }
    ]
  }
`
  )
  .action(async (file: string, opts: { patch?: string; input?: string; output?: string }) => {
    try {
      await runPatch(file, { patchFile: opts.patch, input: opts.input, output: opts.output });
    } catch (err) {
      console.error(`Error: ${err instanceof Error ? err.message : err}`);
      process.exit(1);
    }
  });

program.parse();
