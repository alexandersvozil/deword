import * as docx from "docx";
import JSZip from "jszip";
import { writeFile } from "fs/promises";
import { dirname, resolve } from "path";
import { fileURLToPath } from "url";

const here = dirname(fileURLToPath(import.meta.url));
const outputPath = resolve(here, "latex_text_to_word.docx");

const mr = (text) => new docx.MathRun(text);
const sup = (base, exponent) =>
  new docx.MathSuperScript({ children: asMathArray(base), superScript: asMathArray(exponent) });
const sub = (base, subScript) =>
  new docx.MathSubScript({ children: asMathArray(base), subScript: asMathArray(subScript) });
const frac = (numerator, denominator) =>
  new docx.MathFraction({ numerator: asMathArray(numerator), denominator: asMathArray(denominator) });
const paren = (children) => new docx.MathRoundBrackets({ children: asMathArray(children) });
const math = (children, alignment = docx.AlignmentType.CENTER) =>
  new docx.Paragraph({
    alignment,
    spacing: { before: 120, after: 120 },
    children: [new docx.Math({ children: asMathArray(children) })],
  });

function asMathArray(value) {
  return Array.isArray(value) ? value : [value];
}

function bodyParagraph(children, options = {}) {
  return new docx.Paragraph({
    spacing: { after: 180 },
    ...options,
    children: Array.isArray(children) ? children : [new docx.TextRun(children)],
  });
}

function heading(text) {
  return new docx.Paragraph({
    text,
    heading: docx.HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
  });
}

function noBorderCell(children, widthPct, options = {}) {
  return new docx.TableCell({
    width: { size: widthPct, type: docx.WidthType.PERCENTAGE },
    borders: docx.TableBorders.NONE,
    margins: { top: 40, bottom: 40, left: 60, right: 60 },
    ...options,
    children: Array.isArray(children) ? children : [children],
  });
}

const numberedEquation = new docx.Table({
  width: { size: 100, type: docx.WidthType.PERCENTAGE },
  borders: docx.TableBorders.NONE,
  rows: [
    new docx.TableRow({
      children: [
        noBorderCell(
          math([mr("E"), mr("="), mr("m"), sup(mr("c"), mr("2"))]),
          85
        ),
        noBorderCell(
          new docx.Paragraph({
            alignment: docx.AlignmentType.RIGHT,
            spacing: { before: 120, after: 120 },
            children: [new docx.TextRun("(1)")],
          }),
          15
        ),
      ],
    }),
  ],
});

const sumEquation = math([
  new docx.MathSum({
    children: [mr("i")],
    subScript: [mr("i"), mr("="), mr("1")],
    superScript: [mr("N")],
  }),
  mr("="),
  frac(
    [mr("N"), paren([mr("N"), mr("+"), mr("1")])],
    [mr("2")]
  ),
]);

const alignEq1 = math([
  mr("f"),
  paren([mr("x")]),
  mr("="),
  sup(paren([mr("x"), mr("+"), mr("1")]), mr("2")),
]);

const alignEq2 = math([
  mr("="),
  sup(mr("x"), mr("2")),
  mr("+"),
  mr("2"),
  mr("x"),
  mr("+"),
  mr("1"),
]);

const matrixTable = new docx.Table({
  alignment: docx.AlignmentType.CENTER,
  width: { size: 52, type: docx.WidthType.PERCENTAGE },
  borders: docx.TableBorders.NONE,
  rows: [
    new docx.TableRow({
      children: [
        noBorderCell(new docx.Paragraph(""), 20),
        noBorderCell(new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [new docx.TextRun("⎛")] }), 8),
        noBorderCell(new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, children: [new docx.TextRun("a₁₁")] }), 32),
        noBorderCell(new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, children: [new docx.TextRun("a₁₂")] }), 32),
        noBorderCell(new docx.Paragraph({ alignment: docx.AlignmentType.LEFT, children: [new docx.TextRun("⎞")] }), 8),
      ],
    }),
    new docx.TableRow({
      children: [
        noBorderCell(new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [new docx.TextRun("A =")] }), 20),
        noBorderCell(new docx.Paragraph({ alignment: docx.AlignmentType.RIGHT, children: [new docx.TextRun("⎝")] }), 8),
        noBorderCell(new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, children: [new docx.TextRun("a₂₁")] }), 32),
        noBorderCell(new docx.Paragraph({ alignment: docx.AlignmentType.CENTER, children: [new docx.TextRun("a₂₂")] }), 32),
        noBorderCell(new docx.Paragraph({ alignment: docx.AlignmentType.LEFT, children: [new docx.TextRun("⎠")] }), 8),
      ],
    }),
  ],
});

const doc = new docx.Document({
  creator: "My Name",
  title: "A Sample Document with Mathematics",
  description: "Word rendering of benchmarks/latex_test/latex_text_to_word.tex",
  styles: {
    default: {
      document: {
        run: {
          font: "Times New Roman",
          size: 24,
        },
        paragraph: {
          spacing: { line: 276 },
        },
      },
    },
  },
  sections: [
    {
      children: [
        new docx.Paragraph({
          alignment: docx.AlignmentType.CENTER,
          spacing: { before: 240, after: 160 },
          children: [new docx.TextRun({ text: "A Sample Document with Mathematics", size: 34 })],
        }),
        new docx.Paragraph({
          alignment: docx.AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new docx.TextRun({ text: "My Name", size: 26 })],
        }),
        new docx.Paragraph({
          alignment: docx.AlignmentType.CENTER,
          spacing: { after: 280 },
          children: [new docx.TextRun({ text: "March 31, 2026", size: 24 })],
        }),

        heading("1 Introduction"),
        bodyParagraph([
          new docx.TextRun("LaTeX is excellent for writing documents with mathematical notation. Mathematical expressions can be included inline within a paragraph using single dollar signs, like this: the famous Pythagorean theorem is often written as "),
          new docx.Math({
            children: [sup(mr("a"), mr("2")), mr("+"), sup(mr("b"), mr("2")), mr("="), sup(mr("c"), mr("2"))],
          }),
          new docx.TextRun("."),
        ]),

        heading("2 Displayed Equations"),
        bodyParagraph("For equations that need to be on their own line and possibly numbered, the 'equation' environment is ideal."),
        numberedEquation,
        bodyParagraph("Equation 1 is the mass-energy equivalence principle."),
        bodyParagraph("For unnumbered displayed equations, you can use double dollar signs ('$$...$$') or '\\[...\\]' (the latter is preferred)."),
        sumEquation,

        heading("3 Advanced Mathematics"),
        bodyParagraph("The 'amsmath' package provides environments for more complex or aligned equations, such as the 'align*' environment (the asterisk prevents numbering)."),
        alignEq1,
        alignEq2,
        new docx.Paragraph({ children: [new docx.PageBreak()] }),
        bodyParagraph("We can also include matrices using environments like 'pmatrix' or 'array'."),
        matrixTable,
      ],
    },
  ],
});

const buffer = await docx.Packer.toBuffer(doc);
const normalizedBuffer = await forceBlackHeadingStyles(buffer);
await writeFile(outputPath, normalizedBuffer);
console.log(`Wrote ${outputPath}`);

async function forceBlackHeadingStyles(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const stylesFile = zip.file("word/styles.xml");
  if (!stylesFile) return buffer;

  let stylesXml = await stylesFile.async("string");
  stylesXml = stylesXml.replace(
    /(<w:style\b[^>]*w:styleId="Heading[1-9](?:Char)?"[\s\S]*?<w:rPr>[\s\S]*?)<w:color\b[^>]*\/>/g,
    "$1<w:color w:val=\"000000\"/>"
  );
  zip.file("word/styles.xml", stylesXml);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
}
