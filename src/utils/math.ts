import JSZip from "jszip";
import {
  AlignmentType,
  Document,
  Math as DocxMath,
  MathFraction,
  MathFunction,
  MathRadical,
  MathRoundBrackets,
  MathCurlyBrackets,
  MathSquareBrackets,
  MathAngledBrackets,
  MathRun,
  MathSubScript,
  MathSubSuperScript,
  MathSuperScript,
  Packer,
  Paragraph,
  type MathComponent,
} from "docx";
import { findTopLevelBlocks, getBalancedTagRange } from "./wordModel.js";

export type FormulaAlignment = "left" | "center" | "right";

const SYMBOLS = new Map<string, string>([
  ["alpha", "α"], ["beta", "β"], ["gamma", "γ"], ["delta", "δ"], ["epsilon", "ε"], ["varepsilon", "ε"],
  ["zeta", "ζ"], ["eta", "η"], ["theta", "θ"], ["vartheta", "ϑ"], ["iota", "ι"], ["kappa", "κ"],
  ["lambda", "λ"], ["mu", "μ"], ["nu", "ν"], ["xi", "ξ"], ["pi", "π"], ["varpi", "ϖ"],
  ["rho", "ρ"], ["varrho", "ϱ"], ["sigma", "σ"], ["varsigma", "ς"], ["tau", "τ"], ["upsilon", "υ"],
  ["phi", "φ"], ["varphi", "ϕ"], ["chi", "χ"], ["psi", "ψ"], ["omega", "ω"],
  ["Gamma", "Γ"], ["Delta", "Δ"], ["Theta", "Θ"], ["Lambda", "Λ"], ["Xi", "Ξ"], ["Pi", "Π"],
  ["Sigma", "Σ"], ["Phi", "Φ"], ["Psi", "Ψ"], ["Omega", "Ω"],
  ["cdot", "·"], ["times", "×"], ["pm", "±"], ["mp", "∓"], ["neq", "≠"], ["ne", "≠"],
  ["leq", "≤"], ["geq", "≥"], ["approx", "≈"], ["sim", "∼"], ["to", "→"], ["rightarrow", "→"],
  ["leftarrow", "←"], ["leftrightarrow", "↔"], ["infty", "∞"], ["partial", "∂"], ["nabla", "∇"],
  ["forall", "∀"], ["exists", "∃"], ["in", "∈"], ["notin", "∉"], ["subset", "⊂"], ["subseteq", "⊆"],
  ["supset", "⊃"], ["supseteq", "⊇"], ["cup", "∪"], ["cap", "∩"], ["land", "∧"], ["lor", "∨"],
  ["sum", "∑"], ["prod", "∏"], ["int", "∫"], ["oint", "∮"], ["degree", "°"], ["cdots", "⋯"],
  ["ldots", "…"], ["dots", "…"], ["prime", "′"], ["neg", "¬"], ["oplus", "⊕"], ["otimes", "⊗"],
]);

const FUNCTIONS = new Set([
  "sin", "cos", "tan", "sec", "csc", "cot", "log", "ln", "exp", "lim", "max", "min", "det", "Pr"
]);

class LatexMathParser {
  private index = 0;

  constructor(private readonly input: string) {}

  parse(): MathComponent[] {
    const out = this.parseExpression();
    if (this.index < this.input.length) {
      throw new Error(`Unexpected trailing input near: ${this.input.slice(this.index, this.index + 20)}`);
    }
    return out;
  }

  private parseExpression(stoppers: string[] = []): MathComponent[] {
    const out: MathComponent[] = [];
    while (!this.eof()) {
      this.skipWhitespace();
      const ch = this.peek();
      if (!ch) break;
      if (stoppers.includes(ch)) break;
      out.push(...this.parseAtomWithScripts());
    }
    return out;
  }

  private parseAtomWithScripts(): MathComponent[] {
    let base = this.parseAtom();
    let subScript: MathComponent[] | undefined;
    let superScript: MathComponent[] | undefined;

    while (true) {
      this.skipWhitespace();
      const ch = this.peek();
      if (ch === "_") {
        this.index++;
        subScript = this.parseScriptArgument();
        continue;
      }
      if (ch === "^") {
        this.index++;
        superScript = this.parseScriptArgument();
        continue;
      }
      break;
    }

    if (base.length === 0) return [];
    if (subScript && superScript) {
      return [new MathSubSuperScript({ children: base, subScript, superScript })];
    }
    if (subScript) {
      return [new MathSubScript({ children: base, subScript })];
    }
    if (superScript) {
      return [new MathSuperScript({ children: base, superScript })];
    }
    return base;
  }

  private parseAtom(): MathComponent[] {
    this.skipWhitespace();
    const ch = this.peek();
    if (!ch) return [];

    if (ch === "{") return this.parseGroup();
    if (ch === "(") return this.parseBracketGroup("(");
    if (ch === "[") return this.parseBracketGroup("[");
    if (ch === "⟨") return this.parseBracketGroup("⟨");
    if (ch === "<") return this.parseBracketGroup("<");
    if (ch === "\\") return this.parseCommand();

    return [new MathRun(this.nextCodePoint())];
  }

  private parseGroup(): MathComponent[] {
    this.expect("{");
    const inner = this.parseExpression(["}"]);
    this.expect("}");
    return inner;
  }

  private parseBracketGroup(open: string): MathComponent[] {
    const { close, wrap } = this.bracketSpec(open);
    this.expect(open);
    const inner = this.parseExpression([close]);
    this.expect(close);
    return [wrap(inner)];
  }

  private parseScriptArgument(): MathComponent[] {
    this.skipWhitespace();
    const ch = this.peek();
    if (!ch) throw new Error("Expected script argument after ^ or _.");
    if (ch === "{") return this.parseGroup();
    if (ch === "(") return this.parseBracketGroup("(");
    if (ch === "[") return this.parseBracketGroup("[");
    if (ch === "\\") return this.parseCommand();
    return [new MathRun(this.nextCodePoint())];
  }

  private parseCommand(): MathComponent[] {
    this.expect("\\");
    const next = this.peek();
    if (!next) throw new Error("Dangling backslash at end of formula.");

    const name = /[A-Za-z]/.test(next) ? this.readWhile(/[A-Za-z]/) : this.nextCodePoint();

    if (name === "left" || name === "right") {
      return [];
    }

    if (name === "," || name === ";" || name === "!" || name === " ") {
      return [new MathRun(" ")];
    }

    if (name === "frac") {
      const numerator = this.parseRequiredGroup("\\frac numerator");
      const denominator = this.parseRequiredGroup("\\frac denominator");
      return [new MathFraction({ numerator, denominator })];
    }

    if (name === "sqrt") {
      this.skipWhitespace();
      let degree: MathComponent[] | undefined;
      if (this.peek() === "[") {
        this.expect("[");
        degree = this.parseExpression(["]"]);
        this.expect("]");
      }
      const children = this.parseRequiredGroup("\\sqrt radicand");
      return [new MathRadical({ children, degree })];
    }

    if (name === "text" || name === "mathrm" || name === "operatorname") {
      const raw = this.parseRawBraceText(name);
      return [...raw].map((ch) => new MathRun(ch));
    }

    if (FUNCTIONS.has(name)) {
      this.skipWhitespace();
      if (this.eof()) {
        return [new MathRun(name)];
      }
      const nextChar = this.peek();
      if (!nextChar || ["+", "-", "=", ">", "<", ",", ";", ")", "]", "}"] .includes(nextChar)) {
        return [new MathRun(name)];
      }
      const argument = this.parseAtomWithScripts();
      return [new MathFunction({ name: [new MathRun(name)], children: argument.length ? argument : [new MathRun("")] })];
    }

    const symbol = SYMBOLS.get(name);
    if (symbol) {
      return [new MathRun(symbol)];
    }

    if (name.length === 1) {
      return [new MathRun(name)];
    }

    throw new Error(`Unsupported LaTeX command: \\${name}`);
  }

  private parseRequiredGroup(context: string): MathComponent[] {
    this.skipWhitespace();
    if (this.peek() !== "{") {
      throw new Error(`Expected { ... } for ${context}.`);
    }
    return this.parseGroup();
  }

  private parseRawBraceText(commandName: string): string {
    this.skipWhitespace();
    if (this.peek() !== "{") {
      throw new Error(`Expected { ... } after \\${commandName}.`);
    }
    this.expect("{");
    let depth = 1;
    let out = "";
    while (!this.eof() && depth > 0) {
      const ch = this.nextCodePoint();
      if (ch === "{") {
        depth++;
        if (depth > 1) out += ch;
        continue;
      }
      if (ch === "}") {
        depth--;
        if (depth > 0) out += ch;
        continue;
      }
      out += ch;
    }
    if (depth !== 0) {
      throw new Error(`Unclosed { ... } after \\${commandName}.`);
    }
    return out;
  }

  private bracketSpec(open: string): {
    close: string;
    wrap: (children: MathComponent[]) => MathComponent;
  } {
    if (open === "(") return { close: ")", wrap: (children) => new MathRoundBrackets({ children }) };
    if (open === "[") return { close: "]", wrap: (children) => new MathSquareBrackets({ children }) };
    if (open === "⟨" || open === "<") return { close: open === "<" ? ">" : "⟩", wrap: (children) => new MathAngledBrackets({ children }) };
    return { close: "}", wrap: (children) => new MathCurlyBrackets({ children }) };
  }

  private readWhile(re: RegExp): string {
    let out = "";
    while (!this.eof() && re.test(this.peek()!)) {
      out += this.nextCodePoint();
    }
    return out;
  }

  private skipWhitespace(): void {
    while (!this.eof() && /\s/.test(this.peek()!)) {
      this.index++;
    }
  }

  private peek(): string | undefined {
    if (this.index >= this.input.length) return undefined;
    const cp = this.input.codePointAt(this.index);
    return cp == null ? undefined : String.fromCodePoint(cp);
  }

  private nextCodePoint(): string {
    const cp = this.input.codePointAt(this.index);
    if (cp == null) return "";
    const ch = String.fromCodePoint(cp);
    this.index += ch.length;
    return ch;
  }

  private expect(ch: string): void {
    const actual = this.nextCodePoint();
    if (actual !== ch) {
      throw new Error(`Expected '${ch}' but found '${actual || "<eof>"}'.`);
    }
  }

  private eof(): boolean {
    return this.index >= this.input.length;
  }
}

export function parseLatexMath(input: string): MathComponent[] {
  const trimmed = input.trim();
  if (!trimmed) {
    throw new Error("Formula is empty.");
  }
  const parsed = new LatexMathParser(trimmed).parse();
  if (parsed.length === 0) {
    throw new Error("Formula did not produce any math content.");
  }
  return parsed;
}

function mapAlignment(alignment: FormulaAlignment) {
  switch (alignment) {
    case "left":
      return AlignmentType.LEFT;
    case "right":
      return AlignmentType.RIGHT;
    default:
      return AlignmentType.CENTER;
  }
}

export async function buildFormulaParagraphXml(input: string, alignment: FormulaAlignment = "center"): Promise<string> {
  const mathChildren = parseLatexMath(input);
  const paragraph = new Paragraph({
    alignment: mapAlignment(alignment),
    children: [new DocxMath({ children: mathChildren })],
  });

  const doc = new Document({
    sections: [{ children: [paragraph] }],
  });

  const buffer = await Packer.toBuffer(doc);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")?.async("string");
  if (!xml) {
    throw new Error("Failed to generate formula XML.");
  }

  const body = getBalancedTagRange(xml, "w:body");
  if (!body) {
    throw new Error("Generated formula document is missing <w:body>.");
  }

  const paragraphBlock = findTopLevelBlocks(body.innerXml, ["w:p"])[0]?.xml;
  if (!paragraphBlock) {
    throw new Error("Generated formula document is missing the equation paragraph.");
  }

  return paragraphBlock;
}

export function ensureMathNamespace(documentXml: string): string {
  if (/xmlns:m="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/math"/.test(documentXml)) {
    return documentXml;
  }
  return documentXml.replace(
    /<w:document\b/,
    '<w:document xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
  );
}
