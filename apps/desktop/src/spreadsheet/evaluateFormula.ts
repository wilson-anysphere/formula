import { parseA1Range, type RangeAddress } from "./a1.js";
import { tokenizeFormula } from "../formula-bar/highlight/tokenizeFormula.js";

export type SpreadsheetValue = number | string | boolean | null;
export type CellValue = SpreadsheetValue | SpreadsheetValue[];

function isErrorCode(value: unknown): value is string {
  return typeof value === "string" && value.startsWith("#");
}

function toNumber(value: SpreadsheetValue): number | null {
  if (value === null) return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return 0;
    const num = Number(trimmed);
    return Number.isFinite(num) ? num : null;
  }
  return null;
}

type EvalToken =
  | { type: "number"; value: number }
  | { type: "string"; value: string }
  | { type: "boolean"; value: boolean }
  | { type: "error"; value: string }
  | { type: "reference"; value: string }
  | { type: "function"; value: string }
  | { type: "operator"; value: string }
  | { type: "paren"; value: "(" | ")" }
  | { type: "comma"; value: "," };

function lex(formula: string): EvalToken[] {
  return tokenizeFormula(formula)
    .filter((token) => token.type !== "whitespace")
    .map((token): EvalToken => {
      switch (token.type) {
        case "number":
          return { type: "number", value: Number(token.text) };
        case "string":
          return { type: "string", value: token.text.slice(1, token.text.endsWith('"') ? -1 : token.text.length) };
        case "error":
          return { type: "error", value: token.text };
        case "reference":
          return { type: "reference", value: token.text };
        case "function":
          return { type: "function", value: token.text.toUpperCase() };
        case "identifier": {
          const upper = token.text.toUpperCase();
          if (upper === "TRUE") return { type: "boolean", value: true };
          if (upper === "FALSE") return { type: "boolean", value: false };
          return { type: "error", value: "#NAME?" };
        }
        case "operator":
          return { type: "operator", value: token.text };
        case "punctuation":
          if (token.text === "(" || token.text === ")") return { type: "paren", value: token.text };
          if (token.text === ",") return { type: "comma", value: "," };
          return { type: "error", value: "#VALUE!" };
        default:
          return { type: "error", value: "#VALUE!" };
      }
    });
}

class Parser {
  readonly tokens: EvalToken[];
  index = 0;

  constructor(tokens: EvalToken[]) {
    this.tokens = tokens;
  }

  peek(): EvalToken | null {
    return this.tokens[this.index] ?? null;
  }

  consume(): EvalToken | null {
    const tok = this.peek();
    if (tok) this.index += 1;
    return tok;
  }

  match(type: EvalToken["type"], value?: string): boolean {
    const tok = this.peek();
    if (!tok || tok.type !== type) return false;
    if (value !== undefined && "value" in tok && (tok as any).value !== value) return false;
    this.index += 1;
    return true;
  }
}

function flattenNumbers(values: CellValue[], out: number[]): string | null {
  for (const val of values) {
    if (isErrorCode(val)) return val;
    if (Array.isArray(val)) {
      const nested: CellValue[] = val as SpreadsheetValue[];
      const err = flattenNumbers(nested, out);
      if (err) return err;
      continue;
    }
    const num = toNumber(val as SpreadsheetValue);
    if (num !== null) out.push(num);
  }
  return null;
}

type GetCellValue = (address: string) => SpreadsheetValue;

function evalFunction(name: string, args: CellValue[], getCellValue: GetCellValue): SpreadsheetValue {
  const upper = name.toUpperCase();
  if (upper === "SUM") {
    const nums: number[] = [];
    const err = flattenNumbers(args, nums);
    if (err) return err;
    return nums.reduce((a, b) => a + b, 0);
  }

  if (upper === "AVERAGE") {
    const nums: number[] = [];
    const err = flattenNumbers(args, nums);
    if (err) return err;
    if (nums.length === 0) return 0;
    return nums.reduce((a, b) => a + b, 0) / nums.length;
  }

  if (upper === "IF") {
    const cond = args[0] ?? null;
    if (isErrorCode(cond)) return cond;
    const condNum = Array.isArray(cond) ? null : toNumber(cond as SpreadsheetValue);
    const truthy = condNum !== null ? condNum !== 0 : Boolean(cond);
    const chosen = truthy ? (args[1] ?? null) : (args[2] ?? null);
    if (isErrorCode(chosen)) return chosen;
    if (Array.isArray(chosen)) return (chosen[0] ?? null) as SpreadsheetValue;
    return chosen as SpreadsheetValue;
  }

  if (upper === "VLOOKUP") {
    return "#N/A";
  }

  return "#NAME?";
}

function readReference(refText: string, getCellValue: GetCellValue): CellValue {
  const range = parseA1Range(refText);
  if (!range) return "#REF!";
  if (range.start.row === range.end.row && range.start.col === range.end.col) {
    return getCellValue(toA1(range.start));
  }
  const values: SpreadsheetValue[] = [];
  for (let r = range.start.row; r <= range.end.row; r += 1) {
    for (let c = range.start.col; c <= range.end.col; c += 1) {
      const addr = toA1({ row: r, col: c });
      values.push(getCellValue(addr));
    }
  }
  return values;
}

function toA1(addr: { row: number; col: number }): string {
  let n = addr.col + 1;
  let col = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return `${col}${addr.row + 1}`;
}

function parsePrimary(parser: Parser, getCellValue: GetCellValue): CellValue {
  const tok = parser.peek();
  if (!tok) return null;

  if (tok.type === "error") {
    parser.consume();
    return tok.value;
  }

  if (tok.type === "number") {
    parser.consume();
    return tok.value;
  }

  if (tok.type === "string") {
    parser.consume();
    return tok.value;
  }

  if (tok.type === "boolean") {
    parser.consume();
    return tok.value;
  }

  if (tok.type === "reference") {
    parser.consume();
    return readReference(tok.value, getCellValue);
  }

  if (tok.type === "function") {
    const name = tok.value;
    parser.consume();
    if (!parser.match("paren", "(")) return "#VALUE!";

    const args: CellValue[] = [];
    if (!parser.match("paren", ")")) {
      while (true) {
        args.push(parseExpression(parser, getCellValue));
        if (parser.match("comma", ",")) continue;
        if (parser.match("paren", ")")) break;
        return "#VALUE!";
      }
    }

    return evalFunction(name, args, getCellValue);
  }

  if (tok.type === "paren" && tok.value === "(") {
    parser.consume();
    const inner = parseExpression(parser, getCellValue);
    if (!parser.match("paren", ")")) return "#VALUE!";
    return inner;
  }

  return "#VALUE!";
}

function parseUnary(parser: Parser, getCellValue: GetCellValue): CellValue {
  const tok = parser.peek();
  if (tok?.type === "operator" && (tok.value === "+" || tok.value === "-")) {
    parser.consume();
    const rhs = parseUnary(parser, getCellValue);
    if (isErrorCode(rhs)) return rhs;
    if (Array.isArray(rhs)) return "#VALUE!";
    const num = toNumber(rhs as SpreadsheetValue);
    if (num === null) return "#VALUE!";
    return tok.value === "-" ? -num : num;
  }
  return parsePrimary(parser, getCellValue);
}

function parseTerm(parser: Parser, getCellValue: GetCellValue): CellValue {
  let left = parseUnary(parser, getCellValue);
  while (true) {
    const tok = parser.peek();
    if (!tok || tok.type !== "operator" || (tok.value !== "*" && tok.value !== "/")) break;
    parser.consume();
    const right = parseUnary(parser, getCellValue);
    if (isErrorCode(left)) return left;
    if (isErrorCode(right)) return right;
    if (Array.isArray(left) || Array.isArray(right)) return "#VALUE!";
    const leftNum = toNumber(left as SpreadsheetValue);
    const rightNum = toNumber(right as SpreadsheetValue);
    if (leftNum === null || rightNum === null) return "#VALUE!";
    if (tok.value === "/") {
      if (rightNum === 0) return "#DIV/0!";
      left = leftNum / rightNum;
    } else {
      left = leftNum * rightNum;
    }
  }
  return left;
}

function parseExpression(parser: Parser, getCellValue: GetCellValue): CellValue {
  let left = parseTerm(parser, getCellValue);
  while (true) {
    const tok = parser.peek();
    if (!tok || tok.type !== "operator" || (tok.value !== "+" && tok.value !== "-")) break;
    parser.consume();
    const right = parseTerm(parser, getCellValue);
    if (isErrorCode(left)) return left;
    if (isErrorCode(right)) return right;
    if (Array.isArray(left) || Array.isArray(right)) return "#VALUE!";
    const leftNum = toNumber(left as SpreadsheetValue);
    const rightNum = toNumber(right as SpreadsheetValue);
    if (leftNum === null || rightNum === null) return "#VALUE!";
    left = tok.value === "+" ? leftNum + rightNum : leftNum - rightNum;
  }
  return left;
}

export function evaluateFormula(formulaText: string, getCellValue: GetCellValue): SpreadsheetValue {
  const text = formulaText.trim();
  if (!text.startsWith("=")) {
    const maybeNum = Number(text);
    if (text !== "" && Number.isFinite(maybeNum)) return maybeNum;
    return text === "" ? null : text;
  }

  const tokens = lex(text.slice(1));
  const parser = new Parser(tokens);
  const value = parseExpression(parser, getCellValue);
  if (isErrorCode(value)) return value;
  if (Array.isArray(value)) return (value[0] ?? null) as SpreadsheetValue;
  return value as SpreadsheetValue;
}

export function rangeToAddresses(range: RangeAddress): string[] {
  const out: string[] = [];
  for (let r = range.start.row; r <= range.end.row; r += 1) {
    for (let c = range.start.col; c <= range.end.col; c += 1) {
      out.push(toA1({ row: r, col: c }));
    }
  }
  return out;
}
