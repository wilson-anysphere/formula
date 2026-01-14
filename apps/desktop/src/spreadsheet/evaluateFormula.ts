import { parseA1Range, type RangeAddress } from "./a1.js";
import { tokenizeFormula } from "@formula/spreadsheet-frontend/formula/tokenizeFormula";

export type SpreadsheetValue = number | string | boolean | null;
export const PROVENANCE_REF_SEPARATOR = "\u001f";
export type ProvenanceCellValue = { __cellRef: string; value: SpreadsheetValue };
export type CellValue = SpreadsheetValue | SpreadsheetValue[] | ProvenanceCellValue | ProvenanceCellValue[];

export type AiFunctionArgumentProvenance = {
  /**
   * Canonical cell references referenced while evaluating the argument expression.
   *
   * These may be sheet-qualified (e.g. `Sheet1!A1`) depending on the caller's
   * `cellAddress` and the formula text.
   */
  cells: string[];
  /**
   * Canonical rectangular range references referenced while evaluating the argument expression.
   *
   * These may be sheet-qualified (e.g. `Sheet1!A1:B2`) depending on the caller's
   * `cellAddress` and the formula text.
   */
  ranges: string[];
};

export interface AiFunctionEvaluator {
  evaluateAiFunction(params: {
    name: string;
    args: CellValue[];
    cellAddress?: string;
    argProvenance?: AiFunctionArgumentProvenance[];
  }): SpreadsheetValue;
  /**
   * Optional range sampling limit for direct AI() range arguments.
   *
   * When provided, `evaluateFormula` will sample referenced ranges to this size when
   * producing AI function arguments, preventing it from materializing unbounded arrays.
   */
  rangeSampleLimit?: number;
}

export interface EvaluateFormulaOptions {
  ai?: AiFunctionEvaluator;
  cellAddress?: string;
  /**
   * Optional name resolver used by lightweight evaluation (e.g. formula-bar preview).
   *
   * When provided, identifiers that aren't TRUE/FALSE are looked up here and, if
   * resolved, treated as A1 references/ranges.
   */
  resolveNameToReference?: (name: string) => string | null;
  /**
   * When directly passing a range reference as an AI function argument, cap the number of
   * cells materialized to avoid unbounded range serialization.
   *
   * Nested non-AI functions (e.g. `SUM(A1:A1000)` inside an AI argument) still materialize
   * the full range so their semantics remain intact.
   */
  aiRangeSampleLimit?: number;
  /**
   * Safety guard: maximum number of cells to fully materialize for a rectangular range reference.
   *
   * This evaluator is used in the UI (e.g. formula previews, AI provenance, fallback computed values).
   * Fully materializing very large ranges (hundreds of thousands to millions of cells) can exhaust
   * memory or freeze the UI thread. When exceeded, the evaluator returns `#VALUE!`.
   *
   * Set to `Infinity` to disable the guard.
   */
  maxRangeCells?: number;
}

// Spreadsheet error codes are a small closed set in Excel (plus a couple of app-specific
// sentinel values like `#DLP!` for blocked AI calls). We intentionally *do not* treat
// arbitrary `#`-prefixed strings as errors (e.g. `#hashtag`), because those are common
// as plain text.
export const SPREADSHEET_ERROR_CODE_REGEX =
  /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|SPILL!|VALUE!|CALC!|GETTING_DATA|FIELD!|CONNECT!|BLOCKED!|UNKNOWN!|DLP!|AI!)$/;

export function isSpreadsheetErrorCode(value: unknown): value is string {
  return typeof value === "string" && SPREADSHEET_ERROR_CODE_REGEX.test(value);
}

function isErrorCode(value: unknown): value is string {
  return isSpreadsheetErrorCode(value);
}

function isProvenanceCellValue(value: unknown): value is ProvenanceCellValue {
  if (!value || typeof value !== "object") return false;
  const v = value as any;
  return typeof v.__cellRef === "string" && "value" in v;
}

function unwrapProvenance(value: CellValue): CellValue {
  return isProvenanceCellValue(value) ? value.value : value;
}

function splitProvenanceRefs(refs: string): string[] {
  return String(refs)
    .split(PROVENANCE_REF_SEPARATOR)
    .map((ref) => ref.trim())
    .filter(Boolean);
}

function rangeRefFromArray(value: CellValue): string | null {
  if (!Array.isArray(value)) return null;
  const ref = (value as any).__rangeRef;
  return typeof ref === "string" && ref.trim() ? ref.trim() : null;
}

function provenanceRefs(value: CellValue): string[] {
  if (isProvenanceCellValue(value)) return splitProvenanceRefs(value.__cellRef);
  const rangeRef = rangeRefFromArray(value);
  if (rangeRef) return [rangeRef];
  return [];
}

function wrapWithProvenance(value: SpreadsheetValue, refs: string[]): CellValue {
  const uniq = [...new Set(refs.map((r) => r.trim()).filter(Boolean))];
  if (uniq.length === 0) return value;
  return { __cellRef: uniq.join(PROVENANCE_REF_SEPARATOR), value };
}

function toNumber(value: CellValue): number | null {
  const unwrapped = unwrapProvenance(value);
  if (Array.isArray(unwrapped)) return null;
  const scalar = unwrapped as SpreadsheetValue;
  if (scalar === null) return 0;
  if (typeof scalar === "number") return Number.isFinite(scalar) ? scalar : null;
  if (typeof scalar === "boolean") return scalar ? 1 : 0;
  if (typeof scalar === "string") {
    const trimmed = scalar.trim();
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

function lex(formula: string, options: EvaluateFormulaOptions): EvalToken[] {
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
          const resolved = options.resolveNameToReference?.(token.text);
          if (resolved) return { type: "reference", value: resolved };
          return { type: "error", value: "#NAME?" };
        }
        case "operator":
          return { type: "operator", value: token.text };
        case "punctuation":
          if (token.text === "(" || token.text === ")") return { type: "paren", value: token.text };
          // Support locale-specific argument separators (`;` in many locales) by treating them
          // equivalently to commas. This evaluator is used in UI previews where locale-aware
          // parsing should avoid returning misleading errors.
          if (token.text === "," || token.text === ";") return { type: "comma", value: "," };
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
    if (Array.isArray(val)) {
      const nested: CellValue[] = val as CellValue[];
      const err = flattenNumbers(nested, out);
      if (err) return err;
      continue;
    }
    const scalar = unwrapProvenance(val);
    if (isErrorCode(scalar)) return scalar;
    const num = toNumber(scalar);
    if (num !== null) out.push(num);
  }
  return null;
}

type GetCellValue = (address: string) => SpreadsheetValue;

type EvalContext = {
  /**
   * When true, A1 references inside expressions are returned as provenance objects
   * (`{__cellRef, value}`) instead of raw scalars/arrays.
   *
   * This is used for AI cell functions so downstream DLP enforcement can resolve
   * classifications from the referenced cell address (not just the cell value).
   */
  preserveReferenceProvenance: boolean;
  /**
   * When true, multi-cell range references are sampled (first N) instead of fully materialized.
   */
  sampleRangeReferences: boolean;
  maxRangeCells: number;
};

const DEFAULT_CONTEXT: EvalContext = { preserveReferenceProvenance: false, sampleRangeReferences: false, maxRangeCells: 0 };

const DEFAULT_AI_RANGE_SAMPLE_LIMIT = 200;
const DEFAULT_AI_RANGE_SAMPLE_PREFIX = 30;
const DEFAULT_MAX_EVAL_RANGE_CELLS = 200_000;

function clampInt(value: number, opts: { min: number; max: number }): number {
  const n = Number.isFinite(value) ? Math.trunc(value) : opts.min;
  return Math.max(opts.min, Math.min(opts.max, n));
}

function splitSheetQualifier(input: string): { sheetName: string | null; ref: string } {
  const s = String(input).trim();

  const quoted = s.match(/^'((?:[^']|'')+)'!(.+)$/);
  if (quoted) {
    return { sheetName: quoted[1].replace(/''/g, "'"), ref: quoted[2] };
  }

  const unquoted = s.match(/^([^!]+)!(.+)$/);
  if (unquoted) return { sheetName: unquoted[1], ref: unquoted[2] };

  return { sheetName: null, ref: s };
}

function sheetIdFromCellAddress(cellAddress?: string): string | null {
  if (!cellAddress) return null;
  const bang = cellAddress.indexOf("!");
  if (bang === -1) return null;
  return cellAddress.slice(0, bang);
}

function evalFunction(
  name: string,
  args: CellValue[],
  getCellValue: GetCellValue,
  options: EvaluateFormulaOptions,
  context: EvalContext,
  argProvenance?: AiFunctionArgumentProvenance[],
): CellValue {
  const upper = name.toUpperCase();
  if (upper === "AI" || upper === "AI.EXTRACT" || upper === "AI.CLASSIFY" || upper === "AI.TRANSLATE") {
    return options.ai
      ? options.ai.evaluateAiFunction({ name: upper, args, cellAddress: options.cellAddress, argProvenance })
      : "#NAME?";
  }

  if (upper === "AND" || upper === "OR") {
    const refs: string[] = [];
    const isAnd = upper === "AND";
    for (const arg of args) {
      const scalar = unwrapProvenance(arg);
      if (isErrorCode(scalar)) return scalar;
      if (Array.isArray(scalar)) return "#VALUE!";
      const num = toNumber(scalar);
      const truthy = num !== null ? num !== 0 : Boolean(scalar);
      if (context.preserveReferenceProvenance) refs.push(...provenanceRefs(arg));
      if (isAnd && !truthy) return context.preserveReferenceProvenance ? wrapWithProvenance(false, refs) : false;
      if (!isAnd && truthy) return context.preserveReferenceProvenance ? wrapWithProvenance(true, refs) : true;
    }
    const value = isAnd;
    return context.preserveReferenceProvenance ? wrapWithProvenance(value, refs) : value;
  }

  if (upper === "NOT") {
    const arg = args[0] ?? null;
    const scalar = unwrapProvenance(arg);
    if (isErrorCode(scalar)) return scalar;
    if (Array.isArray(scalar)) return "#VALUE!";
    const num = toNumber(scalar);
    const truthy = num !== null ? num !== 0 : Boolean(scalar);
    const value = !truthy;
    if (!context.preserveReferenceProvenance) return value;
    return wrapWithProvenance(value, provenanceRefs(arg));
  }

  if (upper === "SUM") {
    const nums: number[] = [];
    const err = flattenNumbers(args, nums);
    if (err) return err;
    const value = nums.reduce((a, b) => a + b, 0);
    if (!context.preserveReferenceProvenance) return value;
    const refs = args.flatMap((arg) => provenanceRefs(arg));
    return wrapWithProvenance(value, refs);
  }

  if (upper === "AVERAGE") {
    const nums: number[] = [];
    const err = flattenNumbers(args, nums);
    if (err) return err;
    if (nums.length === 0) return 0;
    const value = nums.reduce((a, b) => a + b, 0) / nums.length;
    if (!context.preserveReferenceProvenance) return value;
    const refs = args.flatMap((arg) => provenanceRefs(arg));
    return wrapWithProvenance(value, refs);
  }

  if (upper === "IF") {
    const cond = args[0] ?? null;
    const condScalar = unwrapProvenance(cond);
    if (isErrorCode(condScalar)) return condScalar;
    const condNum = Array.isArray(condScalar) ? null : toNumber(condScalar);
    const truthy = condNum !== null ? condNum !== 0 : Boolean(condScalar);
    const chosen = truthy ? (args[1] ?? null) : (args[2] ?? null);
    const chosenScalar = Array.isArray(chosen) ? ((chosen[0] ?? null) as CellValue) : chosen;
    const chosenUnwrapped = unwrapProvenance(chosenScalar);
    if (isErrorCode(chosenUnwrapped)) return chosenUnwrapped;
    if (!context.preserveReferenceProvenance) return chosenScalar;
    const refs = [...provenanceRefs(cond), ...provenanceRefs(chosenScalar)];
    return wrapWithProvenance(chosenUnwrapped as SpreadsheetValue, refs);
  }

  if (upper === "IFERROR") {
    const first = args[0] ?? null;
    const firstScalar = unwrapProvenance(first);
    const refs: string[] = [];
    if (context.preserveReferenceProvenance) refs.push(...provenanceRefs(first));

    if (!isErrorCode(firstScalar)) {
      if (!context.preserveReferenceProvenance) return first;
      // Preserve provenance: unwrap scalar and wrap with any referenced cells.
      const scalar = Array.isArray(first) ? ((first[0] ?? null) as CellValue) : first;
      const unwrapped = unwrapProvenance(scalar);
      if (isErrorCode(unwrapped)) return unwrapped;
      return wrapWithProvenance(unwrapped as SpreadsheetValue, refs);
    }

    const fallback = args[1] ?? null;
    if (context.preserveReferenceProvenance) refs.push(...provenanceRefs(fallback));
    const scalar = Array.isArray(fallback) ? ((fallback[0] ?? null) as CellValue) : fallback;
    const unwrapped = unwrapProvenance(scalar);
    if (isErrorCode(unwrapped)) return unwrapped;
    if (!context.preserveReferenceProvenance) return scalar;
    return wrapWithProvenance(unwrapped as SpreadsheetValue, refs);
  }

  if (upper === "VLOOKUP") {
    return "#N/A";
  }

  return "#NAME?";
}

function readReference(
  refText: string,
  getCellValue: GetCellValue,
  options: EvaluateFormulaOptions,
  context: EvalContext
): CellValue {
  const { sheetName, ref } = splitSheetQualifier(refText);
  const range = parseA1Range(ref);
  if (!range) return "#REF!";

  const defaultSheet = sheetIdFromCellAddress(options.cellAddress);
  const provenanceSheet = sheetName ?? defaultSheet;
  const getPrefix = sheetName ? `${sheetName}!` : "";

  const readCell = (addr: string) => getCellValue(`${getPrefix}${addr}`);

  if (range.start.row === range.end.row && range.start.col === range.end.col) {
    const addr = toA1(range.start);
    const value = readCell(addr);
    if (!context.preserveReferenceProvenance) return value;
    const cellRef = provenanceSheet ? `${provenanceSheet}!${addr}` : addr;
    return { __cellRef: cellRef, value };
  }

  const totalCells = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
  const maxRangeCells = options.maxRangeCells ?? DEFAULT_MAX_EVAL_RANGE_CELLS;
  const maxRangeCellsNormalized =
    typeof maxRangeCells === "number" && Number.isFinite(maxRangeCells) ? Math.max(0, Math.trunc(maxRangeCells)) : Infinity;
  const willMaterializeFullRange = !context.sampleRangeReferences;
  if (willMaterializeFullRange && totalCells > maxRangeCellsNormalized) {
    return "#VALUE!";
  }

  if (!context.preserveReferenceProvenance) {
    const values: SpreadsheetValue[] = [];
    for (let r = range.start.row; r <= range.end.row; r += 1) {
      for (let c = range.start.col; c <= range.end.col; c += 1) {
        const addr = toA1({ row: r, col: c });
        values.push(readCell(addr));
      }
    }
    return values;
  }

  const maxCells = context.sampleRangeReferences ? Math.min(totalCells, Math.max(1, context.maxRangeCells)) : totalCells;

  const values: ProvenanceCellValue[] = [];
  const start = toA1(range.start);
  const end = toA1(range.end);
  const rangeRef = start === end ? start : `${start}:${end}`;
  const fullRangeRef = provenanceSheet ? `${provenanceSheet}!${rangeRef}` : rangeRef;

  const cols = range.end.col - range.start.col + 1;
  const addCellAtIndex = (index: number): void => {
    const rowOffset = Math.floor(index / cols);
    const colOffset = index % cols;
    const addr = toA1({ row: range.start.row + rowOffset, col: range.start.col + colOffset });
    const cellRef = provenanceSheet ? `${provenanceSheet}!${addr}` : addr;
    values.push({ __cellRef: cellRef, value: readCell(addr) });
  };

  // When sampling, include a prefix of the range (deterministic "top rows") plus a
  // deterministic random sample from the remainder. This avoids always sending only
  // the top of large ranges to the model.
  if (context.sampleRangeReferences && totalCells > maxCells) {
    const prefixCount = Math.min(maxCells, DEFAULT_AI_RANGE_SAMPLE_PREFIX);
    for (let i = 0; i < prefixCount; i += 1) addCellAtIndex(i);

    const remaining = maxCells - prefixCount;
    if (remaining > 0) {
      const seed = hashText(`${fullRangeRef}:${totalCells}:${maxCells}`);
      const rand = mulberry32(seed);
      const sample = pickSampleIndices({ total: totalCells - prefixCount, count: remaining, rand }).map((i) => i + prefixCount);
      for (const idx of sample) addCellAtIndex(idx);
    }
  } else {
    for (let i = 0; i < maxCells; i += 1) addCellAtIndex(i);
  }

  (values as any).__rangeRef = fullRangeRef;
  (values as any).__totalCells = totalCells;
  return values;
}

function hashText(text: string): number {
  // FNV-1a 32-bit for deterministic, dependency-free hashing.
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}

function mulberry32(seed: number): () => number {
  let t = seed >>> 0;
  return () => {
    t += 0x6d2b79f5;
    let x = Math.imul(t ^ (t >>> 15), 1 | t);
    x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
    return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
  };
}

function pickSampleIndices(params: { total: number; count: number; rand: () => number }): number[] {
  const count = Math.max(0, Math.min(params.count, params.total));
  if (count === 0) return [];
  const out = new Set<number>();
  const maxAttempts = Math.max(100, count * 20);
  let attempts = 0;

  while (out.size < count && attempts < maxAttempts) {
    attempts += 1;
    const idx = Math.floor(params.rand() * params.total);
    out.add(idx);
  }

  // Fall back to deterministic fill if collisions prevented reaching the target count.
  for (let i = 0; out.size < count && i < params.total; i += 1) out.add(i);

  return Array.from(out).sort((a, b) => a - b);
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

function isAiFunctionName(name: string): boolean {
  const upper = name.toUpperCase();
  return upper === "AI" || upper === "AI.EXTRACT" || upper === "AI.CLASSIFY" || upper === "AI.TRANSLATE";
}

function aiFunctionArgumentProvenance(value: CellValue): AiFunctionArgumentProvenance {
  const cells = new Set<string>();
  const ranges = new Set<string>();

  for (const refText of provenanceRefs(value)) {
    const cleaned = String(refText).replaceAll("$", "").trim();
    if (!cleaned) continue;
    const { sheetName, ref } = splitSheetQualifier(cleaned);
    const parsed = parseA1Range(ref);
    if (!parsed) continue;

    const start = toA1(parsed.start);
    const end = toA1(parsed.end);
    const prefix = sheetName ? `${sheetName}!` : "";

    if (start === end) cells.add(`${prefix}${start}`);
    else ranges.add(`${prefix}${start}:${end}`);
  }

  return { cells: Array.from(cells).sort(), ranges: Array.from(ranges).sort() };
}

function parsePrimary(
  parser: Parser,
  getCellValue: GetCellValue,
  options: EvaluateFormulaOptions,
  context: EvalContext,
): CellValue {
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
    return readReference(tok.value, getCellValue, options, context);
  }

  if (tok.type === "function") {
    const name = tok.value;
    parser.consume();
    if (!parser.match("paren", "(")) return "#VALUE!";

    const args: CellValue[] = [];
    const argContext = isAiFunctionName(name)
      ? {
          preserveReferenceProvenance: true,
          sampleRangeReferences: true,
          maxRangeCells: clampInt(
            options.aiRangeSampleLimit ?? options.ai?.rangeSampleLimit ?? DEFAULT_AI_RANGE_SAMPLE_LIMIT,
            { min: 1, max: 10_000 },
          ),
        }
      : { ...context, sampleRangeReferences: false };
    const argProvenance: AiFunctionArgumentProvenance[] = [];
    const isAiFn = isAiFunctionName(name) && Boolean(options.ai);
    if (!parser.match("paren", ")")) {
      while (true) {
        const argValue = parseComparison(parser, getCellValue, options, argContext);
        args.push(argValue);
        if (isAiFn) argProvenance.push(aiFunctionArgumentProvenance(argValue));
        if (parser.match("comma", ",")) continue;
        if (parser.match("paren", ")")) break;
        return "#VALUE!";
      }
    }

    return evalFunction(name, args, getCellValue, options, context, isAiFn ? argProvenance : undefined);
  }

  if (tok.type === "paren" && tok.value === "(") {
    parser.consume();
    const inner = parseComparison(parser, getCellValue, options, context);
    if (!parser.match("paren", ")")) return "#VALUE!";
    return inner;
  }

  return "#VALUE!";
}

function parseUnary(
  parser: Parser,
  getCellValue: GetCellValue,
  options: EvaluateFormulaOptions,
  context: EvalContext,
): CellValue {
  const tok = parser.peek();
  if (tok?.type === "operator" && (tok.value === "+" || tok.value === "-")) {
    parser.consume();
    const before = parser.index;
    const rhs = parseUnary(parser, getCellValue, options, context);
    if (parser.index === before) return "#VALUE!";
    const rhsScalar = unwrapProvenance(rhs);
    if (isErrorCode(rhsScalar)) return rhsScalar;
    if (Array.isArray(rhsScalar)) return "#VALUE!";
    const num = toNumber(rhsScalar);
    if (num === null) return "#VALUE!";
    const result = tok.value === "-" ? -num : num;
    if (!context.preserveReferenceProvenance) return result;
    return wrapWithProvenance(result, provenanceRefs(rhs));
  }
  return parsePrimary(parser, getCellValue, options, context);
}

function parseTerm(
  parser: Parser,
  getCellValue: GetCellValue,
  options: EvaluateFormulaOptions,
  context: EvalContext,
): CellValue {
  let left = parseUnary(parser, getCellValue, options, context);
  while (true) {
    const tok = parser.peek();
    if (!tok || tok.type !== "operator" || (tok.value !== "*" && tok.value !== "/")) break;
    parser.consume();
    const before = parser.index;
    const right = parseUnary(parser, getCellValue, options, context);
    if (parser.index === before) return "#VALUE!";
    const leftScalar = unwrapProvenance(left);
    const rightScalar = unwrapProvenance(right);
    if (isErrorCode(leftScalar)) return leftScalar;
    if (isErrorCode(rightScalar)) return rightScalar;
    if (Array.isArray(leftScalar) || Array.isArray(rightScalar)) return "#VALUE!";
    const leftNum = toNumber(leftScalar);
    const rightNum = toNumber(rightScalar);
    if (leftNum === null || rightNum === null) return "#VALUE!";
    const refs = context.preserveReferenceProvenance ? [...provenanceRefs(left), ...provenanceRefs(right)] : [];
    if (tok.value === "/") {
      if (rightNum === 0) return "#DIV/0!";
      const value = leftNum / rightNum;
      left = context.preserveReferenceProvenance ? wrapWithProvenance(value, refs) : value;
    } else {
      const value = leftNum * rightNum;
      left = context.preserveReferenceProvenance ? wrapWithProvenance(value, refs) : value;
    }
  }
  return left;
}

function parseExpression(
  parser: Parser,
  getCellValue: GetCellValue,
  options: EvaluateFormulaOptions,
  context: EvalContext,
): CellValue {
  let left = parseTerm(parser, getCellValue, options, context);
  while (true) {
    const tok = parser.peek();
    if (!tok || tok.type !== "operator" || (tok.value !== "+" && tok.value !== "-")) break;
    parser.consume();
    const before = parser.index;
    const right = parseTerm(parser, getCellValue, options, context);
    if (parser.index === before) return "#VALUE!";
    const leftScalar = unwrapProvenance(left);
    const rightScalar = unwrapProvenance(right);
    if (isErrorCode(leftScalar)) return leftScalar;
    if (isErrorCode(rightScalar)) return rightScalar;
    if (Array.isArray(leftScalar) || Array.isArray(rightScalar)) return "#VALUE!";
    const leftNum = toNumber(leftScalar);
    const rightNum = toNumber(rightScalar);
    if (leftNum === null || rightNum === null) return "#VALUE!";
    const refs = context.preserveReferenceProvenance ? [...provenanceRefs(left), ...provenanceRefs(right)] : [];
    const value = tok.value === "+" ? leftNum + rightNum : leftNum - rightNum;
    left = context.preserveReferenceProvenance ? wrapWithProvenance(value, refs) : value;
  }
  return left;
}

function parseConcat(
  parser: Parser,
  getCellValue: GetCellValue,
  options: EvaluateFormulaOptions,
  context: EvalContext,
): CellValue {
  let left = parseExpression(parser, getCellValue, options, context);
  while (true) {
    const tok = parser.peek();
    if (!tok || tok.type !== "operator" || tok.value !== "&") break;
    parser.consume();
    const before = parser.index;
    const right = parseExpression(parser, getCellValue, options, context);
    if (parser.index === before) return "#VALUE!";

    const leftScalar = unwrapProvenance(left);
    const rightScalar = unwrapProvenance(right);
    if (isErrorCode(leftScalar)) return leftScalar;
    if (isErrorCode(rightScalar)) return rightScalar;
    if (Array.isArray(leftScalar) || Array.isArray(rightScalar)) return "#VALUE!";

    const refs = context.preserveReferenceProvenance ? [...provenanceRefs(left), ...provenanceRefs(right)] : [];
    const value = `${leftScalar == null ? "" : String(leftScalar)}${rightScalar == null ? "" : String(rightScalar)}`;
    left = context.preserveReferenceProvenance ? wrapWithProvenance(value, refs) : value;
  }
  return left;
}

function compareScalars(left: SpreadsheetValue, right: SpreadsheetValue, op: string): boolean | string {
  // Match Excel's loose coercion in a simplified way: prefer numeric comparisons when both sides
  // can be coerced to numbers, otherwise fall back to case-insensitive string comparison.
  const leftNum = toNumber(left);
  const rightNum = toNumber(right);
  if (leftNum !== null && rightNum !== null) {
    switch (op) {
      case "=":
        return leftNum === rightNum;
      case "<>":
        return leftNum !== rightNum;
      case ">":
        return leftNum > rightNum;
      case ">=":
        return leftNum >= rightNum;
      case "<":
        return leftNum < rightNum;
      case "<=":
        return leftNum <= rightNum;
      default:
        return "#VALUE!";
    }
  }

  const l = (left == null ? "" : String(left)).toUpperCase();
  const r = (right == null ? "" : String(right)).toUpperCase();
  switch (op) {
    case "=":
      return l === r;
    case "<>":
      return l !== r;
    case ">":
      return l > r;
    case ">=":
      return l >= r;
    case "<":
      return l < r;
    case "<=":
      return l <= r;
    default:
      return "#VALUE!";
  }
}

function parseComparison(
  parser: Parser,
  getCellValue: GetCellValue,
  options: EvaluateFormulaOptions,
  context: EvalContext,
): CellValue {
  let left = parseConcat(parser, getCellValue, options, context);
  while (true) {
    const tok = parser.peek();
    if (
      !tok ||
      tok.type !== "operator" ||
      (tok.value !== "=" &&
        tok.value !== "<>" &&
        tok.value !== ">" &&
        tok.value !== ">=" &&
        tok.value !== "<" &&
        tok.value !== "<=")
    ) {
      break;
    }

    const op = tok.value;
    parser.consume();
    const before = parser.index;
    const right = parseConcat(parser, getCellValue, options, context);
    if (parser.index === before) return "#VALUE!";

    const leftScalar = unwrapProvenance(left);
    const rightScalar = unwrapProvenance(right);
    if (isErrorCode(leftScalar)) return leftScalar;
    if (isErrorCode(rightScalar)) return rightScalar;
    if (Array.isArray(leftScalar) || Array.isArray(rightScalar)) return "#VALUE!";

    const compared = compareScalars(leftScalar as SpreadsheetValue, rightScalar as SpreadsheetValue, op);
    if (typeof compared === "string") return compared;
    const refs = context.preserveReferenceProvenance ? [...provenanceRefs(left), ...provenanceRefs(right)] : [];
    left = context.preserveReferenceProvenance ? wrapWithProvenance(compared, refs) : compared;
  }
  return left;
}

export function evaluateFormula(
  formulaText: string,
  getCellValue: GetCellValue,
  options: EvaluateFormulaOptions = {}
): SpreadsheetValue {
  const text = formulaText.trim();
  if (!text.startsWith("=")) {
    const maybeNum = Number(text);
    if (text !== "" && Number.isFinite(maybeNum)) return maybeNum;
    return text === "" ? null : text;
  }

  const tokens = lex(text.slice(1), options);
  const parser = new Parser(tokens);
  const value = parseComparison(parser, getCellValue, options, DEFAULT_CONTEXT);
  // If we didn't consume the full token stream, treat this as invalid/unsupported syntax.
  if (parser.peek()) return "#VALUE!";
  if (isErrorCode(value)) return value;
  if (isProvenanceCellValue(value)) return value.value;
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
