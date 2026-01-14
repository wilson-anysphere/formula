import { parseA1Range, type RangeAddress } from "./a1.js";
import { tokenizeFormula } from "@formula/spreadsheet-frontend/formula/tokenizeFormula";

// Translation tables from the Rust engine (canonical <-> localized function names).
// Keep these in sync with `crates/formula-engine/src/locale/data/*.tsv`.
//
// This lightweight evaluator is used in UI contexts (formula-bar previews, AI provenance
// fallbacks). When the UI locale uses localized function names (e.g. de-DE `SUMME`),
// we canonicalize them so built-in function handlers (SUM, IF, ...) still work.
import DE_DE_FUNCTION_TSV from "../../../../crates/formula-engine/src/locale/data/de-DE.tsv?raw";
import ES_ES_FUNCTION_TSV from "../../../../crates/formula-engine/src/locale/data/es-ES.tsv?raw";
import FR_FR_FUNCTION_TSV from "../../../../crates/formula-engine/src/locale/data/fr-FR.tsv?raw";

import DE_DE_ERRORS_TSV from "../../../../crates/formula-engine/src/locale/data/de-DE.errors.tsv?raw";
import ES_ES_ERRORS_TSV from "../../../../crates/formula-engine/src/locale/data/es-ES.errors.tsv?raw";
import FR_FR_ERRORS_TSV from "../../../../crates/formula-engine/src/locale/data/fr-FR.errors.tsv?raw";
import { normalizeFormulaLocaleId } from "./formulaLocale.js";

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
   * Optional workbook file metadata used by Excel-compatible worksheet information functions
   * like `CELL("filename")` and `INFO("directory")`.
   *
   * When omitted (or when `filename` is missing), these functions should behave like an unsaved
   * workbook and return `""`.
   */
  workbookFileMetadata?: { directory: string | null; filename: string | null } | null;
  /**
   * Optional current-sheet display name.
   *
   * This is used by functions like `CELL("filename")` when no explicit reference argument is
   * provided, or when the reference is unqualified (e.g. `A1`).
   *
   * When omitted, callers can still provide `cellAddress: "Sheet1!A1"` and the evaluator will
   * best-effort infer the sheet name from that.
   */
  currentSheetName?: string;
  /**
   * Optional locale identifier used for canonicalizing localized function names (e.g. de-DE `SUMME` -> `SUM`).
   *
   * Callers should generally pass the current UI/workbook locale when evaluating user-visible previews.
   * When omitted, the evaluator assumes canonical English function names.
   */
  localeId?: string;
  /**
   * Optional name resolver used by lightweight evaluation (e.g. formula-bar preview).
   *
   * When provided, identifiers that aren't TRUE/FALSE are looked up here and, if
   * resolved, treated as A1 references/ranges.
   */
  resolveNameToReference?: (name: string) => string | null;
  /**
   * Optional structured-reference resolver used by lightweight evaluation (e.g. formula-bar previews).
   *
   * When provided, reference tokens that are not valid A1 refs/ranges (e.g. `Table1[Amount]`)
   * are passed through this resolver before falling back to `#REF!`.
   *
   * The resolver should return an A1 reference/range string, optionally sheet-qualified
   * (e.g. `Sheet1!A2:A10`).
   */
  resolveStructuredRefToReference?: (refText: string) => string | null;
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

function casefoldIdent(ident: string): string {
  // Mirror Rust's locale behavior (`casefold_ident` / `casefold`): Unicode-aware uppercasing.
  return String(ident ?? "").toUpperCase();
}

type FunctionTranslationMap = Map<string, string>;

function parseFunctionTranslationsTsv(tsv: string): FunctionTranslationMap {
  const localizedToCanonical: FunctionTranslationMap = new Map();
  for (const rawLine of String(tsv ?? "").split(/\r?\n/)) {
    const trimmed = rawLine.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;

    // Parse the raw line (not the trimmed line) so trailing empty columns (`SUM\tSUMME\t`)
    // do not silently pass.
    const parts = rawLine.split("\t");
    if (parts.length !== 2) continue;
    const canonical = parts[0].trim();
    const localized = parts[1].trim();
    if (!canonical || !localized) continue;

    const canonUpper = casefoldIdent(canonical);
    const locUpper = casefoldIdent(localized);
    // Only store translations that differ; identity entries can fall back to `casefoldIdent`.
    if (canonUpper && locUpper && canonUpper !== locUpper) {
      localizedToCanonical.set(locUpper, canonUpper);
    }
  }
  return localizedToCanonical;
}

const FUNCTION_TRANSLATIONS_BY_LOCALE: Record<string, FunctionTranslationMap> = {
  "de-DE": parseFunctionTranslationsTsv(DE_DE_FUNCTION_TSV),
  "fr-FR": parseFunctionTranslationsTsv(FR_FR_FUNCTION_TSV),
  "es-ES": parseFunctionTranslationsTsv(ES_ES_FUNCTION_TSV),
};

type ErrorTranslationMap = Map<string, string>;

function parseErrorTranslationsTsv(tsv: string): ErrorTranslationMap {
  const localizedToCanonical: ErrorTranslationMap = new Map();
  for (const rawLine of String(tsv ?? "").split(/\r?\n/)) {
    const trimmed = rawLine.trim();
    if (!trimmed) continue;
    // Error literals themselves start with `#`, so treat comments as `#` followed by whitespace
    // (or a bare `#`) rather than treating all `#` lines as comments.
    const isComment = trimmed === "#" || (trimmed.startsWith("#") && /\s/u.test(trimmed[1] ?? ""));
    if (isComment) continue;

    // Parse the raw line (not the trimmed line) so trailing empty columns (`#VALUE!\t#WERT!\t`)
    // do not silently pass.
    const parts = rawLine.split("\t");
    if (parts.length !== 2) continue;
    const canonical = parts[0].trim();
    const localized = parts[1].trim();
    if (!canonical || !localized) continue;
    if (!canonical.startsWith("#") || !localized.startsWith("#")) continue;

    const canonUpper = casefoldIdent(canonical);
    const locUpper = casefoldIdent(localized);
    // Only store translations that differ; identity entries can fall back to `casefoldIdent`.
    if (canonUpper && locUpper && canonUpper !== locUpper) {
      localizedToCanonical.set(locUpper, canonUpper);
    }
  }
  return localizedToCanonical;
}

const ERROR_TRANSLATIONS_BY_LOCALE: Record<string, ErrorTranslationMap> = {
  "de-DE": parseErrorTranslationsTsv(DE_DE_ERRORS_TSV),
  "fr-FR": parseErrorTranslationsTsv(FR_FR_ERRORS_TSV),
  "es-ES": parseErrorTranslationsTsv(ES_ES_ERRORS_TSV),
};

type NumberLocaleConfig = {
  decimalSeparator: "." | ",";
  thousandsSeparator: "." | "\u00A0" | "\u202F" | null;
};

function getNumberLocaleConfig(localeId?: string): NumberLocaleConfig {
  // Mirror `formula_engine::LocaleConfig` defaults for locales the WASM engine currently ships.
  switch (normalizeFormulaLocaleId(localeId)) {
    case "de-DE":
    case "es-ES":
      return { decimalSeparator: ",", thousandsSeparator: "." };
    case "fr-FR":
      return { decimalSeparator: ",", thousandsSeparator: "\u00A0" };
    default:
      // en-US (canonical)
      return { decimalSeparator: ".", thousandsSeparator: null };
  }
}

function splitNumericExponent(raw: string): { mantissa: string; exponent: string } {
  // Port of `formula_engine::LocaleConfig::split_numeric_exponent`.
  if (!/[eE]/.test(raw)) return { mantissa: raw, exponent: "" };
  for (let idx = 0; idx < raw.length; idx += 1) {
    const ch = raw[idx];
    if (ch !== "e" && ch !== "E") continue;
    let rest = raw.slice(idx + 1);
    if (rest.startsWith("+") || rest.startsWith("-")) rest = rest.slice(1);
    if (rest.length === 0) continue;
    if (!/^\d+$/.test(rest)) continue;
    return { mantissa: raw.slice(0, idx), exponent: raw.slice(idx) };
  }
  return { mantissa: raw, exponent: "" };
}

function looksLikeThousandsGrouping(mantissa: string, sep: "."): boolean {
  const parts = mantissa.split(sep);
  const first = parts.shift();
  if (!first) return false;
  if (first.length === 0 || first.length > 3 || !/^\d+$/.test(first)) return false;
  if (parts.length === 0) return false;
  return parts.every((p) => p.length === 3 && /^\d+$/.test(p));
}

function parseLocaleNumber(raw: string, locale: NumberLocaleConfig): number | null {
  // Port of `formula_engine::LocaleConfig::parse_number`.
  const trimmed = String(raw ?? "").trim();
  if (!trimmed) return null;

  const { mantissa: mantissaWithSign, exponent } = splitNumericExponent(trimmed);
  let sign = "";
  let mantissa = mantissaWithSign;
  if (mantissa.startsWith("+") || mantissa.startsWith("-")) {
    sign = mantissa[0] ?? "";
    mantissa = mantissa.slice(1);
  }
  if (!mantissa) return null;

  let decimal: "." | "," | null = null;
  if (mantissa.includes(locale.decimalSeparator)) decimal = locale.decimalSeparator;
  else if (mantissa.includes(".")) decimal = ".";

  // Disambiguate locales where the thousands separator collides with the canonical decimal separator,
  // mirroring the Rust behavior (`de-DE`: '.' grouping, ',' decimal).
  if (
    decimal === "." &&
    locale.decimalSeparator !== "." &&
    locale.thousandsSeparator === "." &&
    looksLikeThousandsGrouping(mantissa, ".")
  ) {
    decimal = null;
  }

  let out = sign;
  let decimalUsed = false;

  for (const ch of mantissa) {
    if (ch >= "0" && ch <= "9") {
      out += ch;
      continue;
    }

    if (decimal && ch === decimal) {
      if (decimalUsed) return null;
      out += ".";
      decimalUsed = true;
      continue;
    }

    const isThousands =
      locale.thousandsSeparator === ch ||
      // Some spreadsheets use narrow NBSP (U+202F) instead of NBSP (U+00A0); accept both when configured.
      (locale.thousandsSeparator === "\u00A0" && ch === "\u202F") ||
      (locale.thousandsSeparator === "\u202F" && ch === "\u00A0");
    if (isThousands && (!decimal || ch !== decimal)) {
      continue;
    }

    return null;
  }

  out += exponent;
  const n = Number(out);
  return Number.isFinite(n) ? n : null;
}

function canonicalizeFunctionNameForLocale(name: string, localeId?: string): string {
  const raw = String(name ?? "");
  if (!raw) return raw;

  const formulaLocaleId = normalizeFormulaLocaleId(localeId);
  const map = formulaLocaleId ? FUNCTION_TRANSLATIONS_BY_LOCALE[formulaLocaleId] : undefined;

  // Mirror `formula_engine::locale::registry::FormulaLocale::canonical_function_name`.
  const PREFIX = "_xlfn.";
  const hasPrefix = raw.length >= PREFIX.length && raw.slice(0, PREFIX.length).toLowerCase() === PREFIX;
  const base = hasPrefix ? raw.slice(PREFIX.length) : raw;
  const upper = casefoldIdent(base);
  const mapped = map?.get(upper) ?? upper;
  return hasPrefix ? `${PREFIX}${mapped}` : mapped;
}

function localizedBooleanLiteral(identUpper: string, localeId?: string): boolean | null {
  switch (normalizeFormulaLocaleId(localeId)) {
    case "de-DE": {
      if (identUpper === "WAHR") return true;
      if (identUpper === "FALSCH") return false;
      return null;
    }
    case "fr-FR": {
      if (identUpper === "VRAI") return true;
      if (identUpper === "FAUX") return false;
      return null;
    }
    case "es-ES": {
      if (identUpper === "VERDADERO") return true;
      if (identUpper === "FALSO") return false;
      return null;
    }
    default:
      return null;
  }
}

function canonicalizeErrorCodeForLocale(errorLiteral: string, localeId?: string): string {
  const raw = String(errorLiteral ?? "").trim();
  if (!raw.startsWith("#")) return raw;
  const upper = casefoldIdent(raw);

  // Excel (and our engine) treat `#N/A` as the canonical form, but many spreadsheets emit `#N/A!`.
  if (upper === "#N/A!") return "#N/A";

  if (isSpreadsheetErrorCode(upper)) return upper;

  const formulaLocaleId = normalizeFormulaLocaleId(localeId);
  const map = formulaLocaleId ? ERROR_TRANSLATIONS_BY_LOCALE[formulaLocaleId] : undefined;
  const mapped = map?.get(upper) ?? upper;
  return isSpreadsheetErrorCode(mapped) ? mapped : raw;
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
  const locale = getNumberLocaleConfig(options.localeId);
  const tokens = tokenizeFormula(formula).filter((token) => token.type !== "whitespace");
  const out: EvalToken[] = [];

  for (let i = 0; i < tokens.length; i += 1) {
    const token = tokens[i]!;

    // Merge locale-specific decimal-comma and thousands-separator constructs into a single number token.
    if (token.type === "number") {
      let raw = token.text;
      let end = token.end;
      let j = i;

      const canThousandsMerge = () => {
        if (locale.thousandsSeparator !== ".") return false;
        // Avoid merging when the mantissa doesn't look like a grouping prefix (e.g. `1.2`).
        const { mantissa, exponent } = splitNumericExponent(raw);
        if (exponent) return false;
        const m = mantissa.trim();
        if (!m) return false;
        // Accept a leading sign.
        const normalized = m.startsWith("+") || m.startsWith("-") ? m.slice(1) : m;
        return /^\d{1,3}(?:\.\d{3})*$/.test(normalized);
      };

      // Merge repeated thousands group separators (e.g. `1.234.567` in de-DE).
      while (
        canThousandsMerge() &&
        // Tokenizer can represent the additional group as either:
        // - punctuation "." + number "567", or
        // - number ".567" (because numbers may start with a leading decimal point).
        // Support both shapes so inputs like `1.234.567,89` in de-DE parse correctly.
        ((tokens[j + 1]?.type === "punctuation" &&
          tokens[j + 1]?.text === "." &&
          tokens[j + 2]?.type === "number" &&
          /^\d{3}$/.test(tokens[j + 2]!.text) &&
          tokens[j + 1]!.start === end &&
          tokens[j + 2]!.start === tokens[j + 1]!.end) ||
          (tokens[j + 1]?.type === "number" &&
            /^\.\d{3}$/.test(tokens[j + 1]!.text) &&
            tokens[j + 1]!.start === end))
      ) {
        const next = tokens[j + 1]!;
        if (next.type === "number") {
          raw += next.text;
          end = next.end;
          j += 1;
          continue;
        }
        raw += `.${tokens[j + 2]!.text}`;
        end = tokens[j + 2]!.end;
        j += 2;
      }

      // Merge NBSP-based thousands separators (e.g. `1 234,56` or `1 234,56` in fr-FR).
      if (locale.thousandsSeparator === "\u00A0" || locale.thousandsSeparator === "\u202F") {
        while (
          tokens[j + 1]?.type === "unknown" &&
          (tokens[j + 1]?.text === "\u00A0" || tokens[j + 1]?.text === "\u202F") &&
          tokens[j + 2]?.type === "number" &&
          /^\d{3}$/.test(tokens[j + 2]!.text) &&
          tokens[j + 1]!.start === end &&
          tokens[j + 2]!.start === tokens[j + 1]!.end
        ) {
          raw += `${tokens[j + 1]!.text}${tokens[j + 2]!.text}`;
          end = tokens[j + 2]!.end;
          j += 2;
        }
      }

      // Merge decimal comma (e.g. `1,5` in de-DE).
      if (
        locale.decimalSeparator === "," &&
        tokens[j + 1]?.type === "punctuation" &&
        tokens[j + 1]?.text === "," &&
        tokens[j + 2]?.type === "number" &&
        tokens[j + 1]!.start === end &&
        tokens[j + 2]!.start === tokens[j + 1]!.end
      ) {
        raw += `,${tokens[j + 2]!.text}`;
        end = tokens[j + 2]!.end;
        j += 2;
      }

      const value = parseLocaleNumber(raw, locale);
      if (value == null) {
        out.push({ type: "error", value: "#VALUE!" });
      } else {
        out.push({ type: "number", value });
      }
      i = j;
      continue;
    }

    switch (token.type) {
      case "string":
        out.push({ type: "string", value: token.text.slice(1, token.text.endsWith('"') ? -1 : token.text.length) });
        break;
      case "error":
        out.push({ type: "error", value: canonicalizeErrorCodeForLocale(token.text, options.localeId) });
        break;
      case "reference":
        out.push({ type: "reference", value: token.text });
        break;
      case "function":
        out.push({ type: "function", value: canonicalizeFunctionNameForLocale(token.text, options.localeId) });
        break;
      case "identifier": {
        const upper = casefoldIdent(token.text);
        if (upper === "TRUE") {
          out.push({ type: "boolean", value: true });
          break;
        }
        if (upper === "FALSE") {
          out.push({ type: "boolean", value: false });
          break;
        }
        const localizedBool = localizedBooleanLiteral(upper, options.localeId);
        if (localizedBool !== null) {
          out.push({ type: "boolean", value: localizedBool });
          break;
        }
        const resolved = options.resolveNameToReference?.(token.text);
        if (resolved) {
          out.push({ type: "reference", value: resolved });
          break;
        }

        // `tokenizeFormula` only labels a token as a "function" when the opening paren is adjacent
        // (e.g. `SUM(A1)`). Excel permits whitespace between a function name and `(`, so treat an
        // unresolved identifier followed by `(` as a function call too (e.g. `SUM (A1)`).
        if (tokens[i + 1]?.type === "punctuation" && tokens[i + 1]?.text === "(") {
          out.push({ type: "function", value: canonicalizeFunctionNameForLocale(token.text, options.localeId) });
          break;
        }

        out.push({ type: "error", value: "#NAME?" });
        break;
      }
      case "operator":
        out.push({ type: "operator", value: token.text });
        break;
      case "punctuation":
        if (token.text === "(" || token.text === ")") {
          out.push({ type: "paren", value: token.text });
          break;
        }
        // Support locale-specific argument separators (`;` in many locales) by treating them
        // equivalently to commas. This evaluator is used in UI previews where locale-aware
        // parsing should avoid returning misleading errors.
        if (token.text === "," || token.text === ";") {
          out.push({ type: "comma", value: "," });
          break;
        }
        out.push({ type: "error", value: "#VALUE!" });
        break;
      default:
        out.push({ type: "error", value: "#VALUE!" });
        break;
    }
  }

  return out;
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

function normalizeFunctionName(name: string): string {
  const upper = String(name ?? "").toUpperCase();
  const PREFIX = "_XLFN.";
  if (upper.startsWith(PREFIX)) return upper.slice(PREFIX.length);
  return upper;
}

function workbookDirForExcel(dir: string): string {
  const raw = String(dir ?? "");
  if (!raw) return "";
  if (raw.endsWith("/") || raw.endsWith("\\")) return raw;

  const lastSlash = raw.lastIndexOf("/");
  const lastBackslash = raw.lastIndexOf("\\");
  const sep =
    lastSlash >= 0 || lastBackslash >= 0
      ? lastSlash > lastBackslash
        ? "/"
        : "\\"
      : "/";
  return `${raw}${sep}`;
}

function evalFunction(
  name: string,
  args: CellValue[],
  getCellValue: GetCellValue,
  options: EvaluateFormulaOptions,
  context: EvalContext,
  argProvenance?: AiFunctionArgumentProvenance[],
): CellValue {
  const upper = normalizeFunctionName(name);
  if (upper === "AI" || upper === "AI.EXTRACT" || upper === "AI.CLASSIFY" || upper === "AI.TRANSLATE") {
    return options.ai
      ? options.ai.evaluateAiFunction({ name: upper, args, cellAddress: options.cellAddress, argProvenance })
      : "#NAME?";
  }

  if (upper === "CELL") {
    const infoTypeArg = args[0] ?? null;
    const infoTypeScalar = unwrapProvenance(infoTypeArg);
    if (Array.isArray(infoTypeScalar)) return "#VALUE!";
    const infoType = typeof infoTypeScalar === "string" ? infoTypeScalar.trim().toLowerCase() : "";
    if (!infoType) return "#VALUE!";
    if (infoType !== "filename") return "#VALUE!";

    const meta = options.workbookFileMetadata ?? null;
    const filename = typeof meta?.filename === "string" ? meta.filename.trim() : "";
    if (!filename) return "";

    const dirRaw = typeof meta?.directory === "string" ? meta.directory : null;
    const dir = dirRaw != null && dirRaw.trim() !== "" ? workbookDirForExcel(dirRaw) : "";

    const currentSheetName = (() => {
      const raw = typeof options.currentSheetName === "string" ? options.currentSheetName.trim() : "";
      if (raw) return raw;
      const inferred = sheetIdFromCellAddress(options.cellAddress);
      return inferred ? inferred.trim() : "";
    })();

    const defaultSheetToken = sheetIdFromCellAddress(options.cellAddress)?.trim() ?? "";
    const refSheetName = (() => {
      const refArg = args[1] ?? null;
      if (refArg == null) return currentSheetName;

      const refs = provenanceRefs(refArg);
      const firstRef = refs[0] ?? "";
      if (!firstRef) return currentSheetName;
      const split = splitSheetQualifier(firstRef);
      const sheetToken = split.sheetName;
      if (!sheetToken) return currentSheetName;
      // `readReference` uses `cellAddress` as the default sheet for unqualified references,
      // but SpreadsheetApp may pass a stable internal sheet id there. When the reference
      // resolves to that same token, treat it as the current sheet and use the caller-supplied
      // display name (Excel output uses display names, not stable ids).
      if (defaultSheetToken && sheetToken.toLowerCase() === defaultSheetToken.toLowerCase()) {
        return currentSheetName;
      }
      return sheetToken;
    })();

    const value = dir ? `${dir}[${filename}]${refSheetName}` : `[${filename}]${refSheetName}`;
    if (!context.preserveReferenceProvenance) return value;
    const refs = args.flatMap((arg) => provenanceRefs(arg));
    return wrapWithProvenance(value, refs);
  }

  if (upper === "INFO") {
    const typeTextArg = args[0] ?? null;
    const typeTextScalar = unwrapProvenance(typeTextArg);
    if (Array.isArray(typeTextScalar)) return "#VALUE!";
    const typeText = typeof typeTextScalar === "string" ? typeTextScalar.trim().toLowerCase() : "";
    if (!typeText) return "#VALUE!";
    if (typeText !== "directory") return "#VALUE!";

    const meta = options.workbookFileMetadata ?? null;
    const filename = typeof meta?.filename === "string" ? meta.filename.trim() : "";
    const dirRaw = typeof meta?.directory === "string" ? meta.directory : null;
    const dir = dirRaw != null && dirRaw.trim() !== "" ? workbookDirForExcel(dirRaw) : "";
    // Match the Rust engine's Excel semantics: `INFO("directory")` returns `#N/A` until the workbook
    // has a known directory (typically after first save). Filename-only metadata (web) also yields
    // `#N/A`.
    if (!filename || !dir) return "#N/A";

    if (!context.preserveReferenceProvenance) return dir;
    const refs = args.flatMap((arg) => provenanceRefs(arg));
    return wrapWithProvenance(dir, refs);
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

  let effectiveSheetName = sheetName;
  let effectiveRef = ref;
  let range = parseA1Range(effectiveRef);

  if (!range && typeof options.resolveStructuredRefToReference === "function") {
    const resolved = options.resolveStructuredRefToReference(refText);
    if (resolved) {
      const split = splitSheetQualifier(resolved);
      effectiveSheetName = split.sheetName;
      effectiveRef = split.ref;
      range = parseA1Range(effectiveRef);
    }
  }

  if (!range) return "#REF!";

  const defaultSheet = sheetIdFromCellAddress(options.cellAddress);
  const provenanceSheet = effectiveSheetName ?? defaultSheet;
  const getPrefix = effectiveSheetName ? `${effectiveSheetName}!` : "";

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
  const upper = normalizeFunctionName(name);
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
    const upperName = normalizeFunctionName(name);
    const argContext =
      isAiFunctionName(upperName)
        ? {
            preserveReferenceProvenance: true,
            sampleRangeReferences: true,
            maxRangeCells: clampInt(
              options.aiRangeSampleLimit ?? options.ai?.rangeSampleLimit ?? DEFAULT_AI_RANGE_SAMPLE_LIMIT,
              { min: 1, max: 10_000 },
            ),
          }
        : upperName === "CELL"
          ? {
              // `CELL(info_type, [reference])` needs access to the reference's sheet name
              // (and should not materialize large range values just to capture provenance).
              preserveReferenceProvenance: true,
              sampleRangeReferences: true,
              maxRangeCells: 1,
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
