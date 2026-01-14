import {
  parsePartialFormula as parsePartialFormulaFallback,
  FunctionRegistry,
  type PartialFormulaContext,
} from "@formula/ai-completion";
import { getLocale } from "../../i18n/index.js";
import { normalizeFormulaLocaleId, type FormulaLocaleId } from "../../spreadsheet/formulaLocale.js";

// Translation tables from the Rust engine (canonical <-> localized function names).
// Keep these in sync with `crates/formula-engine/src/locale/data/*.tsv`.
//
// We only need the localized->canonical direction so the completion engine can
// look up signatures/range-arg metadata against the canonical function registry.
import DE_DE_FUNCTION_TSV from "../../../../../crates/formula-engine/src/locale/data/de-DE.tsv?raw";
import ES_ES_FUNCTION_TSV from "../../../../../crates/formula-engine/src/locale/data/es-ES.tsv?raw";
import FR_FR_FUNCTION_TSV from "../../../../../crates/formula-engine/src/locale/data/fr-FR.tsv?raw";

type RangeArgRegistry = {
  isRangeArg: (fnName: string, argIndex: number) => boolean;
  // Optional: some registry implementations (e.g. FunctionRegistry) provide `search()` which we can
  // use to disambiguate A1-looking function names like "LOG10" from cell references like "A1".
  search?: (prefix: string, options?: { limit?: number }) => any[];
};

type EngineClientLike = {
  parseFormulaPartial: (
    formula: string,
    cursor?: number,
    options?: { localeId?: string },
    rpcOptions?: { timeoutMs?: number },
  ) => Promise<{ context?: { function?: { name: string; argIndex: number } | null } | null }>;
};

function currentLocaleId(): string {
  try {
    const raw = typeof document !== "undefined" ? document.documentElement?.lang : "";
    const trimmed = String(raw ?? "").trim();
    if (trimmed) return trimmed;
  } catch {
    // ignore
  }
  try {
    return getLocale();
  } catch {
    return "en-US";
  }
}

function safeLocaleId(getLocaleId: () => string): string {
  try {
    const raw = getLocaleId();
    const trimmed = String(raw ?? "").trim();
    return trimmed || "en-US";
  } catch {
    return "en-US";
  }
}

function clampCursor(input: string, cursorPosition: number): number {
  const len = typeof input === "string" ? input.length : 0;
  if (!Number.isInteger(cursorPosition)) return len;
  if (cursorPosition < 0) return 0;
  if (cursorPosition > len) return len;
  return cursorPosition;
}

function casefoldIdent(ident: string): string {
  // Mirror Rust's locale behavior (`casefold_ident` / `casefold`):
  // - ASCII identifiers: ASCII uppercase (Excel-style case-insensitive)
  // - Non-ASCII identifiers: Unicode-aware uppercasing (`ä` -> `Ä`, `ß` -> `SS`, ...)
  //
  // JS `toUpperCase()` performs Unicode uppercasing and is stable enough for our use here.
  return String(ident ?? "").toUpperCase();
}

type FunctionTranslationMap = Map<string, string>;

type FunctionTranslationTables = {
  localizedToCanonical: FunctionTranslationMap;
  canonicalToLocalized: Map<string, string>;
};

function parseFunctionTranslationsTsv(tsv: string): FunctionTranslationTables {
  const localizedToCanonical: FunctionTranslationMap = new Map();
  const canonicalToLocalized: Map<string, string> = new Map();
  for (const rawLine of String(tsv ?? "").split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const [canonical, localized] = line.split("\t");
    if (!canonical || !localized) continue;
    const canonUpper = casefoldIdent(canonical.trim());
    const localizedTrimmed = localized.trim();
    const locUpper = casefoldIdent(localizedTrimmed);
    // Only store translations that differ; identity entries can fall back to `to_ascii_uppercase`.
    if (canonUpper && locUpper && canonUpper !== locUpper) {
      localizedToCanonical.set(locUpper, canonUpper);
      // Preserve the exact localized spelling from the TSV (typically already uppercase), so
      // completion suggestions can emit it without casefold expansions (e.g. ß -> SS).
      canonicalToLocalized.set(canonUpper, localizedTrimmed);
    }
  }
  return { localizedToCanonical, canonicalToLocalized };
}

const FUNCTION_TRANSLATIONS_BY_LOCALE: Record<FormulaLocaleId, FunctionTranslationTables> = {
  "de-DE": parseFunctionTranslationsTsv(DE_DE_FUNCTION_TSV),
  "fr-FR": parseFunctionTranslationsTsv(FR_FR_FUNCTION_TSV),
  "es-ES": parseFunctionTranslationsTsv(ES_ES_FUNCTION_TSV),
  // Minimal locales (no translated function names yet, but they are valid engine locale ids).
  "ja-JP": { localizedToCanonical: new Map(), canonicalToLocalized: new Map() },
  "zh-CN": { localizedToCanonical: new Map(), canonicalToLocalized: new Map() },
  "zh-TW": { localizedToCanonical: new Map(), canonicalToLocalized: new Map() },
  "ko-KR": { localizedToCanonical: new Map(), canonicalToLocalized: new Map() },
  "en-US": { localizedToCanonical: new Map(), canonicalToLocalized: new Map() },
};

type FormulaArgSeparator = "," | ";";

function getLocaleArgSeparator(localeId: string): FormulaArgSeparator {
  // Keep in sync with `crates/formula-engine/src/ast.rs` (`LocaleConfig::*`).
  //
  // The WASM engine currently only ships these locales. Their separators match Excel:
  // - en-US: `,` args + `.` decimals
  // - de-DE/fr-FR/es-ES: `;` args + `,` decimals
  // - other engine locales (ja/zh/ko): currently share en-US punctuation
  //
  // Treat unknown locales as canonical `,` separator.
  switch (normalizeFormulaLocaleId(localeId)) {
    case "de-DE":
    case "fr-FR":
    case "es-ES":
      return ";";
    default:
      return ",";
  }
}

function canonicalizeFunctionNameForLocale(name: string, localeId: string): string {
  const raw = String(name ?? "");
  if (!raw) return raw;

  const effectiveLocaleId = normalizeFormulaLocaleId(localeId) ?? "en-US";
  const localeMap = FUNCTION_TRANSLATIONS_BY_LOCALE[effectiveLocaleId];
  if (!localeMap) return casefoldIdent(raw);

  // Mirror `formula_engine::locale::registry::FormulaLocale::canonical_function_name`.
  const PREFIX = "_xlfn.";
  const hasPrefix = raw.length >= PREFIX.length && raw.slice(0, PREFIX.length).toLowerCase() === PREFIX;
  const base = hasPrefix ? raw.slice(PREFIX.length) : raw;
  const upper = casefoldIdent(base);
  const mapped = localeMap.localizedToCanonical.get(upper) ?? upper;
  return hasPrefix ? `${PREFIX}${mapped}` : mapped;
}

type LocaleAliasMeta = {
  __formulaLocaleId?: string;
  __formulaCanonicalName?: string;
};

/**
 * Desktop-only `FunctionRegistry` that prefers localized function-name completions when
 * a locale translation table is available.
 *
 * It registers localized aliases (e.g. `SUMME`) alongside canonical names (`SUM`) and
 * filters `search()` results so:
 * - callers still get canonical matches when no localized alias matches the prefix
 * - when a localized alias *does* match, its canonical counterpart is suppressed so the
 *   localized name wins the completion ranking (important for pure-insertion UX)
 */
class LocaleAwareFunctionRegistry extends FunctionRegistry {
  readonly #getLocaleId: () => string;

  constructor(options: { getLocaleId?: () => string } = {}) {
    super();
    this.#getLocaleId = options.getLocaleId ?? currentLocaleId;

    // Register localized aliases for any locales the WASM engine supports. The desktop UI
    // may not expose all of these locales yet, but registering them eagerly keeps this
    // registry future-proof and allows tests to validate the mapping tables.
    for (const [localeId, tables] of Object.entries(FUNCTION_TRANSLATIONS_BY_LOCALE)) {
      for (const [canonUpper, localizedName] of tables.canonicalToLocalized.entries()) {
        // Avoid overriding canonical entries or existing aliases (collisions are possible across locales).
        if (this.getFunction(localizedName)) continue;

        const spec = this.getFunction(canonUpper);
        if (!spec) continue;

        const alias: any = {
          ...spec,
          name: localizedName,
          // Small boost so localized aliases win ranking when both a localized alias (e.g. SUMME)
          // and other canonical functions (e.g. SUMIF) match the same short prefix ("SU").
          completionBoost: 0.05,
          __formulaLocaleId: localeId,
          __formulaCanonicalName: spec.name,
        } satisfies LocaleAliasMeta;
        try {
          this.register(alias);
        } catch {
          // Be defensive: localized alias registration is best-effort and should never
          // prevent the completion engine from booting.
        }
      }
    }
  }

  search(prefix: string, options?: { limit?: number }): any[] {
    const rawLimit = options?.limit ?? 10;
    // Fetch a larger candidate set, then filter down by locale. This ensures we still return
    // enough localized results even when the prefix matches many canonical names.
    //
    // Note: because we register aliases for multiple locales, `super.search()` can return many
    // "wrong-locale" aliases before reaching the canonical functions. Keep this large enough to
    // avoid accidentally filtering everything out for some prefixes in other locales.
    const candidateLimit = Math.max(rawLimit * 50, 500);

    const localeId = safeLocaleId(this.#getLocaleId);
    const formulaLocaleId = normalizeFormulaLocaleId(localeId) ?? "en-US";

    const candidates: any[] = super.search(prefix, { ...options, limit: candidateLimit } as any);

    /** @type {any[]} */
    const localized: any[] = [];
    /** @type {any[]} */
    const canonical: any[] = [];
    /** @type {Set<string>} */
    const suppressCanonical = new Set<string>();

    for (const spec of candidates) {
      const meta = spec as any as LocaleAliasMeta;
      const isAlias = typeof meta.__formulaLocaleId === "string" && meta.__formulaLocaleId.length > 0;
      if (isAlias) {
        if (meta.__formulaLocaleId === formulaLocaleId) {
          localized.push(spec);
          if (typeof meta.__formulaCanonicalName === "string") {
            suppressCanonical.add(casefoldIdent(meta.__formulaCanonicalName));
          }
        }
        continue;
      }
      canonical.push(spec);
    }

    // Dedupe canonical specs that have localized aliases for this locale, then merge localized first
    // so the UI can prefer localized function names while still surfacing untranslated functions.
    const canonicalFiltered = canonical.filter((spec) => !suppressCanonical.has(casefoldIdent(spec?.name)));
    return [...localized, ...canonicalFiltered].slice(0, rawLimit);
  }
}

export function createLocaleAwareFunctionRegistry(options: { getLocaleId?: () => string } = {}): FunctionRegistry {
  return new LocaleAwareFunctionRegistry(options);
}

const CANONICAL_STARTER_FUNCTIONS = ["SUM(", "AVERAGE(", "IF(", "XLOOKUP(", "VLOOKUP(", "INDEX(", "MATCH("];

function localizeFunctionNameForLocale(canonicalName: string, localeId: string): string {
  const tables = FUNCTION_TRANSLATIONS_BY_LOCALE[normalizeFormulaLocaleId(localeId) ?? "en-US"];
  if (!tables) return canonicalName;
  const upper = casefoldIdent(canonicalName);
  return tables.canonicalToLocalized.get(upper) ?? canonicalName;
}

function localizeStarter(starter: string, localeId: string): string {
  const raw = String(starter ?? "");
  const idx = raw.indexOf("(");
  if (idx < 0) return raw;
  const name = raw.slice(0, idx);
  const suffix = raw.slice(idx);
  const localizedName = localizeFunctionNameForLocale(name, localeId);
  return `${localizedName}${suffix}`;
}

const UNICODE_LETTER_RE: RegExp | null = (() => {
  try {
    return new RegExp("^\\p{Alphabetic}$", "u");
  } catch {
    // Older JS engines may not support Unicode property escapes.
    return null;
  }
})();

const UNICODE_ALNUM_RE: RegExp | null = (() => {
  try {
    // Match the Rust backend's `char::is_alphanumeric` (Alphabetic || Number).
    return new RegExp("^[\\p{Alphabetic}\\p{Number}]$", "u");
  } catch {
    // Older JS engines may not support Unicode property escapes.
    return null;
  }
})();

function isUnicodeAlphabetic(ch: string): boolean {
  if (UNICODE_LETTER_RE) return UNICODE_LETTER_RE.test(ch);
  return (ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z");
}

function isUnicodeAlphanumeric(ch: string): boolean {
  if (UNICODE_ALNUM_RE) return UNICODE_ALNUM_RE.test(ch);
  return isUnicodeAlphabetic(ch) || (ch >= "0" && ch <= "9");
}

/**
 * True when `ch` is a valid identifier character for function names / name prefixes.
 *
 * This intentionally matches `packages/ai-completion/src/formulaPartialParser.js` so the desktop
 * tab-completion behavior stays aligned with the JS fallback parser.
 */
function isIdentChar(ch: string): boolean {
  if (!ch) return false;
  // Fast path for ASCII.
  const code = ch.charCodeAt(0);
  if (
    (code >= 48 && code <= 57) || // 0-9
    (code >= 65 && code <= 90) || // A-Z
    (code >= 97 && code <= 122) || // a-z
    code === 46 || // .
    code === 95 // _
  ) {
    return true;
  }
  // Best-effort Unicode support for localized function names (e.g. ZÄHLENWENN).
  return Boolean(UNICODE_ALNUM_RE && UNICODE_ALNUM_RE.test(ch));
}

/**
 * Best-effort check: does the host function registry contain any functions that start with
 * the given prefix?
 *
 * This is used to disambiguate a small set of Excel functions that look like A1 cell references
 * (e.g. `LOG10` looks like column `LOG`, row `10`).
 */
function hasFunctionPrefix(functionRegistry: unknown, prefix: string): boolean {
  const search =
    functionRegistry && typeof (functionRegistry as any).search === "function" ? (functionRegistry as any).search : null;
  if (!search) return false;
  try {
    const matches = search.call(functionRegistry, prefix, { limit: 1 });
    return Array.isArray(matches) && matches.length > 0;
  } catch {
    return false;
  }
}

function functionNameFromIdent(identToken: string | null, functionRegistry: unknown): string | null {
  const token = typeof identToken === "string" ? identToken : "";
  if (!token) return null;
  // Avoid returning something that is obviously a cell ref like "A1".
  if (/^[A-Za-z]{1,3}\d+$/.test(token) && !hasFunctionPrefix(functionRegistry, token)) return null;
  return casefoldIdent(token);
}

/**
 * Returns a starter-function list that matches the current UI locale when translation tables are available.
 *
 * This is intended for tab completion when the user has typed only "=" and expects localized function names
 * (e.g. de-DE `=SUMME(`) instead of canonical English (`=SUM(`).
 */
export function createLocaleAwareStarterFunctions(options: { getLocaleId?: () => string } = {}): () => string[] {
  const getLocaleId = options.getLocaleId ?? currentLocaleId;
  return () => {
    const localeId = safeLocaleId(getLocaleId);
    const normalizedLocaleId = normalizeFormulaLocaleId(localeId) ?? "en-US";
    const tables = FUNCTION_TRANSLATIONS_BY_LOCALE[normalizedLocaleId];
    if (!tables) return CANONICAL_STARTER_FUNCTIONS;
    return CANONICAL_STARTER_FUNCTIONS.map((s) => localizeStarter(s, normalizedLocaleId));
  };
}

function findOpenParenIndex(prefix: string, functionRegistry: unknown): number | null {
  // Port of `packages/ai-completion/src/formulaPartialParser.js` open-paren scan.
  /** @type {{ index: number; functionName: string | null }[]} */
  const openParens: Array<{ index: number; functionName: string | null }> = [];
  let inString = false;
  let inSheetQuote = false;
  // Track whether the cursor is currently inside a `[...]` segment.
  //
  // Note: In Excel formulas, `]` inside structured references and external workbook prefixes
  // is escaped as `]]`, which is ambiguous with nested bracket closure (e.g. `[[Col]]`).
  // Use `findMatchingBracketEnd` to skip complete bracket segments and avoid naive depth
  // counting errors that would treat `]]` as two closings.
  let bracketDepth = 0;
  let braceDepth = 0;
  // Track the most recent identifier token so we can cheaply associate it with a following '('
  // (function call). This avoids O(n^2) rescans for formulas with many nested/grouping parens.
  let identStart: number | null = null;
  let pendingIdent: string | null = null;
  for (let i = 0; i < prefix.length; i++) {
    const ch = prefix[i];
    if (inString) {
      if (ch === '"') {
        // Excel escapes quotes inside string literals via doubled quotes: "".
        if (prefix[i + 1] === '"') {
          i += 1;
          continue;
        }
        inString = false;
      }
      continue;
    }
    if (inSheetQuote) {
      if (ch === "'") {
        // Excel escapes apostrophes inside sheet names via doubled quotes: ''.
        if (prefix[i + 1] === "'") {
          i += 1;
          continue;
        }
        inSheetQuote = false;
      }
      continue;
    }
    // Only track identifiers outside structured references. Identifiers inside `[...]` are
    // table/column names and shouldn't be considered function names.
    if (bracketDepth === 0 && isIdentChar(ch!)) {
      if (identStart === null) identStart = i;
      continue;
    }
    if (identStart !== null) {
      pendingIdent = prefix.slice(identStart, i);
      identStart = null;
    }
    if (ch === '"') {
      inString = true;
      pendingIdent = null;
      continue;
    }
    if (ch === "[") {
      pendingIdent = null;
      const end = findMatchingBracketEnd(prefix, i, prefix.length);
      if (end == null) {
        bracketDepth = 1;
        break;
      }
      i = end - 1;
      pendingIdent = null;
      continue;
    }
    if (ch === "{") {
      braceDepth += 1;
      pendingIdent = null;
      continue;
    }
    if (ch === "}") {
      braceDepth = Math.max(0, braceDepth - 1);
      pendingIdent = null;
      continue;
    }
    if (ch === "'" && bracketDepth === 0) {
      inSheetQuote = true;
      pendingIdent = null;
      continue;
    }
    if (bracketDepth !== 0) continue;
    if (ch === "(") {
      openParens.push({ index: i, functionName: functionNameFromIdent(pendingIdent, functionRegistry) });
      pendingIdent = null;
    } else if (ch === ")") {
      openParens.pop();
      pendingIdent = null;
    } else if (!/\s/.test(ch!)) {
      // Any other non-whitespace token breaks the identifier->'(' link.
      pendingIdent = null;
    }
  }

  for (let i = openParens.length - 1; i >= 0; i--) {
    const fnName = openParens[i]?.functionName;
    if (typeof fnName === "string" && fnName.length > 0) return openParens[i]!.index;
  }
  return null;
}

function getArgContextWithSeparator(
  prefix: string,
  openParenIndex: number,
  cursorPosition: number,
  argSeparator: FormulaArgSeparator,
): { argIndex: number; currentArg: { text: string; start: number; end: number } } {
  // Port of `packages/ai-completion/src/formulaPartialParser.js` arg scanner, but with a
  // caller-provided `argSeparator` so locale-aware callers can avoid mis-parsing decimal commas.
  const baseDepth = 1;
  let depth = baseDepth;
  let argIndex = 0;
  let lastArgSeparatorIndex = -1;
  let inString = false;
  let inSheetQuote = false;
  let braceDepth = 0;

  for (let i = openParenIndex + 1; i < cursorPosition; i++) {
    const ch = prefix[i];
    if (inString) {
      if (ch === '"') {
        if (prefix[i + 1] === '"') {
          i += 1;
          continue;
        }
        inString = false;
      }
      continue;
    }
    if (inSheetQuote) {
      if (ch === "'") {
        if (prefix[i + 1] === "'") {
          i += 1;
          continue;
        }
        inSheetQuote = false;
      }
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === "'") {
      inSheetQuote = true;
      continue;
    }
    if (ch === "[") {
      const end = findMatchingBracketEnd(prefix, i, cursorPosition);
      if (end == null) break;
      i = end - 1;
      continue;
    }
    if (ch === "{") {
      braceDepth += 1;
      continue;
    }
    if (ch === "}") {
      braceDepth = Math.max(0, braceDepth - 1);
      continue;
    }
    if (ch === "(") depth++;
    else if (ch === ")") depth = Math.max(baseDepth, depth - 1);
    else if (depth === baseDepth && braceDepth === 0) {
      if (ch === argSeparator) {
        argIndex += 1;
        lastArgSeparatorIndex = i;
      }
    }
  }

  let rawStart = lastArgSeparatorIndex === -1 ? openParenIndex + 1 : lastArgSeparatorIndex + 1;
  let start = rawStart;
  while (start < cursorPosition && /\s/.test(prefix[start])) start++;
  const currentArg = {
    start,
    end: cursorPosition,
    text: prefix.slice(start, cursorPosition),
  };

  return { argIndex, currentArg };
}

function findWorkbookPrefixEnd(src: string, startIndex: number, max: number): number | null {
  // External workbook prefixes escape closing brackets by doubling: `]]` -> literal `]`.
  //
  // Workbook names may also contain `[` characters; treat them as plain text (no nesting).
  if (src[startIndex] !== "[") return null;
  let i = startIndex + 1;
  while (i < max) {
    if (src[i] === "]") {
      if (i + 1 < max && src[i + 1] === "]") {
        i += 2;
        continue;
      }
      return i + 1;
    }
    i += 1;
  }
  return null;
}

function findWorkbookPrefixEndIfValid(src: string, startIndex: number, max: number): number | null {
  const end = findWorkbookPrefixEnd(src, startIndex, max);
  if (!end) return null;

  // Heuristic: only treat this as an external workbook prefix if it is immediately followed by
  // an unquoted sheet name and `!` (e.g. `[Book.xlsx]Sheet1!A1`). This avoids incorrectly treating
  // incomplete structured references as workbook prefixes.
  if (end >= max) return null;
  const first = src[end] ?? "";
  if (!(first === "_" || isUnicodeAlphabetic(first))) return null;

  let i = end + 1;
  while (i < max) {
    const ch = src[i] ?? "";
    if (ch === "!") return end;
    if (ch === "_" || ch === "." || ch === ":" || isUnicodeAlphanumeric(ch)) {
      i += 1;
      continue;
    }
    break;
  }
  return null;
}

function findMatchingStructuredRefBracketEnd(src: string, startIndex: number, max: number): number | null {
  if (src[startIndex] !== "[") return null;

  let i = startIndex;
  let depth = 0;
  const escapeChoices: Array<{ i: number; depth: number }> = [];

  const backtrack = (): boolean => {
    const choice = escapeChoices.pop();
    if (!choice) return false;
    i = choice.i;
    depth = choice.depth;
    // Reinterpret the first `]` of the `]]` pair as a real closing bracket.
    depth -= 1;
    i += 1;
    return true;
  };

  while (true) {
    if (i >= max) {
      if (!backtrack()) return null;
      continue;
    }

    const ch = src[i] ?? "";
    if (ch === "[") {
      depth += 1;
      i += 1;
      continue;
    }
    if (ch === "]") {
      if (src[i + 1] === "]" && depth > 0 && i + 1 < max) {
        // Prefer treating `]]` as an escaped literal `]` inside the bracket segment.
        escapeChoices.push({ i, depth });
        i += 2;
        continue;
      }
      depth -= 1;
      i += 1;
      if (depth === 0) return i;
      if (depth < 0) {
        // Too many closing brackets - try reinterpreting an earlier escape.
        if (!backtrack()) return null;
      }
      continue;
    }

    i += 1;
  }
}

function findMatchingBracketEnd(src: string, startIndex: number, limit: number): number | null {
  const max = Number.isFinite(limit) ? Math.max(0, Math.min(src.length, Math.trunc(limit))) : src.length;
  if (startIndex < 0 || startIndex >= max) return null;
  if (src[startIndex] !== "[") return null;

  return findMatchingStructuredRefBracketEnd(src, startIndex, max) ?? findWorkbookPrefixEndIfValid(src, startIndex, max);
}

function canonicalizeInCallContext(
  ctx: PartialFormulaContext,
  localeId: string,
  functionRegistry: RangeArgRegistry,
): PartialFormulaContext {
  if (!ctx?.isFormula || !ctx.inFunctionCall || typeof ctx.functionName !== "string") return ctx;
  const argIndex = Number.isInteger(ctx.argIndex) ? (ctx.argIndex as number) : 0;
  const canonicalFnName = canonicalizeFunctionNameForLocale(ctx.functionName, localeId);
  if (!canonicalFnName) return ctx;

  const expectingRange = Boolean(functionRegistry?.isRangeArg?.(canonicalFnName, argIndex));
  const needsUpdate = canonicalFnName !== ctx.functionName || expectingRange !== Boolean(ctx.expectingRange);
  if (!needsUpdate) return ctx;
  return {
    ...ctx,
    functionName: canonicalFnName,
    expectingRange,
  };
}

export function createLocaleAwarePartialFormulaParser(options: {
  /**
   * Return the current EngineClient instance when available.
   *
   * The desktop app initializes the WASM engine asynchronously; this indirection allows
   * tab completion to stay responsive while the engine is still booting.
   */
  getEngineClient?: () => EngineClientLike | null;
  /**
   * Optional locale-id getter used for locale-aware parsing.
   *
   * When omitted, the parser uses the best-effort document/i18n locale (`currentLocaleId()`).
   */
  getLocaleId?: () => string;
  /**
   * Maximum time (ms) we're willing to wait for the engine parser before falling back
   * to the JS implementation.
   */
  timeoutMs?: number;
}): (
  input: string,
  cursorPosition: number,
  functionRegistry: RangeArgRegistry,
) => PartialFormulaContext | Promise<PartialFormulaContext> {
  const getEngineClient = options.getEngineClient ?? (() => null);
  const getLocaleId = options.getLocaleId ?? currentLocaleId;
  const timeoutMs = Number.isFinite(options.timeoutMs) ? Math.max(1, Math.trunc(options.timeoutMs as number)) : 15;
  const unsupportedLocaleIds = new Set<string>();

  return (input: string, cursorPosition: number, functionRegistry: RangeArgRegistry) => {
    const cursor = clampCursor(input, cursorPosition);
    const prefix = input.slice(0, cursor);
    if (!prefix.startsWith("=")) {
      return { isFormula: false, inFunctionCall: false };
    }

    const localeId = safeLocaleId(getLocaleId);
    const normalizedLocaleId = normalizeFormulaLocaleId(localeId) ?? "en-US";

    // Use the built-in JS parser for spans (currentArg.start/end, etc.) and then
    // optionally refine function/arg context via the locale-aware WASM engine.
    let baseline: PartialFormulaContext;
    try {
      baseline = parsePartialFormulaFallback(input, cursor, functionRegistry);
    } catch {
      baseline = { isFormula: true, inFunctionCall: false };
    }
    const localeArgSeparator = getLocaleArgSeparator(normalizedLocaleId);
    const shouldFixArgContext = localeArgSeparator === ";" || prefix.includes(";");
    if (shouldFixArgContext) {
      // The fallback parser is intentionally locale-agnostic:
      // - It treats `,` as an argument separator until a `;` appears.
      // - If a `;` appears, it may switch into "semicolon locale" mode.
      //
      // When we know the locale (or are falling back to en-US for unsupported locales),
      // recompute the arg span/index using the effective locale separator so completion
      // behavior matches parsing semantics.
      const openParenIndex =
        baseline.inFunctionCall && typeof (baseline as any).openParenIndex === "number"
          ? ((baseline as any).openParenIndex as number)
          : findOpenParenIndex(prefix, functionRegistry);
      if (openParenIndex != null) {
        const { argIndex, currentArg } = getArgContextWithSeparator(
          prefix,
          openParenIndex,
          cursor,
          localeArgSeparator,
        );
        baseline = { ...baseline, argIndex, currentArg };
      }
    }
    baseline = canonicalizeInCallContext(baseline, normalizedLocaleId, functionRegistry);

    const engine = (() => {
      try {
        return getEngineClient();
      } catch {
        return null;
      }
    })();

    if (!engine) return baseline;
    if (unsupportedLocaleIds.has(normalizedLocaleId)) return baseline;

    const parsePromise = Promise.resolve()
      .then(() => engine.parseFormulaPartial(input, cursor, { localeId: normalizedLocaleId }, { timeoutMs }))
      .catch((err) => {
        const message = err instanceof Error ? err.message : String(err);
        if (typeof message === "string") {
          const trimmed = message.trim();
          const prefix = "unknown localeId:";
          if (trimmed.startsWith(prefix)) {
            const unknown = trimmed.slice(prefix.length).trim();
            // Only cache when the engine is rejecting the exact locale id we passed. This keeps the
            // defensive "unsupported locale" fast-path from triggering on unrelated errors that
            // happen to share the same prefix.
            if (unknown && unknown.toLowerCase() === normalizedLocaleId.toLowerCase()) {
              unsupportedLocaleIds.add(normalizedLocaleId);
            }
          }
        }
        return null;
      });

    return withTimeout(parsePromise, timeoutMs)
      .then((result) => {
        const ctx = result?.context?.function ?? null;
        const rawName = typeof (ctx as any)?.name === "string" ? String((ctx as any).name).trim() : "";
        const argIndex =
          typeof (ctx as any)?.argIndex === "number" && Number.isInteger((ctx as any).argIndex)
            ? (ctx as any).argIndex
            : null;
        if (!rawName || argIndex == null || argIndex < 0) return baseline;

        const canonicalFnName = canonicalizeFunctionNameForLocale(rawName, normalizedLocaleId);

        return {
          ...baseline,
          inFunctionCall: true,
          functionName: canonicalFnName,
          argIndex,
          expectingRange: Boolean(functionRegistry?.isRangeArg?.(canonicalFnName, argIndex)),
        };
      })
      .catch(() => baseline);
  };
}

function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T | null> {
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) return promise.then((v) => v);
  return new Promise((resolve) => {
    let settled = false;
    const timer = setTimeout(() => {
      settled = true;
      resolve(null);
    }, timeoutMs);
    promise
      .then((value) => {
        if (settled) return;
        clearTimeout(timer);
        settled = true;
        resolve(value);
      })
      .catch(() => {
        if (settled) return;
        clearTimeout(timer);
        settled = true;
        resolve(null);
      });
  });
}
