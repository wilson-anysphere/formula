import {
  parsePartialFormula as parsePartialFormulaFallback,
  FunctionRegistry,
  type PartialFormulaContext,
} from "@formula/ai-completion";
import { getLocale } from "../../i18n/index.js";

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
};

type EngineClientLike = {
  parseFormulaPartial: (
    formula: string,
    cursor?: number,
    options?: { localeId?: string },
    rpcOptions?: { timeoutMs?: number },
  ) => Promise<{ context?: { function?: { name: string; argIndex: number } | null } | null }>;
};

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

const FUNCTION_TRANSLATIONS_BY_LOCALE: Record<string, FunctionTranslationTables> = {
  "de-DE": parseFunctionTranslationsTsv(DE_DE_FUNCTION_TSV),
  "fr-FR": parseFunctionTranslationsTsv(FR_FR_FUNCTION_TSV),
  "es-ES": parseFunctionTranslationsTsv(ES_ES_FUNCTION_TSV),
};

type FormulaArgSeparator = "," | ";";

function getLocaleArgSeparator(localeId: string): FormulaArgSeparator {
  // Keep in sync with `crates/formula-engine/src/ast.rs` (`LocaleConfig::*`).
  //
  // The WASM engine currently only ships these locales. Their separators match Excel:
  // - en-US: `,` args + `.` decimals
  // - de-DE/fr-FR/es-ES: `;` args + `,` decimals
  //
  // Treat unknown locales as canonical `,` separator.
  switch (localeId) {
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

  const localeMap = FUNCTION_TRANSLATIONS_BY_LOCALE[localeId];
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
  constructor() {
    super();

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

    const localeId = (() => {
      try {
        return getLocale();
      } catch {
        return "en-US";
      }
    })();

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
        if (meta.__formulaLocaleId === localeId) {
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

export function createLocaleAwareFunctionRegistry(): FunctionRegistry {
  return new LocaleAwareFunctionRegistry();
}

const CANONICAL_STARTER_FUNCTIONS = ["SUM(", "AVERAGE(", "IF(", "XLOOKUP(", "VLOOKUP(", "INDEX(", "MATCH("];

function localizeFunctionNameForLocale(canonicalName: string, localeId: string): string {
  const tables = FUNCTION_TRANSLATIONS_BY_LOCALE[localeId];
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

/**
 * Returns a starter-function list that matches the current UI locale when translation tables are available.
 *
 * This is intended for tab completion when the user has typed only "=" and expects localized function names
 * (e.g. de-DE `=SUMME(`) instead of canonical English (`=SUM(`).
 */
export function createLocaleAwareStarterFunctions(): () => string[] {
  return () => {
    const localeId = (() => {
      try {
        return getLocale();
      } catch {
        return "en-US";
      }
    })();
    const tables = FUNCTION_TRANSLATIONS_BY_LOCALE[localeId];
    if (!tables) return CANONICAL_STARTER_FUNCTIONS;
    return CANONICAL_STARTER_FUNCTIONS.map((s) => localizeStarter(s, localeId));
  };
}

function findOpenParenIndex(prefix: string): number | null {
  // Port of `packages/ai-completion/src/formulaPartialParser.js` open-paren scan.
  /** @type {number[]} */
  const openParens: number[] = [];
  let inString = false;
  let inSheetQuote = false;
  let bracketDepth = 0;
  let braceDepth = 0;
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
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === "[") {
      bracketDepth += 1;
      continue;
    }
    if (ch === "]") {
      bracketDepth = Math.max(0, bracketDepth - 1);
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
    if (ch === "'" && bracketDepth === 0) {
      inSheetQuote = true;
      continue;
    }
    if (bracketDepth !== 0) continue;
    if (ch === "(") {
      openParens.push(i);
    } else if (ch === ")") {
      openParens.pop();
    }
  }

  if (openParens.length === 0) return null;
  return openParens[openParens.length - 1] ?? null;
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
  let bracketDepth = 0;
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
    if (ch === "'" && bracketDepth === 0) {
      inSheetQuote = true;
      continue;
    }
    if (ch === "[") {
      bracketDepth += 1;
      continue;
    }
    if (ch === "]") {
      bracketDepth = Math.max(0, bracketDepth - 1);
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
    if (ch === "(" && bracketDepth === 0) depth++;
    else if (ch === ")" && bracketDepth === 0) depth = Math.max(baseDepth, depth - 1);
    else if (depth === baseDepth && bracketDepth === 0 && braceDepth === 0) {
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
  const timeoutMs = Number.isFinite(options.timeoutMs) ? Math.max(1, Math.trunc(options.timeoutMs as number)) : 15;
  const unsupportedLocaleIds = new Set<string>();

  return (input: string, cursorPosition: number, functionRegistry: RangeArgRegistry) => {
    const cursor = clampCursor(input, cursorPosition);
    const prefix = input.slice(0, cursor);
    if (!prefix.startsWith("=")) {
      return { isFormula: false, inFunctionCall: false };
    }

    const localeId = (() => {
      try {
        return getLocale();
      } catch {
        return "en-US";
      }
    })();

    // Use the built-in JS parser for spans (currentArg.start/end, etc.) and then
    // optionally refine function/arg context via the locale-aware WASM engine.
    let baseline: PartialFormulaContext;
    try {
      baseline = parsePartialFormulaFallback(input, cursor, functionRegistry);
    } catch {
      baseline = { isFormula: true, inFunctionCall: false };
    }
    const localeArgSeparator = getLocaleArgSeparator(localeId);
    if (localeArgSeparator === ";") {
      // The fallback parser is intentionally locale-agnostic, and will treat `,` as an argument
      // separator until a `;` appears. In semicolon locales (where `,` is the decimal separator),
      // that can mis-classify partial numbers like `1,` as "arg 1". Fix up the arg span/index
      // using the known locale separator so completion edits remain stable.
      const openParenIndex = findOpenParenIndex(prefix);
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
    baseline = canonicalizeInCallContext(baseline, localeId, functionRegistry);

    const engine = (() => {
      try {
        return getEngineClient();
      } catch {
        return null;
      }
    })();

    if (!engine) return baseline;
    if (unsupportedLocaleIds.has(localeId)) return baseline;

    const parsePromise = Promise.resolve()
      .then(() => engine.parseFormulaPartial(input, cursor, { localeId }, { timeoutMs }))
      .catch((err) => {
        const message = err instanceof Error ? err.message : String(err);
        if (typeof message === "string" && message.startsWith("unknown localeId:")) {
          unsupportedLocaleIds.add(localeId);
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

        const canonicalFnName = canonicalizeFunctionNameForLocale(rawName, localeId);

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
