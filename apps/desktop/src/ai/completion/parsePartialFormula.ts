import {
  parsePartialFormula as parsePartialFormulaFallback,
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

function toAsciiUpperCase(str: string): string {
  // Mirror Rust's `to_ascii_uppercase` to keep comparisons stable for non-ASCII identifiers.
  return str.replace(/[a-z]/g, (ch) => ch.toUpperCase());
}

type FunctionTranslationMap = Map<string, string>;

function parseFunctionTranslationsTsv(tsv: string): FunctionTranslationMap {
  const map: FunctionTranslationMap = new Map();
  for (const rawLine of String(tsv ?? "").split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const [canonical, localized] = line.split("\t");
    if (!canonical || !localized) continue;
    const canonUpper = toAsciiUpperCase(canonical.trim());
    const locUpper = toAsciiUpperCase(localized.trim());
    // Only store translations that differ; identity entries can fall back to `to_ascii_uppercase`.
    if (canonUpper && locUpper && canonUpper !== locUpper) {
      map.set(locUpper, canonUpper);
    }
  }
  return map;
}

const FUNCTION_TRANSLATIONS_BY_LOCALE: Record<string, FunctionTranslationMap> = {
  "de-DE": parseFunctionTranslationsTsv(DE_DE_FUNCTION_TSV),
  "fr-FR": parseFunctionTranslationsTsv(FR_FR_FUNCTION_TSV),
  "es-ES": parseFunctionTranslationsTsv(ES_ES_FUNCTION_TSV),
};

function canonicalizeFunctionNameForLocale(name: string, localeId: string): string {
  const raw = String(name ?? "");
  if (!raw) return raw;

  const localeMap = FUNCTION_TRANSLATIONS_BY_LOCALE[localeId];
  if (!localeMap) return toAsciiUpperCase(raw);

  // Mirror `formula_engine::locale::registry::FormulaLocale::canonical_function_name`.
  const PREFIX = "_xlfn.";
  const hasPrefix = raw.length >= PREFIX.length && raw.slice(0, PREFIX.length).toLowerCase() === PREFIX;
  const base = hasPrefix ? raw.slice(PREFIX.length) : raw;
  const upper = toAsciiUpperCase(base);
  const mapped = localeMap.get(upper) ?? upper;
  return hasPrefix ? `${PREFIX}${mapped}` : mapped;
}

function canonicalizeInCallContext(
  ctx: PartialFormulaContext,
  localeId: string,
  functionRegistry: RangeArgRegistry,
): PartialFormulaContext {
  if (!ctx?.isFormula || !ctx.inFunctionCall || typeof ctx.functionName !== "string") return ctx;
  const argIndex = Number.isInteger(ctx.argIndex) ? (ctx.argIndex as number) : 0;
  const canonicalFnName = canonicalizeFunctionNameForLocale(ctx.functionName, localeId);
  if (!canonicalFnName || canonicalFnName === ctx.functionName) return ctx;
  return {
    ...ctx,
    functionName: canonicalFnName,
    expectingRange: Boolean(functionRegistry?.isRangeArg?.(canonicalFnName, argIndex)),
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

