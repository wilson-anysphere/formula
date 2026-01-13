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

type FunctionCallFrame = {
  name: string;
  parenDepth: number;
  braceDepth: number;
  bracketDepth: number;
  openParenIndex: number;
  argIndex: number;
  lastArgSepIndex: number | null;
};

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

function isAsciiLetter(ch: string): boolean {
  return ch >= "A" && ch <= "Z" ? true : ch >= "a" && ch <= "z";
}

function isAsciiDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

// Avoid Unicode-property RegExp literals (e.g. `/\p{L}/u`) so the bundle still parses
// in JS engines that don't support them. Fall back to ASCII-only heuristics there.
const UNICODE_LETTER_RE: RegExp | null = (() => {
  try {
    return new RegExp("^\\p{Alphabetic}$", "u");
  } catch {
    return null;
  }
})();

const UNICODE_ALNUM_RE: RegExp | null = (() => {
  try {
    return new RegExp("^[\\p{Alphabetic}\\p{Number}]$", "u");
  } catch {
    return null;
  }
})();

function isUnicodeAlphabetic(ch: string): boolean {
  if (UNICODE_LETTER_RE) return UNICODE_LETTER_RE.test(ch);
  return isAsciiLetter(ch);
}

function isUnicodeAlphanumeric(ch: string): boolean {
  if (UNICODE_ALNUM_RE) return UNICODE_ALNUM_RE.test(ch);
  return isAsciiLetter(ch) || isAsciiDigit(ch);
}

function clampCursor(input: string, cursorPosition: number): number {
  const len = typeof input === "string" ? input.length : 0;
  if (!Number.isInteger(cursorPosition)) return len;
  if (cursorPosition < 0) return 0;
  if (cursorPosition > len) return len;
  return cursorPosition;
}

function isIdentStartChar(ch: string): boolean {
  if (!ch) return false;
  if (ch === "$" || ch === "_" || ch === "\\") return true;
  // Unicode identifiers: mirror the Rust lexer which allows non-ASCII alphabetic.
  return isUnicodeAlphabetic(ch);
}

function isIdentContChar(ch: string): boolean {
  if (!ch) return false;
  if (ch === "$" || ch === "_" || ch === "\\" || ch === ".") return true;
  return isUnicodeAlphanumeric(ch);
}

/**
 * Best-effort scan for the current innermost function call frame in `formulaPrefix`.
 *
 * This mirrors the fallback scanner used by the Rust WASM tooling for lex errors
 * (see `crates/formula-wasm/src/lib.rs::scan_fallback_function_context`), but is
 * implemented in terms of JS UTF-16 code unit offsets.
 */
function scanFunctionCallFrame(formulaPrefix: string, argSeparator: string): FunctionCallFrame | null {
  type Mode = "normal" | "string" | "quotedIdent";
  let mode: Mode = "normal";
  let parenDepth = 0;
  let braceDepth = 0;
  let bracketDepth = 0;
  /** @type {FunctionCallFrame[]} */
  const stack: FunctionCallFrame[] = [];

  for (let i = 0; i < formulaPrefix.length; ) {
    const ch = formulaPrefix[i]!;

    if (mode === "string") {
      if (ch === '"') {
        if (formulaPrefix[i + 1] === '"') {
          i += 2;
          continue;
        }
        mode = "normal";
        i += 1;
        continue;
      }
      i += 1;
      continue;
    }

    if (mode === "quotedIdent") {
      if (ch === "'") {
        if (formulaPrefix[i + 1] === "'") {
          i += 2;
          continue;
        }
        mode = "normal";
        i += 1;
        continue;
      }
      i += 1;
      continue;
    }

    // Mode: normal
    if (bracketDepth === 0) {
      if (ch === '"') {
        mode = "string";
        i += 1;
        continue;
      }
      if (ch === "'") {
        mode = "quotedIdent";
        i += 1;
        continue;
      }
    }

    if (bracketDepth > 0) {
      // Inside structured-ref/workbook brackets, treat everything as raw text except nested brackets.
      if (ch === "[") {
        bracketDepth += 1;
      } else if (ch === "]") {
        if (bracketDepth === 1 && formulaPrefix.startsWith("]]", i)) {
          // Excel escapes literal `]` as `]]` inside brackets.
          i += 2;
          continue;
        }
        bracketDepth = Math.max(0, bracketDepth - 1);
      }
      i += 1;
      continue;
    }

    switch (ch) {
      case "[":
        bracketDepth += 1;
        i += 1;
        continue;
      case "]":
        bracketDepth = Math.max(0, bracketDepth - 1);
        i += 1;
        continue;
      case "{":
        braceDepth += 1;
        i += 1;
        continue;
      case "}":
        braceDepth = Math.max(0, braceDepth - 1);
        i += 1;
        continue;
      case "(":
        parenDepth += 1;
        i += 1;
        continue;
      case ")":
        if (parenDepth > 0) {
          const top = stack[stack.length - 1];
          if (top && top.parenDepth === parenDepth) {
            stack.pop();
          }
          parenDepth -= 1;
        }
        i += 1;
        continue;
      default:
        break;
    }

    if (ch === argSeparator) {
      const top = stack[stack.length - 1];
      if (
        top &&
        parenDepth === top.parenDepth &&
        braceDepth === top.braceDepth &&
        bracketDepth === top.bracketDepth
      ) {
        top.argIndex += 1;
        top.lastArgSepIndex = i;
      }
      i += 1;
      continue;
    }

    if (isIdentStartChar(ch)) {
      const start = i;
      let end = i + 1;
      while (end < formulaPrefix.length && isIdentContChar(formulaPrefix[end]!)) end += 1;
      const ident = formulaPrefix.slice(start, end);

      // Look ahead for `(`, allowing whitespace.
      let j = end;
      while (j < formulaPrefix.length && /\s/.test(formulaPrefix[j]!)) j += 1;

      if (j < formulaPrefix.length && formulaPrefix[j] === "(") {
        parenDepth += 1;
        stack.push({
          name: toAsciiUpperCase(ident),
          parenDepth,
          braceDepth,
          bracketDepth,
          openParenIndex: j,
          argIndex: 0,
          lastArgSepIndex: null,
        });
        i = j + 1;
        continue;
      }

      i = end;
      continue;
    }

    i += 1;
  }

  return stack.length > 0 ? stack[stack.length - 1]! : null;
}

function buildContextFromFunctionCall(params: {
  input: string;
  cursor: number;
  rawFnName: string;
  canonicalFnName: string;
  argIndex: number;
  functionRegistry: RangeArgRegistry;
}): PartialFormulaContext {
  const { input, cursor, rawFnName, canonicalFnName, argIndex, functionRegistry } = params;

  const prefix = input.slice(0, cursor);
  const candidates = [",", ";"];
  let frame: FunctionCallFrame | null = null;

  // Prefer a candidate whose frame matches both the function name and argIndex from the engine.
  for (const sep of candidates) {
    const scanned = scanFunctionCallFrame(prefix, sep);
    if (!scanned) continue;
    if (scanned.name !== rawFnName) continue;
    if (scanned.argIndex !== argIndex) continue;
    frame = scanned;
    break;
  }

  // Fall back to any frame that matches the function name (even if argIndex is off).
  if (!frame) {
    for (const sep of candidates) {
      const scanned = scanFunctionCallFrame(prefix, sep);
      if (!scanned) continue;
      if (scanned.name !== rawFnName) continue;
      frame = scanned;
      break;
    }
  }

  // Last resort: no frame. Still return the function + arg index so completion logic works,
  // but use a conservative currentArg span.
  let spanStart = cursor;
  if (frame) {
    if (argIndex === 0) {
      spanStart = frame.openParenIndex + 1;
    } else if (frame.lastArgSepIndex != null) {
      spanStart = frame.lastArgSepIndex + 1;
    } else {
      spanStart = frame.openParenIndex + 1;
    }
  }

  while (spanStart < cursor && /\s/.test(input[spanStart]!)) spanStart += 1;

  const currentArg = {
    start: spanStart,
    end: cursor,
    text: input.slice(spanStart, cursor),
  };

  return {
    isFormula: true,
    inFunctionCall: true,
    functionName: canonicalFnName,
    argIndex,
    currentArg,
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
}): (input: string, cursorPosition: number, functionRegistry: RangeArgRegistry) => Promise<PartialFormulaContext> {
  const getEngineClient = options.getEngineClient ?? (() => null);
  const timeoutMs = Number.isFinite(options.timeoutMs) ? Math.max(1, Math.trunc(options.timeoutMs as number)) : 15;
  const unsupportedLocaleIds = new Set<string>();

  return async (
    input: string,
    cursorPosition: number,
    functionRegistry: RangeArgRegistry
  ): Promise<PartialFormulaContext> => {
    const cursor = clampCursor(input, cursorPosition);
    const prefix = input.slice(0, cursor);
    if (!prefix.startsWith("=")) {
      return { isFormula: false, inFunctionCall: false };
    }

    const localeId = getLocale();
    const engine = getEngineClient();
    if (!engine) {
      return canonicalizeInCallContext(
        parsePartialFormulaFallback(input, cursor, functionRegistry),
        localeId,
        functionRegistry
      );
    }
    if (unsupportedLocaleIds.has(localeId)) {
      return canonicalizeInCallContext(
        parsePartialFormulaFallback(input, cursor, functionRegistry),
        localeId,
        functionRegistry
      );
    }

    try {
      const result = await engine.parseFormulaPartial(input, cursor, { localeId }, { timeoutMs });
      const ctx = result?.context?.function ?? null;
      if (ctx && typeof ctx.name === "string" && Number.isInteger(ctx.argIndex) && ctx.argIndex >= 0) {
        const canonicalFnName = canonicalizeFunctionNameForLocale(ctx.name, localeId);
        return buildContextFromFunctionCall({
          input,
          cursor,
          rawFnName: ctx.name,
          canonicalFnName,
          argIndex: ctx.argIndex,
          functionRegistry,
        });
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (typeof message === "string" && message.startsWith("unknown localeId:")) {
        unsupportedLocaleIds.add(localeId);
      }
    }

    return canonicalizeInCallContext(
      parsePartialFormulaFallback(input, cursor, functionRegistry),
      localeId,
      functionRegistry
    );
  };
}
