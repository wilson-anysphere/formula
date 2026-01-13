import {
  parsePartialFormula as parsePartialFormulaFallback,
  type PartialFormulaContext,
} from "@formula/ai-completion";
import { getLocale } from "../../i18n/index.js";

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
  return /\p{L}/u.test(ch);
}

function isIdentContChar(ch: string): boolean {
  if (!ch) return false;
  if (ch === "$" || ch === "_" || ch === "\\" || ch === ".") return true;
  return /[\p{L}\p{N}]/u.test(ch);
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
  fnName: string;
  argIndex: number;
  functionRegistry: RangeArgRegistry;
}): PartialFormulaContext {
  const { input, cursor, fnName, argIndex, functionRegistry } = params;

  const prefix = input.slice(0, cursor);
  const candidates = [",", ";"];
  let frame: FunctionCallFrame | null = null;

  // Prefer a candidate whose frame matches both the function name and argIndex from the engine.
  for (const sep of candidates) {
    const scanned = scanFunctionCallFrame(prefix, sep);
    if (!scanned) continue;
    if (scanned.name !== fnName) continue;
    if (scanned.argIndex !== argIndex) continue;
    frame = scanned;
    break;
  }

  // Fall back to any frame that matches the function name (even if argIndex is off).
  if (!frame) {
    for (const sep of candidates) {
      const scanned = scanFunctionCallFrame(prefix, sep);
      if (!scanned) continue;
      if (scanned.name !== fnName) continue;
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
    functionName: fnName,
    argIndex,
    currentArg,
    expectingRange: Boolean(functionRegistry?.isRangeArg?.(fnName, argIndex)),
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

    const engine = getEngineClient();
    if (!engine) {
      return parsePartialFormulaFallback(input, cursor, functionRegistry);
    }

    const localeId = getLocale();

    try {
      const result = await engine.parseFormulaPartial(input, cursor, { localeId }, { timeoutMs });
      const ctx = result?.context?.function ?? null;
      if (ctx && typeof ctx.name === "string" && Number.isInteger(ctx.argIndex) && ctx.argIndex >= 0) {
        return buildContextFromFunctionCall({
          input,
          cursor,
          fnName: ctx.name,
          argIndex: ctx.argIndex,
          functionRegistry,
        });
      }
    } catch {
      // ignore; fall back below
    }

    return parsePartialFormulaFallback(input, cursor, functionRegistry);
  };
}
