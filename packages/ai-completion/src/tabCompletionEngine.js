import { LRUCache } from "./lruCache.js";
import { FunctionRegistry } from "./functionRegistry.js";
import { parsePartialFormula } from "./formulaPartialParser.js";
import { suggestRanges } from "./rangeSuggester.js";
import { suggestPatternValues } from "./patternSuggester.js";
import { normalizeCellRef, toA1, columnIndexToLetter } from "./a1.js";

/**
 * @typedef {"formula" | "value" | "function_arg" | "range"} SuggestionType
 *
 * @typedef {{
 *   text: string,
 *   displayText: string,
 *   type: SuggestionType,
 *   confidence: number,
 *   preview?: any
 * }} Suggestion
 *
 * @typedef {{
 *   currentInput: string,
 *   cursorPosition: number,
 *   cellRef: {row:number,col:number} | string,
 *   surroundingCells: { getCellValue: (row:number, col:number) => any, getCacheKey?: () => string }
 * }} CompletionContext
 */

export class TabCompletionEngine {
  /**
   * @param {{
   *   functionRegistry?: FunctionRegistry,
   *   parsePartialFormula?: typeof parsePartialFormula,
   *   localModel?: { complete: (prompt: string, options?: any) => Promise<string> } | null,
   *   cache?: LRUCache,
   *   cacheSize?: number,
   *   maxSuggestions?: number,
   *   localModelTimeoutMs?: number
   * }} [options]
   */
  constructor(options = {}) {
    this.functionRegistry = options.functionRegistry ?? new FunctionRegistry();
    this.parsePartialFormula = options.parsePartialFormula ?? parsePartialFormula;
    this.localModel = options.localModel ?? null;
    this.cache = options.cache ?? new LRUCache(options.cacheSize ?? 200);
    this.maxSuggestions = options.maxSuggestions ?? 5;
    this.localModelTimeoutMs = options.localModelTimeoutMs ?? 60;
  }

  /**
   * @param {CompletionContext} context
   * @returns {Promise<Suggestion[]>}
   */
  async getSuggestions(context) {
    const input = context?.currentInput ?? "";
    const cursorPosition = clampCursor(input, context?.cursorPosition);

    const cacheKey = this.buildCacheKey({
      ...context,
      currentInput: input,
      cursorPosition,
    });

    const cached = this.cache.get(cacheKey);
    if (cached) return cached;

    const parsed = this.parsePartialFormula(input, cursorPosition, this.functionRegistry);

    const [ruleBased, patternBased, localModelBased] = await Promise.all([
      this.getRuleBasedSuggestions(context, parsed),
      this.getPatternSuggestions(context, parsed),
      this.getLocalModelSuggestions(context, parsed),
    ]);

    const ranked = rankAndDedupe([
      ...ruleBased,
      ...patternBased,
      ...localModelBased,
    ]).slice(0, this.maxSuggestions);

    this.cache.set(cacheKey, ranked);
    return ranked;
  }

  buildCacheKey(context) {
    const cell = normalizeCellRef(context.cellRef);
    const surroundingKey =
      typeof context.surroundingCells?.getCacheKey === "function"
        ? context.surroundingCells.getCacheKey()
        : "";
    return JSON.stringify({
      input: context.currentInput,
      cursor: context.cursorPosition,
      cell,
      surroundingKey,
    });
  }

  /**
   * @param {CompletionContext} context
   * @param {ReturnType<typeof parsePartialFormula>} parsed
   * @returns {Promise<Suggestion[]>}
   */
  async getRuleBasedSuggestions(context, parsed) {
    if (!parsed.isFormula) return [];

    // 1) Function name completion
    if (!parsed.inFunctionCall && parsed.functionNamePrefix) {
      return this.suggestFunctionNames(context, parsed.functionNamePrefix);
    }

    if (!parsed.inFunctionCall || !parsed.functionName || !parsed.currentArg) return [];

    // 2) Range completion based on contiguous data
    if (parsed.expectingRange) {
      return this.suggestRangeCompletions(context, parsed);
    }

    // 3) Argument value hints for non-range args
    return this.suggestArgumentValues(context, parsed);
  }

  /**
   * @param {CompletionContext} context
   * @param {ReturnType<typeof parsePartialFormula>} parsed
   * @returns {Promise<Suggestion[]>}
   */
  async getPatternSuggestions(context, parsed) {
    if (parsed.isFormula) return [];
    const candidates = suggestPatternValues(context);
    return candidates.map(c => ({
      text: c.text,
      displayText: c.text,
      type: "value",
      confidence: c.confidence,
    }));
  }

  /**
   * @param {CompletionContext} context
   * @param {ReturnType<typeof parsePartialFormula>} parsed
   * @returns {Promise<Suggestion[]>}
   */
  async getLocalModelSuggestions(context, parsed) {
    if (!this.localModel) return [];
    if (!parsed.isFormula) return [];

    // Rule-based completions are often better than LLM for function names and
    // argument structure. Only ask the model when we have an actual formula body.
    if (parsed.functionNamePrefix) return [];

    const input = context.currentInput ?? "";
    const cursor = clampCursor(input, context.cursorPosition);
    const cell = normalizeCellRef(context.cellRef);

    const prompt = buildLocalModelPrompt({
      input,
      cursorPosition: cursor,
      cellA1: toA1(cell),
    });

    try {
      const completion = await withTimeout(
        this.localModel.complete(prompt, {
          maxTokens: 50,
          temperature: 0.1,
          stop: [")", ",", "\n"],
        }),
        this.localModelTimeoutMs
      );

      const suggestionText = normalizeLocalCompletion(input, cursor, completion);
      if (!suggestionText) return [];

      return [
        {
          text: suggestionText,
          displayText: suggestionText,
          type: "formula",
          confidence: 0.35,
        },
      ];
    } catch {
      // LLM is optional; ignore failures and timeouts.
      return [];
    }
  }

  /**
   * @param {CompletionContext} context
   * @param {{text:string,start:number,end:number}} token
   * @returns {Suggestion[]}
   */
  suggestFunctionNames(context, token) {
    const input = context.currentInput ?? "";
    const prefix = token.text;
    const matches = this.functionRegistry.search(prefix, { limit: 10 });

    /** @type {Suggestion[]} */
    const suggestions = [];

    for (const spec of matches) {
      const completedName = applyNameCase(spec.name, prefix);
      // Preserve the user-typed prefix exactly so inline "ghost text" can be
      // rendered as a pure insertion at the cursor.
      const remainder = completedName.slice(prefix.length);
      const replacement = `${prefix}${remainder}(`;
      const newText = replaceSpan(input, token.start, token.end, replacement);
      suggestions.push({
        text: newText,
        displayText: replacement,
        type: "formula",
        confidence: clamp01(0.6 + (prefix.length / spec.name.length) * 0.4),
      });

      // Provide lightweight "modern alternative" suggestions for some legacy functions.
      if (spec.name === "VLOOKUP") {
        suggestions.push({
          text: replaceSpan(input, token.start, token.end, "XLOOKUP("),
          displayText: "XLOOKUP(",
          type: "formula",
          confidence: 0.35,
        });
      }
    }

    return rankAndDedupe(suggestions).slice(0, this.maxSuggestions);
  }

  /**
   * @param {CompletionContext} context
   * @param {ReturnType<typeof parsePartialFormula>} parsed
   * @returns {Suggestion[]}
   */
  suggestRangeCompletions(context, parsed) {
    const input = context.currentInput ?? "";
    const cursor = clampCursor(input, context.cursorPosition);
    const cellRef = normalizeCellRef(context.cellRef);

    const rangeCandidates = suggestRanges({
      currentArgText: parsed.currentArg.text,
      cellRef,
      surroundingCells: context.surroundingCells,
    });

    /** @type {Suggestion[]} */
    const suggestions = [];

    for (const candidate of rangeCandidates) {
      let newText = replaceSpan(input, parsed.currentArg.start, parsed.currentArg.end, candidate.range);

      // If the user is at the end of input and we still have unbalanced parens,
      // auto-close them. This matches the common "type =SUM(A<tab>" workflow.
      if (cursor === input.length) {
        newText = closeUnbalancedParens(newText);
      }

      suggestions.push({
        text: newText,
        displayText: newText,
        type: "range",
        confidence: candidate.confidence,
      });
    }

    return suggestions;
  }

  /**
   * @param {CompletionContext} context
   * @param {ReturnType<typeof parsePartialFormula>} parsed
   * @returns {Suggestion[]}
   */
  suggestArgumentValues(context, parsed) {
    const input = context.currentInput ?? "";
    const cursor = clampCursor(input, context.cursorPosition);
    const cellRef = normalizeCellRef(context.cellRef);

    const fnName = parsed.functionName;
    const argIndex = parsed.argIndex ?? 0;
    const argType = this.functionRegistry.getArgType(fnName, argIndex) ?? "any";

    const spanStart = parsed.currentArg?.start ?? cursor;
    const spanEnd = parsed.currentArg?.end ?? cursor;

    /** @type {Suggestion[]} */
    const suggestions = [];

    if (argType === "boolean") {
      for (const boolLiteral of ["TRUE", "FALSE"]) {
        suggestions.push({
          text: replaceSpan(input, spanStart, spanEnd, boolLiteral),
          displayText: boolLiteral,
          type: "function_arg",
          confidence: 0.5,
        });
      }
      return suggestions;
    }

    if (argType === "number") {
      for (const n of ["1", "0"]) {
        suggestions.push({
          text: replaceSpan(input, spanStart, spanEnd, n),
          displayText: n,
          type: "function_arg",
          confidence: 0.4,
        });
      }
      return suggestions;
    }

    if (argType === "value") {
      // Common heuristic: reference the cell to the left.
      if (cellRef.col > 0) {
        const leftA1 = `${columnIndexToLetter(cellRef.col - 1)}${cellRef.row + 1}`;
        suggestions.push({
          text: replaceSpan(input, spanStart, spanEnd, leftA1),
          displayText: leftA1,
          type: "function_arg",
          confidence: 0.35,
        });
      }
      return suggestions;
    }

    return [];
  }
}

function buildLocalModelPrompt({ input, cursorPosition, cellA1 }) {
  return [
    "You are a spreadsheet formula completion engine.",
    "Return ONLY the text to insert at the cursor.",
    `Cell: ${cellA1}`,
    `Input: ${input}`,
    `CursorPosition: ${cursorPosition}`,
    "Completion:",
  ].join("\n");
}

function normalizeLocalCompletion(input, cursorPosition, completion) {
  const raw = (completion ?? "").toString().trim();
  if (!raw) return "";

  // If the model returned a full formula, trust it.
  if (raw.startsWith("=")) return raw;

  // Otherwise treat it as text to insert at the cursor.
  const before = input.slice(0, cursorPosition);
  const after = input.slice(cursorPosition);
  return `${before}${raw}${after}`;
}

function clampCursor(input, cursorPosition) {
  if (!Number.isInteger(cursorPosition)) return input.length;
  if (cursorPosition < 0) return 0;
  if (cursorPosition > input.length) return input.length;
  return cursorPosition;
}

function replaceSpan(input, start, end, replacement) {
  return `${input.slice(0, start)}${replacement}${input.slice(end)}`;
}

function closeUnbalancedParens(input) {
  let balance = 0;
  for (const ch of input) {
    if (ch === "(") balance++;
    else if (ch === ")") balance = Math.max(0, balance - 1);
  }
  return balance > 0 ? input + ")".repeat(balance) : input;
}

function applyNameCase(name, typedPrefix) {
  if (typedPrefix && typedPrefix === typedPrefix.toLowerCase()) {
    return name.toLowerCase();
  }
  if (typedPrefix && typedPrefix === typedPrefix.toUpperCase()) {
    return name.toUpperCase();
  }
  return name;
}

/**
 * @param {Suggestion[]} suggestions
 */
function rankAndDedupe(suggestions) {
  /** @type {Map<string, Suggestion>} */
  const bestByText = new Map();
  for (const s of suggestions) {
    if (!s || typeof s.text !== "string" || s.text.length === 0) continue;
    const existing = bestByText.get(s.text);
    if (!existing || (s.confidence ?? 0) > (existing.confidence ?? 0)) {
      bestByText.set(s.text, s);
    }
  }

  const typePriority = {
    range: 4,
    formula: 3,
    function_arg: 2,
    value: 1,
  };

  return [...bestByText.values()].sort((a, b) => {
    const confDiff = (b.confidence ?? 0) - (a.confidence ?? 0);
    if (confDiff !== 0) return confDiff;
    const typeDiff = (typePriority[b.type] ?? 0) - (typePriority[a.type] ?? 0);
    if (typeDiff !== 0) return typeDiff;
    return a.displayText.localeCompare(b.displayText);
  });
}

function clamp01(v) {
  return Math.max(0, Math.min(1, v));
}

/**
 * @template T
 * @param {Promise<T>} promise
 * @param {number} timeoutMs
 * @returns {Promise<T>}
 */
function withTimeout(promise, timeoutMs) {
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) return promise;
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error("timeout")), timeoutMs)),
  ]);
}
