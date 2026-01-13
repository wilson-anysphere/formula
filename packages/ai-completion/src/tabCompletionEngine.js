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
 *   surroundingCells: { getCellValue: (row:number, col:number, sheetName?: string) => any, getCacheKey?: () => string }
 * }} CompletionContext
 *
 * @typedef {{
 *   name: string,
 *   range?: string
 * }} NamedRangeInfo
 *
 * @typedef {{
 *   name: string,
 *   columns: string[],
 *   sheetName?: string,
 *   startRow?: number,
 *   startCol?: number,
 *   endRow?: number,
 *   endCol?: number
 * }} TableInfo
 *
 * @typedef {{
 *   getNamedRanges?: () => NamedRangeInfo[] | Promise<NamedRangeInfo[]>,
 *   getSheetNames?: () => string[] | Promise<string[]>,
 *   getTables?: () => TableInfo[] | Promise<TableInfo[]>,
 *   getCacheKey?: () => string
 * }} SchemaProvider
 *
 * @typedef {(params: { suggestion: Suggestion, context: CompletionContext }) => any | Promise<any>} PreviewEvaluator
 */

export class TabCompletionEngine {
  /**
   * @param {{
   *   functionRegistry?: FunctionRegistry,
   *   parsePartialFormula?: typeof parsePartialFormula,
   *   completionClient?: { completeTabCompletion: (req: { input: string, cursorPosition: number, cellA1: string }) => Promise<string> } | null,
   *   schemaProvider?: SchemaProvider | null,
   *   cache?: LRUCache,
   *   cacheSize?: number,
   *   maxSuggestions?: number,
   *   completionTimeoutMs?: number
   * }} [options]
   */
  constructor(options = {}) {
    this.functionRegistry = options.functionRegistry ?? new FunctionRegistry();
    this.parsePartialFormula = options.parsePartialFormula ?? parsePartialFormula;
    this.completionClient = options.completionClient ?? null;
    this.schemaProvider = options.schemaProvider ?? null;
    this.cache = options.cache ?? new LRUCache(options.cacheSize ?? 200);
    this.maxSuggestions = options.maxSuggestions ?? 5;
    const rawTimeout = options.completionTimeoutMs ?? 100;
    // Tab completion should stay responsive; clamp to a small budget.
    this.completionTimeoutMs = Number.isFinite(rawTimeout) ? Math.max(1, Math.min(rawTimeout, 200)) : 100;
  }

  /**
   * @param {CompletionContext} context
   * @param {{ previewEvaluator?: PreviewEvaluator }} [options]
   * @returns {Promise<Suggestion[]>}
   */
  async getSuggestions(context, options = {}) {
    const input = context?.currentInput ?? "";
    const cursorPosition = clampCursor(input, context?.cursorPosition);

    const cacheKey = this.buildCacheKey({
      ...context,
      currentInput: input,
      cursorPosition,
    });

    const cached = this.cache.get(cacheKey);
    const baseSuggestions = cached ?? (await this.#computeBaseSuggestions(context, input, cursorPosition));
    if (!cached) this.cache.set(cacheKey, baseSuggestions);

    if (typeof options?.previewEvaluator === "function") {
      return attachPreviews(baseSuggestions, context, options.previewEvaluator);
    }
    return baseSuggestions;
  }

  buildCacheKey(context) {
    const cell = normalizeCellRef(context.cellRef);
    const surroundingKey =
      typeof context.surroundingCells?.getCacheKey === "function"
        ? context.surroundingCells.getCacheKey()
        : "";
    const schemaKey =
      typeof this.schemaProvider?.getCacheKey === "function" ? this.schemaProvider.getCacheKey() : "";
    return JSON.stringify({
      input: context.currentInput,
      cursor: context.cursorPosition,
      cell,
      surroundingKey,
      schemaKey,
    });
  }

  async #computeBaseSuggestions(context, input, cursorPosition) {
    const parsed = this.parsePartialFormula(input, cursorPosition, this.functionRegistry);

    const [ruleBased, patternBased, backendBased] = await Promise.all([
      this.getRuleBasedSuggestions(context, parsed),
      this.getPatternSuggestions(context, parsed),
      this.getCursorBackendSuggestions(context, parsed),
    ]);

    return rankAndDedupe([
      ...ruleBased,
      ...patternBased,
      ...backendBased,
    ]).slice(0, this.maxSuggestions);
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
      const [fnSuggestions, identifierSuggestions] = await Promise.all([
        Promise.resolve(this.suggestFunctionNames(context, parsed.functionNamePrefix)),
        this.suggestWorkbookIdentifiers(context, parsed.functionNamePrefix),
      ]);
      return rankAndDedupe([...fnSuggestions, ...identifierSuggestions]).slice(0, this.maxSuggestions);
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
  async getCursorBackendSuggestions(context, parsed) {
    if (!this.completionClient) return [];
    if (!parsed.isFormula) return [];

    // Rule-based completions are often better than LLM for function names and
    // argument structure. Only ask the backend when we have an actual formula body.
    if (parsed.functionNamePrefix) return [];

    const input = context.currentInput ?? "";
    const cursor = clampCursor(input, context.cursorPosition);
    const cell = normalizeCellRef(context.cellRef);

    try {
      const completion = await withTimeout(
        this.completionClient.completeTabCompletion({
          input,
          cursorPosition: cursor,
          cellA1: toA1(cell),
        }),
        this.completionTimeoutMs
      );

      const suggestionText = normalizeBackendCompletion(input, cursor, completion);
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
      // Backend is optional; ignore failures and timeouts.
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
      const callSuffix = spec.minArgs === 0 && spec.maxArgs === 0 ? "()" : "(";
      const replacement = `${prefix}${remainder}${callSuffix}`;
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
   * @returns {Promise<Suggestion[]>}
   */
  suggestRangeCompletions(context, parsed) {
    const input = context.currentInput ?? "";
    const cursor = clampCursor(input, context.cursorPosition);
    const cellRef = normalizeCellRef(context.cellRef);
    const fnSpec = parsed.functionName ? this.functionRegistry.getFunction(parsed.functionName) : undefined;
    const argIndex = parsed.argIndex ?? 0;
    const argSpecName = fnSpec?.args?.[argIndex]?.name;
    const functionCouldBeComplete = functionCouldBeCompleteAfterArg(fnSpec, argIndex);

    const rangeCandidates = suggestRanges({
      currentArgText: parsed.currentArg.text,
      cellRef,
      surroundingCells: context.surroundingCells,
    });

    // Some functions (VLOOKUP table_array, TAKE array, etc.) almost always want
    // a 2D rectangular range when the surrounding data forms a table. When we
    // have both a 1D and 2D candidate, slightly bias toward the 2D option so
    // tab completion defaults to the more useful table-shaped range.
    const prefersTableRange = argSpecName === "table_array" || argSpecName === "array";
    const hasTableCandidate = rangeCandidates.some(
      (c) => typeof c?.reason === "string" && c.reason.startsWith("contiguous_table")
    );

    /** @type {Suggestion[]} */
    const suggestions = [];

    for (const candidate of rangeCandidates) {
      let confidence = candidate.confidence;
      if (prefersTableRange && hasTableCandidate) {
        const isTable = typeof candidate?.reason === "string" && candidate.reason.startsWith("contiguous_table");
        if (isTable) confidence = clamp01(confidence + 0.2);
        if (
          !isTable &&
          (candidate.reason === "contiguous_above_current_cell" || candidate.reason === "contiguous_down_from_start")
        ) {
          confidence = clamp01(confidence - 0.1);
        }
      }

      let newText = replaceSpan(input, parsed.currentArg.start, parsed.currentArg.end, candidate.range);

      // If the user is at the end of input and we still have unbalanced parens,
      // auto-close them when the function could be syntactically complete. This
      // matches the common "type =SUM(A<tab>" workflow without producing invalid
      // suggestions for functions that require additional arguments (e.g. VLOOKUP).
      if (cursor === input.length && functionCouldBeComplete) {
        newText = closeUnbalancedParens(newText);
      }

      suggestions.push({
        text: newText,
        displayText: newText,
        type: "range",
        confidence,
      });
    }

    return Promise.resolve()
      .then(() => this.suggestSchemaRanges(context, parsed))
      .then((schemaSuggestions) => {
        const merged = rankAndDedupe([...suggestions, ...schemaSuggestions]).slice(0, this.maxSuggestions);
        return merged;
      });
  }

  /**
   * Suggest named ranges, sheet-qualified ranges, and structured references for the current arg.
   *
   * @param {CompletionContext} context
   * @param {ReturnType<typeof parsePartialFormula>} parsed
   * @returns {Promise<Suggestion[]>}
   */
  async suggestSchemaRanges(context, parsed) {
    const provider = this.schemaProvider;
    if (!provider) return [];

    const input = context.currentInput ?? "";
    const cursor = clampCursor(input, context.cursorPosition);
    const cellRef = normalizeCellRef(context.cellRef);

    const fnSpec = parsed.functionName ? this.functionRegistry.getFunction(parsed.functionName) : undefined;
    const argIndex = parsed.argIndex ?? 0;
    const functionCouldBeComplete = functionCouldBeCompleteAfterArg(fnSpec, argIndex);

    const spanStart = parsed.currentArg?.start ?? cursor;
    const spanEnd = parsed.currentArg?.end ?? cursor;
    const rawPrefix = (parsed.currentArg?.text ?? "").trim();

    /** @type {Suggestion[]} */
    const suggestions = [];

    const addReplacement = (replacement, { confidence }) => {
      let newText = replaceSpan(input, spanStart, spanEnd, replacement);
      if (cursor === input.length && functionCouldBeComplete) {
        newText = closeUnbalancedParens(newText);
      }
      suggestions.push({
        text: newText,
        displayText: replacement,
        type: "range",
        confidence,
      });
    };

    // 1) Named ranges
    const namedRanges = await safeProviderCall(provider.getNamedRanges);
    for (const entry of namedRanges) {
      const name = (entry?.name ?? "").toString();
      if (!name) continue;
      if (rawPrefix && !startsWithIgnoreCase(name, rawPrefix)) continue;
      addReplacement(completeIdentifier(name, rawPrefix), {
        confidence: clamp01(0.65 + ratioBoost(rawPrefix, name) * 0.25),
      });
    }

    // 2) Structured references (tables)
    const tables = await safeProviderCall(provider.getTables);
    if (tables.length > 0) {
      const structuredMatches = suggestStructuredRefs(rawPrefix, tables);
      for (const m of structuredMatches) {
        addReplacement(m.text, { confidence: m.confidence });
      }
    }

    // 3) Sheet-qualified A1 ranges (Sheet2!A1:A10)
    const sheetNames = await safeProviderCall(provider.getSheetNames);
    const sheetArg = splitSheetQualifiedArg(rawPrefix);
    if (sheetArg) {
      const typedQuoted = rawPrefix.startsWith("'");
      const sheetName = sheetNames
        .filter((s) => typeof s === "string" && s.length > 0)
        .find((s) => s.toLowerCase() === sheetArg.sheetPrefix.toLowerCase());
      if (sheetName) {
        if (!(needsSheetQuotes(sheetName) && !typedQuoted)) {
          const rangeCandidates = suggestRanges({
            currentArgText: sheetArg.rangePrefix,
            cellRef,
            surroundingCells: context.surroundingCells,
            sheetName,
          });
          for (const candidate of rangeCandidates) {
            addReplacement(`${rawPrefix}${candidate.range.slice(sheetArg.rangePrefix.length)}`, {
              confidence: Math.min(0.85, candidate.confidence + 0.05),
            });
          }
        }
      }
    }

    return rankAndDedupe(suggestions).slice(0, this.maxSuggestions);
  }

  /**
   * Suggest identifiers (named ranges, tables) when not in a function call.
   *
   * @param {CompletionContext} context
   * @param {{text:string,start:number,end:number}} token
   * @returns {Promise<Suggestion[]>}
   */
  async suggestWorkbookIdentifiers(context, token) {
    const provider = this.schemaProvider;
    if (!provider) return [];

    const input = context.currentInput ?? "";
    const prefix = token.text;

    /** @type {Suggestion[]} */
    const suggestions = [];

    const namedRanges = await safeProviderCall(provider.getNamedRanges);
    for (const entry of namedRanges) {
      const name = (entry?.name ?? "").toString();
      if (!name) continue;
      if (!startsWithIgnoreCase(name, prefix)) continue;
      const replacement = completeIdentifier(name, prefix);
      const newText = replaceSpan(input, token.start, token.end, replacement);
      suggestions.push({
        text: newText,
        displayText: replacement,
        type: "range",
        confidence: clamp01(0.55 + ratioBoost(prefix, name) * 0.35),
      });
    }

    const tables = await safeProviderCall(provider.getTables);
    if (tables.length > 0) {
      const structuredMatches = suggestStructuredRefs(prefix, tables);
      for (const m of structuredMatches) {
        const newText = replaceSpan(input, token.start, token.end, m.text);
        suggestions.push({
          text: newText,
          displayText: m.text,
          type: "range",
          confidence: m.confidence,
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

function normalizeBackendCompletion(input, cursorPosition, completion) {
  const raw = (completion ?? "").toString().trim();
  if (!raw) return "";

  // If the backend returned a full formula, trust it.
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

/**
 * Returns true when the formula could be syntactically complete after providing the given arg.
 *
 * For most functions we only need to know whether `minArgs` is satisfied. However, some functions
 * accept repeating *groups* of args where entering an early arg in the group implies additional
 * required args must follow (e.g. SUMIFS(..., criteria_range2, criteria2)).
 *
 * @param {any} fnSpec
 * @param {number} argIndex
 */
function functionCouldBeCompleteAfterArg(fnSpec, argIndex) {
  const minArgs = fnSpec?.minArgs;
  if (!Number.isInteger(minArgs)) return true;
  if (!Number.isInteger(argIndex) || argIndex + 1 < minArgs) return false;

  const args = Array.isArray(fnSpec?.args) ? fnSpec.args : [];
  if (args.length === 0) return true;

  const repeatingStart = args.findIndex((a) => a?.repeating);
  if (repeatingStart < 0) return true;
  if (argIndex < repeatingStart) return true;

  const group = args.slice(repeatingStart);
  if (group.length <= 1) return true;

  const within = (argIndex - repeatingStart) % group.length;
  const requiredAfter = group.slice(within + 1).some((spec) => !spec?.optional);
  return !requiredAfter;
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

function applyIdentifierCase(name, typedPrefix) {
  if (!typedPrefix) return name;

  // If the user is typing in uppercase, treat that as an explicit signal to
  // uppercase the full identifier (common for people who prefer Excel-style
  // uppercase references).
  if (typedPrefix === typedPrefix.toUpperCase()) return name.toUpperCase();

  // If the user is typing in lowercase, preserve internal capitalization for
  // CamelCase identifiers (e.g. named ranges), but avoid producing "suM"-style
  // completions for ALL-CAPS identifiers by downcasing them fully.
  if (typedPrefix === typedPrefix.toLowerCase() && name === name.toUpperCase()) return name.toLowerCase();

  return name;
}

function completeIdentifier(name, typedPrefix) {
  if (!typedPrefix) return name;
  const cased = applyIdentifierCase(name, typedPrefix);
  if (typedPrefix.length >= cased.length) return typedPrefix;
  return `${typedPrefix}${cased.slice(typedPrefix.length)}`;
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

async function safeProviderCall(fn) {
  if (typeof fn !== "function") return [];
  try {
    const result = fn();
    const awaited = await Promise.resolve(result);
    return Array.isArray(awaited) ? awaited : [];
  } catch {
    return [];
  }
}

function startsWithIgnoreCase(text, prefix) {
  if (!prefix) return true;
  return text.toLowerCase().startsWith(prefix.toLowerCase());
}

function ratioBoost(prefix, full) {
  if (!prefix || !full) return 0;
  return Math.min(1, prefix.length / full.length);
}

function splitSheetQualifiedArg(text) {
  if (!text || !text.includes("!")) return null;

  // Handle quoted sheet names: 'Sheet 1'!A
  if (text.startsWith("'")) {
    let name = "";
    let i = 1;
    while (i < text.length) {
      const ch = text[i];
      if (ch === "'") {
        if (text[i + 1] === "'") {
          name += "'";
          i += 2;
          continue;
        }
        // Closing quote must be followed by !
        if (text[i + 1] === "!") {
          const rangePrefix = text.slice(i + 2);
          return { sheetPrefix: name, rangePrefix };
        }
        return null;
      }
      name += ch;
      i += 1;
    }
    return null;
  }

  const bang = text.indexOf("!");
  if (bang <= 0) return null;
  const sheetPrefix = text.slice(0, bang);
  const rangePrefix = text.slice(bang + 1);
  return { sheetPrefix, rangePrefix };
}

function needsSheetQuotes(sheetName) {
  // Match Excel's "unquoted sheet name" rules (roughly identifier-like).
  //
  // Note: avoid emitting ambiguous unquoted prefixes like:
  // - TRUE!A1 / FALSE!A1 (boolean literals)
  // - A1!B2 / XFD1048576!A1 (looks like an A1 cell reference)
  // - R1C1!A1 / RC!A1 (looks like an R1C1 cell reference)
  //
  // See similar logic in the Rust backend (`formula_model::needs_quoting_for_sheet_reference`).
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)) return true;
  const lower = sheetName.toLowerCase();
  if (lower === "true" || lower === "false") return true;
  if (looksLikeA1CellReference(sheetName) || looksLikeR1C1CellReference(sheetName)) return true;
  return false;
}

function formatSheetPrefix(sheetName) {
  const needsQuotes = needsSheetQuotes(sheetName);
  if (!needsQuotes) return `${sheetName}!`;
  const escaped = sheetName.replaceAll("'", "''");
  return `'${escaped}'!`;
}

function looksLikeA1CellReference(name) {
  let i = 0;
  let letters = "";
  while (i < name.length) {
    const ch = name[i];
    if (!ch || !/[A-Za-z]/.test(ch)) break;
    if (letters.length >= 3) return false;
    letters += ch;
    i += 1;
  }
  if (letters.length === 0) return false;

  let digits = "";
  while (i < name.length) {
    const ch = name[i];
    if (!ch || !/[0-9]/.test(ch)) break;
    digits += ch;
    i += 1;
  }
  if (digits.length === 0) return false;
  if (i !== name.length) return false;

  // Convert col letters to 1-based index and compare against Excel max col (XFD = 16384).
  const col = letters
    .toUpperCase()
    .split("")
    .reduce((acc, c) => acc * 26 + (c.charCodeAt(0) - 64), 0);
  return col <= 16384;
}

function looksLikeR1C1CellReference(name) {
  const upper = String(name ?? "").toUpperCase();
  if (upper === "R" || upper === "C") return true;
  if (!upper.startsWith("R")) return false;
  let i = 1;
  while (i < upper.length && upper[i] >= "0" && upper[i] <= "9") i += 1;
  if (i >= upper.length) return false;
  if (upper[i] !== "C") return false;
  i += 1;
  while (i < upper.length && upper[i] >= "0" && upper[i] <= "9") i += 1;
  return i === upper.length;
}

function suggestStructuredRefs(prefix, tables) {
  /** @type {{text:string, confidence:number}[]} */
  const out = [];
  const trimmed = (prefix ?? "").trim();

  // If the user typed something like Table1[Col, use that as a strong hint.
  const bracketIdx = trimmed.indexOf("[");
  const tablePrefix = bracketIdx >= 0 ? trimmed.slice(0, bracketIdx) : trimmed;
  const colPrefix = bracketIdx >= 0 ? trimmed.slice(bracketIdx + 1).replaceAll("]", "") : "";

  for (const table of tables) {
    const tableName = (table?.name ?? "").toString();
    const cols = Array.isArray(table?.columns) ? table.columns : [];
    if (!tableName || cols.length === 0) continue;

    if (tablePrefix && !startsWithIgnoreCase(tableName, tablePrefix)) continue;

    for (const col of cols) {
      const colName = (col ?? "").toString();
      if (!colName) continue;
      if (colPrefix && !startsWithIgnoreCase(colName, colPrefix)) continue;
      const canonical = `${tableName}[${colName}]`;
      if (trimmed && !startsWithIgnoreCase(canonical, trimmed)) continue;
      const text = completeIdentifier(canonical, trimmed);
      out.push({
        text,
        confidence: clamp01(0.6 + ratioBoost(colPrefix || tablePrefix, colName || tableName) * 0.25),
      });
      // Lower confidence alternative that includes #All.
      const allCanonical = `${tableName}[[#All],[${colName}]]`;
      if (trimmed && !startsWithIgnoreCase(allCanonical, trimmed)) continue;
      out.push({ text: completeIdentifier(allCanonical, trimmed), confidence: 0.35 });
    }
  }

  return out;
}

async function attachPreviews(suggestions, context, previewEvaluator) {
  /** @type {Suggestion[]} */
  const out = [];
  for (const s of suggestions) {
    if (!s || typeof s.text !== "string") continue;
    if (s.type !== "formula" && s.type !== "range") {
      out.push(s);
      continue;
    }
    try {
      const preview = await previewEvaluator({ suggestion: s, context });
      if (preview === undefined) {
        out.push(s);
      } else {
        out.push({ ...s, preview });
      }
    } catch {
      out.push(s);
    }
  }
  return out;
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
