import { LRUCache } from "./lruCache.js";
import { FunctionRegistry } from "./functionRegistry.js";
import { parsePartialFormula as parsePartialFormulaFallback } from "./formulaPartialParser.js";
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
   *   parsePartialFormula?: typeof parsePartialFormulaFallback | ((input: string, cursorPosition: number, functionRegistry: FunctionRegistry) => any | Promise<any>),
   *   completionClient?: { completeTabCompletion: (req: { input: string, cursorPosition: number, cellA1: string, signal?: AbortSignal }) => Promise<string> } | null,
   *   schemaProvider?: SchemaProvider | null,
   *   cache?: LRUCache,
   *   cacheSize?: number,
   *   maxSuggestions?: number,
   *   completionTimeoutMs?: number
   * }} [options]
   */
  constructor(options = {}) {
    this.functionRegistry = options.functionRegistry ?? new FunctionRegistry();
    this.parsePartialFormula = options.parsePartialFormula ?? parsePartialFormulaFallback;
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
   * @param {{ previewEvaluator?: PreviewEvaluator, signal?: AbortSignal }} [options]
   * @returns {Promise<Suggestion[]>}
   */
  async getSuggestions(context, options = {}) {
    try {
      const input = safeToString(context?.currentInput);
      const cursorPosition = clampCursor(input, context?.cursorPosition);
      const requestSignal = options?.signal;

      const normalizedContext = {
        currentInput: input,
        cursorPosition,
        // Normalize defensively because tab completion runs on every keystroke.
        cellRef: safeNormalizeCellRef(context?.cellRef),
        surroundingCells: context?.surroundingCells,
      };

      const cacheKey = this.buildCacheKey(normalizedContext);

      const cached = safeCacheGet(this.cache, cacheKey);
      const baseSuggestions = Array.isArray(cached)
        ? cached
        : await this.#computeBaseSuggestions(normalizedContext, input, cursorPosition, requestSignal);

      if (!Array.isArray(cached)) {
        safeCacheSet(this.cache, cacheKey, baseSuggestions);
      }

      if (typeof options?.previewEvaluator === "function") {
        try {
          const withPreviews = await attachPreviews(baseSuggestions, normalizedContext, options.previewEvaluator);
          return Array.isArray(withPreviews) ? withPreviews : [];
        } catch {
          return Array.isArray(baseSuggestions) ? baseSuggestions : [];
        }
      }

      return Array.isArray(baseSuggestions) ? baseSuggestions : [];
    } catch {
      // Tab completion must never crash the host.
      return [];
    }
  }

  buildCacheKey(context) {
    let degraded = false;

    let cell = DEFAULT_CELL_REF;
    try {
      // Defensive normalization for cache keys: invalid refs should not throw.
      const normalized = normalizeCellRef(context?.cellRef);
      if (
        normalized &&
        Number.isInteger(normalized.row) &&
        Number.isInteger(normalized.col) &&
        normalized.row >= 0 &&
        normalized.col >= 0
      ) {
        cell = { row: normalized.row, col: normalized.col };
      } else {
        degraded = true;
        cell = DEFAULT_CELL_REF;
      }
    } catch {
      degraded = true;
      cell = DEFAULT_CELL_REF;
    }

    let surroundingKey = "";
    try {
      surroundingKey =
        typeof context?.surroundingCells?.getCacheKey === "function"
          ? safeToString(context.surroundingCells.getCacheKey())
          : "";
    } catch {
      degraded = true;
      surroundingKey = "";
    }

    let schemaKey = "";
    try {
      schemaKey =
        typeof this.schemaProvider?.getCacheKey === "function" ? safeToString(this.schemaProvider.getCacheKey()) : "";
    } catch {
      degraded = true;
      schemaKey = "";
    }

    const input = safeToString(context?.currentInput);
    const cursor = clampCursor(input, context?.cursorPosition);

    const payload = { input, cursor, cell, surroundingKey, schemaKey };

    // If any component failed, return a degraded-but-stable key instead of throwing.
    if (degraded) {
      return `degraded:${input}:${cursor}:${cell.row},${cell.col}:${surroundingKey}:${schemaKey}`;
    }

    try {
      return JSON.stringify(payload);
    } catch {
      return `degraded:${input}:${cursor}:${cell.row},${cell.col}:${surroundingKey}:${schemaKey}`;
    }
  }

  async #computeBaseSuggestions(context, input, cursorPosition, requestSignal) {
    /** @type {ReturnType<typeof parsePartialFormulaFallback>} */
    let parsed;
    try {
      parsed = await Promise.resolve(this.parsePartialFormula(input, cursorPosition, this.functionRegistry));
    } catch {
      // Parsing is best-effort. If a caller-provided parser throws (e.g. WASM engine unavailable),
      // fall back to the built-in JS parser so tab completion remains responsive.
      try {
        parsed = parsePartialFormulaFallback(input, cursorPosition, this.functionRegistry);
      } catch {
        // Be defensive: if even the built-in parser throws, treat as non-formula so pattern
        // suggestions can still run.
        parsed = { isFormula: false, inFunctionCall: false };
      }
    }

    const [ruleBased, patternBased, backendBased] = await Promise.all([
      safeArrayResult(() => this.getRuleBasedSuggestions(context, parsed)),
      safeArrayResult(() => this.getPatternSuggestions(context, parsed)),
      safeArrayResult(() => this.getCursorBackendSuggestions(context, parsed, requestSignal)),
    ]);

    return rankAndDedupe([...ruleBased, ...patternBased, ...backendBased]).slice(0, this.maxSuggestions);
  }

  /**
   * @param {CompletionContext} context
   * @param {ReturnType<typeof parsePartialFormulaFallback>} parsed
   * @returns {Promise<Suggestion[]>}
   */
  async getRuleBasedSuggestions(context, parsed) {
    if (!parsed.isFormula) return [];

    // 0) Default "starter" functions when the user just typed "=" and hasn't
    // begun entering a function name yet.
    if (!parsed.inFunctionCall && !parsed.functionNamePrefix) {
      const topLevel = this.suggestTopLevelFunctions(context);
      if (topLevel.length > 0) return topLevel;
    }

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
   * @param {ReturnType<typeof parsePartialFormulaFallback>} parsed
   * @returns {Promise<Suggestion[]>}
   */
  async getPatternSuggestions(context, parsed) {
    if (parsed.isFormula) return [];
    try {
      const candidates = suggestPatternValues({
        ...context,
        // patternSuggester normalizes internally; ensure it always receives a valid ref.
        cellRef: safeNormalizeCellRef(context?.cellRef),
      });
      return candidates.map((c) => ({
        text: c.text,
        displayText: c.text,
        type: "value",
        confidence: c.confidence,
      }));
    } catch {
      return [];
    }
  }

  /**
   * @param {CompletionContext} context
   * @param {ReturnType<typeof parsePartialFormulaFallback>} parsed
   * @param {AbortSignal | undefined} requestSignal
   * @returns {Promise<Suggestion[]>}
   */
  async getCursorBackendSuggestions(context, parsed, requestSignal) {
    if (!this.completionClient) return [];
    if (!parsed.isFormula) return [];

    const input = safeToString(context?.currentInput);
    // Avoid asking the backend when the formula body is empty (e.g. "=" / "= ").
    // This keeps tab completion responsive and lets the curated starter functions
    // appear immediately without waiting for the network timeout.
    if (input.startsWith("=") && input.slice(1).trim() === "") {
      return [];
    }

    // Rule-based completions are often better than LLM for function names and
    // argument structure. Only ask the backend when we have an actual formula body.
    if (parsed.functionNamePrefix) return [];

    const cursor = clampCursor(input, context?.cursorPosition);
    const cell = safeNormalizeCellRef(context?.cellRef);

    const controller = new AbortController();
    const removeRequestAbortListener = forwardAbortSignal(requestSignal, controller);

    try {
      const completion = await withTimeout(
        this.completionClient.completeTabCompletion({
          input,
          cursorPosition: cursor,
          cellA1: safeToA1(cell),
          signal: controller.signal,
        }),
        this.completionTimeoutMs,
        () => controller.abort()
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
    } finally {
      removeRequestAbortListener?.();
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
   * @param {ReturnType<typeof parsePartialFormulaFallback>} parsed
   * @returns {Promise<Suggestion[]>}
   */
  suggestRangeCompletions(context, parsed) {
    const input = safeToString(context?.currentInput);
    const cursor = clampCursor(input, context?.cursorPosition);
    const cellRef = safeNormalizeCellRef(context?.cellRef);
    const fnSpec = parsed.functionName ? this.functionRegistry.getFunction(parsed.functionName) : undefined;
    const argIndex = parsed.argIndex ?? 0;
    const argSpecName = fnSpec?.args?.[argIndex]?.name;
    const functionCouldBeComplete = functionCouldBeCompleteAfterArg(fnSpec, argIndex);

    // When the user has typed a function call and the current argument is still
    // empty (e.g. "=SUM("), we can still offer a useful range suggestion by
    // assuming the current column as the prefix token.
    //
    // If we don't have a valid active cell ref (e.g. cursor is detached from a
    // cell), avoid suggesting ranges entirely.
    const hasValidCellRef = Number.isInteger(cellRef?.row) && Number.isInteger(cellRef?.col) && cellRef.row >= 0 && cellRef.col >= 0;
    if (!hasValidCellRef) return [];

    const typedArgText = parsed.currentArg.text ?? "";
    const isEmptyArg = typedArgText.trim().length === 0;
    const currentArgText = isEmptyArg ? columnIndexToLetter(cellRef.col) : typedArgText;

    const rangeCandidates = safeSuggestRanges({
      currentArgText,
      cellRef,
      surroundingCells: context?.surroundingCells,
    });

    // Some functions (VLOOKUP table_array, TAKE array, etc.) almost always want
    // a 2D rectangular range when the surrounding data forms a table. When we
    // have both a 1D and 2D candidate, slightly bias toward the 2D option so
    // tab completion defaults to the more useful table-shaped range.
    const prefersTableRange = argSpecName === "table_array" || argSpecName === "array" || argSpecName === "database";
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
          (candidate.reason === "contiguous_above_current_cell" ||
            candidate.reason === "contiguous_below_current_cell" ||
            candidate.reason === "contiguous_down_from_start")
        ) {
          confidence = clamp01(confidence - 0.1);
        }
      }

      // "Empty arg defaults" are helpful, but slightly less confident than when
      // the user has explicitly typed a prefix.
      if (isEmptyArg) {
        confidence = clamp01((confidence ?? 0) - 0.05);
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
        /** @type {Suggestion[]} */
        const merged = [...suggestions, ...schemaSuggestions];

        // If we couldn't generate a range completion (e.g. the user already typed a full
        // A1 range like "A1:A10"), still provide a useful pure-insertion suggestion by
        // auto-closing parens when the function could be complete.
        //
        // This keeps the common "=SUM(A1:A10<tab>" workflow working even when the range
        // suggester intentionally doesn't attempt to "re-suggest" already-complete ranges.
        if (
          merged.length === 0 &&
          cursor === input.length &&
          functionCouldBeComplete &&
          looksLikeCompleteA1RangeArg(typedArgText)
        ) {
          const closed = closeUnbalancedParens(input);
          if (closed !== input) {
            merged.push({
              text: closed,
              displayText: closed,
              type: "range",
              confidence: 0.25,
            });
          }
        }

        return rankAndDedupe(merged).slice(0, this.maxSuggestions);
      });
  }

  /**
   * Suggest named ranges, sheet-qualified ranges, and structured references for the current arg.
   *
   * @param {CompletionContext} context
   * @param {ReturnType<typeof parsePartialFormulaFallback>} parsed
   * @returns {Promise<Suggestion[]>}
   */
  async suggestSchemaRanges(context, parsed) {
    const provider = this.schemaProvider;
    if (!provider) return [];

    const input = safeToString(context?.currentInput);
    const cursor = clampCursor(input, context?.cursorPosition);
    const cellRef = safeNormalizeCellRef(context?.cellRef);

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
      // Avoid auto-closing parens for incomplete sheet-prefix-only suggestions
      // like "Sheet1!" (or obviously incomplete range prefixes like "A1:").
      if (cursor === input.length && functionCouldBeComplete && !isIncompleteRangeReplacement(replacement)) {
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
            const rangeCandidates = safeSuggestRanges({
              currentArgText: sheetArg.rangePrefix,
              cellRef,
              surroundingCells: context?.surroundingCells,
              sheetName,
            });
            // If the user already typed a complete range like Sheet2!A1:A10, rangeSuggester
            // intentionally returns no candidates (it focuses on *expanding* prefixes).
            // Still offer a useful completion by auto-closing parens when the function could
            // be complete.
            if (
              rangeCandidates.length === 0 &&
              cursor === input.length &&
              functionCouldBeComplete &&
              looksLikeCompleteA1RangeArg(sheetArg.rangePrefix) &&
              closeUnbalancedParens(input) !== input
            ) {
              addReplacement(rawPrefix, { confidence: 0.25 });
            }
            for (const candidate of rangeCandidates) {
              // Only emit suggestions that can be represented as a pure insertion at the
              // caret (i.e. candidate.range must start with the user-typed range prefix).
              //
              // Examples:
              // - Typed: "Sheet2!A1:"  -> Candidate: "A1:A10" (ok; insertion after ':')
              // - Typed: "Sheet2!A:"   -> Candidate: "A1:A10" (not ok; would need to insert before ':')
              //
              // When the candidate doesn't extend the typed prefix, don't emit it.
              if (typeof candidate?.range !== "string") continue;
              if (!candidate.range.startsWith(sheetArg.rangePrefix)) continue;

              const suffix = candidate.range.slice(sheetArg.rangePrefix.length);
              const replacement = `${rawPrefix}${suffix}`;

              // If the candidate doesn't actually extend the typed range prefix, only
              // keep it when it will still be useful by auto-closing unbalanced parens
              // (e.g. "=SUM(Sheet2!A:A" -> "=SUM(Sheet2!A:A)").
              if (suffix.length === 0) {
                const wouldAutoClose =
                  cursor === input.length && functionCouldBeComplete && !replacement.endsWith("!");
                if (!wouldAutoClose) continue;
                if (closeUnbalancedParens(input) === input) continue;
              }

              addReplacement(replacement, { confidence: Math.min(0.85, candidate.confidence + 0.05) });
            }
        }
      }
    } else if (rawPrefix) {
      // 4) Sheet-prefix completion: allow typing "=SUM(She" → "=SUM(Sheet1!"
      // (and for quoted sheet names: "=SUM('My" → "=SUM('My Sheet'!").
      const typedQuoted = rawPrefix.startsWith("'");
      for (const sheetName of sheetNames) {
        if (typeof sheetName !== "string" || sheetName.length === 0) continue;

        const requiresQuotes = needsSheetQuotes(sheetName);
        // Only suggest unquoted sheet names when they don't require quotes.
        if (!typedQuoted && requiresQuotes) continue;
        // Only suggest quoted sheet names when the user started a quote.
        if (typedQuoted && !requiresQuotes) continue;

        const formattedPrefix = formatSheetPrefix(sheetName);
        if (!startsWithIgnoreCase(formattedPrefix, rawPrefix)) continue;

        addReplacement(completeIdentifier(formattedPrefix, rawPrefix), {
          confidence: clamp01(0.6 + ratioBoost(rawPrefix, formattedPrefix) * 0.25),
        });
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

    // 3) Sheet identifiers (Sheet2!)
    const sheetNames = await safeProviderCall(provider.getSheetNames);
    for (const rawName of sheetNames) {
      const sheetName = typeof rawName === "string" ? rawName.trim() : "";
      if (!sheetName) continue;
      // Do not suggest quoted sheet references here. The desktop formula bar tab-complete
      // UI only supports "pure insertion" suggestions, and inserting leading quotes would
      // require modifying text before the cursor.
      if (needsSheetQuotes(sheetName)) continue;
      if (!startsWithIgnoreCase(sheetName, prefix)) continue;

      const completedName = completeIdentifier(sheetName, prefix);
      const replacement = `${completedName}!`;
      const insertedSuffix = replacement.slice(prefix.length);
      const newText = replaceSpan(input, token.start, token.end, replacement);
      suggestions.push({
        text: newText,
        // `displayText` is the part that would be inserted at the caret.
        displayText: insertedSuffix,
        type: "range",
        confidence: clamp01(0.55 + ratioBoost(prefix, sheetName) * 0.35),
      });
    }

    return rankAndDedupe(suggestions).slice(0, this.maxSuggestions);
  }

  /**
   * @param {CompletionContext} context
   * @param {ReturnType<typeof parsePartialFormulaFallback>} parsed
   * @returns {Suggestion[]}
   */
  suggestArgumentValues(context, parsed) {
    const input = safeToString(context?.currentInput);
    const cursor = clampCursor(input, context?.cursorPosition);
    const cellRef = safeNormalizeCellRef(context?.cellRef);

    const fnName = parsed.functionName;
    const argIndex = parsed.argIndex ?? 0;
    const argType = this.functionRegistry.getArgType(fnName, argIndex) ?? "any";

    const spanStart = parsed.currentArg?.start ?? cursor;
    const spanEnd = parsed.currentArg?.end ?? cursor;
    const typedPrefix = (parsed.currentArg?.text ?? "").trim();

    /** @type {Suggestion[]} */
    const suggestions = [];

    /**
     * @param {{ replacement: string, displayText?: string, confidence: number }} entry
     */
    const addReplacement = (entry) => {
      const replacement = entry.replacement;
      const displayText = entry.displayText ?? replacement;
      suggestions.push({
        text: replaceSpan(input, spanStart, spanEnd, replacement),
        displayText,
        type: "function_arg",
        confidence: clamp01(entry.confidence + prefixMatchBoost(typedPrefix, replacement)),
      });
    };

    if (argType === "boolean") {
      const enumEntries =
        getFunctionSpecificArgEnum(fnName, argIndex) ?? getCumulativeDistributionBooleanEnum(fnName, argIndex);
      if (enumEntries?.length) {
        for (const entry of enumEntries) addReplacement(entry);
        return dedupeSuggestions(suggestions);
      }

      for (const boolLiteral of ["TRUE", "FALSE"]) {
        addReplacement({ replacement: boolLiteral, confidence: 0.5 });
      }
      return dedupeSuggestions(suggestions);
    }

    if (argType === "number") {
      const enumEntries = getFunctionSpecificArgEnum(fnName, argIndex);
      if (enumEntries?.length) {
        for (const entry of enumEntries) addReplacement(entry);
        return dedupeSuggestions(suggestions);
      }

      for (const n of ["1", "0"]) {
        addReplacement({ replacement: n, confidence: 0.4 });
      }
      return dedupeSuggestions(suggestions);
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

  /**
   * When the user has only typed "=" (optionally followed by whitespace), we
   * provide a small curated set of common Excel functions so tab completion is
   * immediately useful without any additional typing.
   *
   * These suggestions must be representable as "pure insertions" at the caret
   * (the formula bar ghost text model), so we only trigger when the input has
   * no non-whitespace characters after the leading "=".
   *
   * @param {CompletionContext} context
   * @returns {Suggestion[]}
   */
  suggestTopLevelFunctions(context) {
    const input = context.currentInput ?? "";
    if (typeof input !== "string") return [];

    // Only trigger when the formula body is empty (e.g. "=", "= ", "\n=\t").
    if (input.trim() !== "=") return [];

    const cursor = clampCursor(input, context.cursorPosition);
    // The desktop formula bar completion UI only supports pure insertions at the
    // caret. Avoid suggesting these starter functions unless the caret is at
    // the end of the current input.
    if (cursor !== input.length) return [];

    // Curated list + stable ordering.
    //
    // Keep these just below the backend completion confidence so (when the backend
    // is available) users get the richer, context-aware suggestion instead of a
    // bare function stub. When the backend times out/unavailable, these still
    // provide useful immediate fallbacks.
    const starters = ["SUM(", "AVERAGE(", "IF(", "XLOOKUP(", "VLOOKUP(", "INDEX(", "MATCH("];

    /** @type {Suggestion[]} */
    const suggestions = [];
    for (let i = 0; i < starters.length; i += 1) {
      const displayText = starters[i];
      const text = replaceSpan(input, cursor, cursor, displayText);
      suggestions.push({
        text,
        displayText,
        type: "formula",
        confidence: 0.34 - i * 0.01,
      });
    }

    return suggestions.slice(0, this.maxSuggestions);
  }
}

/**
 * Function-specific enumerations for commonly misunderstood "flag" arguments.
 * These are curated because the function catalog only carries coarse arg types.
 *
 * @type {Record<string, Record<number, {replacement: string, displayText?: string, confidence: number}[]>>}
 */
const FUNCTION_SPECIFIC_ARG_ENUMS = {
  MATCH: {
    // match_type
    2: [
      { replacement: "0", displayText: "0 (exact match)", confidence: 0.72 },
      { replacement: "1", displayText: "1 (largest <= lookup_value)", confidence: 0.64 },
      { replacement: "-1", displayText: "-1 (smallest >= lookup_value)", confidence: 0.63 },
    ],
  },
  XMATCH: {
    // match_mode
    2: [
      { replacement: "0", displayText: "0 (exact match)", confidence: 0.72 },
      { replacement: "-1", displayText: "-1 (exact or next smaller)", confidence: 0.64 },
      { replacement: "1", displayText: "1 (exact or next larger)", confidence: 0.63 },
      { replacement: "2", displayText: "2 (wildcard match)", confidence: 0.6 },
    ],
    // search_mode
    3: [
      { replacement: "1", displayText: "1 (first-to-last)", confidence: 0.7 },
      { replacement: "-1", displayText: "-1 (last-to-first)", confidence: 0.64 },
      { replacement: "2", displayText: "2 (binary search ascending)", confidence: 0.61 },
      { replacement: "-2", displayText: "-2 (binary search descending)", confidence: 0.6 },
    ],
  },
  XLOOKUP: {
    // match_mode
    4: [
      { replacement: "0", displayText: "0 (exact match)", confidence: 0.72 },
      { replacement: "-1", displayText: "-1 (exact or next smaller)", confidence: 0.64 },
      { replacement: "1", displayText: "1 (exact or next larger)", confidence: 0.63 },
      { replacement: "2", displayText: "2 (wildcard match)", confidence: 0.6 },
    ],
    // search_mode
    5: [
      { replacement: "1", displayText: "1 (first-to-last)", confidence: 0.7 },
      { replacement: "-1", displayText: "-1 (last-to-first)", confidence: 0.64 },
      { replacement: "2", displayText: "2 (binary search ascending)", confidence: 0.61 },
      { replacement: "-2", displayText: "-2 (binary search descending)", confidence: 0.6 },
    ],
  },
  SORT: {
    // sort_order
    2: [
      { replacement: "1", displayText: "1 (ascending)", confidence: 0.66 },
      { replacement: "-1", displayText: "-1 (descending)", confidence: 0.65 },
    ],
    // by_col
    3: [
      { replacement: "FALSE", displayText: "FALSE (sort by rows)", confidence: 0.62 },
      { replacement: "TRUE", displayText: "TRUE (sort by columns)", confidence: 0.61 },
    ],
  },
  SORTBY: {
    // sort_order1
    2: [
      { replacement: "1", displayText: "1 (ascending)", confidence: 0.66 },
      { replacement: "-1", displayText: "-1 (descending)", confidence: 0.65 },
    ],
  },
  TAKE: {
    // rows
    1: [
      { replacement: "1", displayText: "1 (first row(s))", confidence: 0.62 },
      { replacement: "-1", displayText: "-1 (last row(s))", confidence: 0.61 },
    ],
    // columns
    2: [
      { replacement: "1", displayText: "1 (first column(s))", confidence: 0.6 },
      { replacement: "-1", displayText: "-1 (last column(s))", confidence: 0.59 },
    ],
  },
  DROP: {
    // rows
    1: [
      { replacement: "1", displayText: "1 (drop first row(s))", confidence: 0.62 },
      { replacement: "-1", displayText: "-1 (drop last row(s))", confidence: 0.61 },
    ],
    // columns
    2: [
      { replacement: "1", displayText: "1 (drop first column(s))", confidence: 0.6 },
      { replacement: "-1", displayText: "-1 (drop last column(s))", confidence: 0.59 },
    ],
  },
  TEXTSPLIT: {
    // ignore_empty
    3: [
      { replacement: "TRUE", displayText: "TRUE (ignore empty)", confidence: 0.66 },
      { replacement: "FALSE", displayText: "FALSE (keep empty)", confidence: 0.62 },
    ],
    // match_mode
    4: [
      { replacement: "0", displayText: "0 (case-sensitive)", confidence: 0.64 },
      { replacement: "1", displayText: "1 (case-insensitive)", confidence: 0.63 },
    ],
  },
  TEXTJOIN: {
    // ignore_empty
    1: [
      { replacement: "TRUE", displayText: "TRUE (ignore empty)", confidence: 0.66 },
      { replacement: "FALSE", displayText: "FALSE (keep empty)", confidence: 0.62 },
    ],
  },
  UNIQUE: {
    // by_col
    1: [
      { replacement: "FALSE", displayText: "FALSE (compare by rows)", confidence: 0.63 },
      { replacement: "TRUE", displayText: "TRUE (compare by columns)", confidence: 0.62 },
    ],
    // exactly_once
    2: [
      { replacement: "FALSE", displayText: "FALSE (include duplicates)", confidence: 0.62 },
      { replacement: "TRUE", displayText: "TRUE (only values occurring once)", confidence: 0.61 },
    ],
  },
  SUBTOTAL: {
    // function_num (1-11 / 101-111)
    0: [
      { replacement: "9", displayText: "9 (SUM)", confidence: 0.7 },
      { replacement: "109", displayText: "109 (SUM, ignore hidden)", confidence: 0.68 },
      { replacement: "3", displayText: "3 (COUNTA)", confidence: 0.66 },
      { replacement: "103", displayText: "103 (COUNTA, ignore hidden)", confidence: 0.65 },
      { replacement: "1", displayText: "1 (AVERAGE)", confidence: 0.64 },
    ],
  },
  AGGREGATE: {
    // function_num
    0: [
      { replacement: "9", displayText: "9 (SUM)", confidence: 0.7 },
      { replacement: "1", displayText: "1 (AVERAGE)", confidence: 0.67 },
      { replacement: "3", displayText: "3 (COUNTA)", confidence: 0.66 },
      { replacement: "4", displayText: "4 (MAX)", confidence: 0.65 },
      { replacement: "5", displayText: "5 (MIN)", confidence: 0.64 },
    ],
    // options
    1: [
      { replacement: "0", displayText: "0 (ignore nested SUBTOTAL/AGGREGATE)", confidence: 0.68 },
      { replacement: "4", displayText: "4 (ignore nothing)", confidence: 0.66 },
      { replacement: "6", displayText: "6 (ignore errors)", confidence: 0.64 },
      { replacement: "7", displayText: "7 (ignore hidden + errors)", confidence: 0.63 },
      { replacement: "3", displayText: "3 (ignore nested + hidden + errors)", confidence: 0.62 },
    ],
  },
  "CEILING.MATH": {
    // mode
    2: [
      { replacement: "0", displayText: "0 (default; negatives toward 0)", confidence: 0.64 },
      { replacement: "1", displayText: "1 (negatives away from 0)", confidence: 0.63 },
    ],
  },
  "FLOOR.MATH": {
    // mode
    2: [
      { replacement: "0", displayText: "0 (default; negatives away from 0)", confidence: 0.64 },
      { replacement: "1", displayText: "1 (negatives toward 0)", confidence: 0.63 },
    ],
  },
  QUARTILE: {
    // quart
    1: [
      { replacement: "1", displayText: "1 (1st quartile / 25%)", confidence: 0.66 },
      { replacement: "2", displayText: "2 (median / 50%)", confidence: 0.65 },
      { replacement: "3", displayText: "3 (3rd quartile / 75%)", confidence: 0.64 },
      { replacement: "0", displayText: "0 (minimum)", confidence: 0.62 },
      { replacement: "4", displayText: "4 (maximum)", confidence: 0.61 },
    ],
  },
  "QUARTILE.INC": {
    // quart
    1: [
      { replacement: "1", displayText: "1 (1st quartile / 25%)", confidence: 0.66 },
      { replacement: "2", displayText: "2 (median / 50%)", confidence: 0.65 },
      { replacement: "3", displayText: "3 (3rd quartile / 75%)", confidence: 0.64 },
      { replacement: "0", displayText: "0 (minimum)", confidence: 0.62 },
      { replacement: "4", displayText: "4 (maximum)", confidence: 0.61 },
    ],
  },
  "QUARTILE.EXC": {
    // quart
    1: [
      { replacement: "1", displayText: "1 (1st quartile / 25%)", confidence: 0.66 },
      { replacement: "2", displayText: "2 (median / 50%)", confidence: 0.65 },
      { replacement: "3", displayText: "3 (3rd quartile / 75%)", confidence: 0.64 },
      { replacement: "0", displayText: "0 (minimum)", confidence: 0.62 },
      { replacement: "4", displayText: "4 (maximum)", confidence: 0.61 },
    ],
  },
  "T.TEST": {
    // tails
    2: [
      { replacement: "1", displayText: "1 (one-tailed)", confidence: 0.68 },
      { replacement: "2", displayText: "2 (two-tailed)", confidence: 0.67 },
    ],
    // type
    3: [
      { replacement: "1", displayText: "1 (paired)", confidence: 0.67 },
      { replacement: "2", displayText: "2 (two-sample equal variance)", confidence: 0.66 },
      { replacement: "3", displayText: "3 (two-sample unequal variance)", confidence: 0.65 },
    ],
  },
  TTEST: {
    // tails
    2: [
      { replacement: "1", displayText: "1 (one-tailed)", confidence: 0.68 },
      { replacement: "2", displayText: "2 (two-tailed)", confidence: 0.67 },
    ],
    // type
    3: [
      { replacement: "1", displayText: "1 (paired)", confidence: 0.67 },
      { replacement: "2", displayText: "2 (two-sample equal variance)", confidence: 0.66 },
      { replacement: "3", displayText: "3 (two-sample unequal variance)", confidence: 0.65 },
    ],
  },
  "RANK.EQ": {
    // order
    2: [
      { replacement: "0", displayText: "0 (descending)", confidence: 0.66 },
      { replacement: "1", displayText: "1 (ascending)", confidence: 0.65 },
    ],
  },
  "RANK.AVG": {
    // order
    2: [
      { replacement: "0", displayText: "0 (descending)", confidence: 0.66 },
      { replacement: "1", displayText: "1 (ascending)", confidence: 0.65 },
    ],
  },
  RANK: {
    // order
    2: [
      { replacement: "0", displayText: "0 (descending)", confidence: 0.66 },
      { replacement: "1", displayText: "1 (ascending)", confidence: 0.65 },
    ],
  },
  WEEKDAY: {
    // return_type
    1: [
      { replacement: "1", displayText: "1 (Sun=1..Sat=7)", confidence: 0.66 },
      { replacement: "2", displayText: "2 (Mon=1..Sun=7)", confidence: 0.65 },
      { replacement: "3", displayText: "3 (Mon=0..Sun=6)", confidence: 0.64 },
    ],
  },
  WEEKNUM: {
    // return_type
    1: [
      { replacement: "1", displayText: "1 (week starts Sunday)", confidence: 0.66 },
      { replacement: "2", displayText: "2 (week starts Monday)", confidence: 0.65 },
      { replacement: "21", displayText: "21 (ISO week numbering)", confidence: 0.63 },
    ],
  },
  "FORECAST.ETS": {
    // seasonality
    3: [
      { replacement: "0", displayText: "0 (auto-detect seasonality)", confidence: 0.67 },
      { replacement: "1", displayText: "1 (no seasonality)", confidence: 0.66 },
      { replacement: "12", displayText: "12 (monthly seasonality)", confidence: 0.64 },
      { replacement: "4", displayText: "4 (quarterly seasonality)", confidence: 0.63 },
    ],
    // data_completion
    4: [
      { replacement: "1", displayText: "1 (interpolate missing points)", confidence: 0.67 },
      { replacement: "0", displayText: "0 (treat missing points as 0)", confidence: 0.66 },
    ],
    // aggregation
    5: [
      { replacement: "1", displayText: "1 (AVERAGE)", confidence: 0.67 },
      { replacement: "7", displayText: "7 (SUM)", confidence: 0.66 },
      { replacement: "2", displayText: "2 (COUNT)", confidence: 0.64 },
      { replacement: "4", displayText: "4 (MAX)", confidence: 0.63 },
      { replacement: "6", displayText: "6 (MIN)", confidence: 0.62 },
    ],
  },
  "FORECAST.ETS.CONFINT": {
    // confidence_level
    3: [
      { replacement: "0.95", displayText: "0.95 (95% confidence)", confidence: 0.67 },
      { replacement: "0.9", displayText: "0.9 (90% confidence)", confidence: 0.66 },
      { replacement: "0.99", displayText: "0.99 (99% confidence)", confidence: 0.65 },
    ],
    // seasonality
    4: [
      { replacement: "0", displayText: "0 (auto-detect seasonality)", confidence: 0.67 },
      { replacement: "1", displayText: "1 (no seasonality)", confidence: 0.66 },
      { replacement: "12", displayText: "12 (monthly seasonality)", confidence: 0.64 },
      { replacement: "4", displayText: "4 (quarterly seasonality)", confidence: 0.63 },
    ],
    // data_completion
    5: [
      { replacement: "1", displayText: "1 (interpolate missing points)", confidence: 0.67 },
      { replacement: "0", displayText: "0 (treat missing points as 0)", confidence: 0.66 },
    ],
    // aggregation
    6: [
      { replacement: "1", displayText: "1 (AVERAGE)", confidence: 0.67 },
      { replacement: "7", displayText: "7 (SUM)", confidence: 0.66 },
      { replacement: "2", displayText: "2 (COUNT)", confidence: 0.64 },
      { replacement: "4", displayText: "4 (MAX)", confidence: 0.63 },
      { replacement: "6", displayText: "6 (MIN)", confidence: 0.62 },
    ],
  },
  "FORECAST.ETS.SEASONALITY": {
    // data_completion
    2: [
      { replacement: "1", displayText: "1 (interpolate missing points)", confidence: 0.67 },
      { replacement: "0", displayText: "0 (treat missing points as 0)", confidence: 0.66 },
    ],
    // aggregation
    3: [
      { replacement: "1", displayText: "1 (AVERAGE)", confidence: 0.67 },
      { replacement: "7", displayText: "7 (SUM)", confidence: 0.66 },
      { replacement: "2", displayText: "2 (COUNT)", confidence: 0.64 },
      { replacement: "4", displayText: "4 (MAX)", confidence: 0.63 },
      { replacement: "6", displayText: "6 (MIN)", confidence: 0.62 },
    ],
  },
  LINEST: {
    // const (TRUE = calculate intercept, FALSE = force intercept=0)
    2: [
      { replacement: "TRUE", displayText: "TRUE (calculate intercept)", confidence: 0.66 },
      { replacement: "FALSE", displayText: "FALSE (force intercept=0)", confidence: 0.65 },
    ],
    // stats
    3: [
      { replacement: "TRUE", displayText: "TRUE (return regression stats)", confidence: 0.66 },
      { replacement: "FALSE", displayText: "FALSE (coefficients only)", confidence: 0.65 },
    ],
  },
  LOGEST: {
    // const
    2: [
      { replacement: "TRUE", displayText: "TRUE (calculate intercept)", confidence: 0.66 },
      { replacement: "FALSE", displayText: "FALSE (force intercept=0)", confidence: 0.65 },
    ],
    // stats
    3: [
      { replacement: "TRUE", displayText: "TRUE (return regression stats)", confidence: 0.66 },
      { replacement: "FALSE", displayText: "FALSE (coefficients only)", confidence: 0.65 },
    ],
  },
  TREND: {
    // const
    3: [
      { replacement: "TRUE", displayText: "TRUE (calculate intercept)", confidence: 0.66 },
      { replacement: "FALSE", displayText: "FALSE (force intercept=0)", confidence: 0.65 },
    ],
  },
  GROWTH: {
    // const
    3: [
      { replacement: "TRUE", displayText: "TRUE (calculate intercept)", confidence: 0.66 },
      { replacement: "FALSE", displayText: "FALSE (force intercept=0)", confidence: 0.65 },
    ],
  },
  DAYS360: {
    // method
    2: [
      { replacement: "FALSE", displayText: "FALSE (US/NASD method)", confidence: 0.66 },
      { replacement: "TRUE", displayText: "TRUE (European method)", confidence: 0.65 },
    ],
  },
  YEARFRAC: {
    // basis
    2: [
      { replacement: "0", displayText: "0 (US/NASD 30/360)", confidence: 0.66 },
      { replacement: "1", displayText: "1 (actual/actual)", confidence: 0.65 },
      { replacement: "2", displayText: "2 (actual/360)", confidence: 0.64 },
      { replacement: "3", displayText: "3 (actual/365)", confidence: 0.63 },
      { replacement: "4", displayText: "4 (European 30/360)", confidence: 0.62 },
    ],
  },
  VLOOKUP: {
    // range_lookup (TRUE = approx match, FALSE = exact)
    3: [
      { replacement: "FALSE", displayText: "FALSE (exact match)", confidence: 0.7 },
      { replacement: "TRUE", displayText: "TRUE (approximate match)", confidence: 0.69 },
    ],
  },
  HLOOKUP: {
    // range_lookup
    3: [
      { replacement: "FALSE", displayText: "FALSE (exact match)", confidence: 0.7 },
      { replacement: "TRUE", displayText: "TRUE (approximate match)", confidence: 0.69 },
    ],
  },
};

/**
 * @param {string | undefined} fnName
 * @param {number} argIndex
 */
function getFunctionSpecificArgEnum(fnName, argIndex) {
  if (typeof fnName !== "string" || !Number.isInteger(argIndex)) return null;
  const upper = fnName.toUpperCase();
  const base = upper.startsWith("_XLFN.") ? upper.slice("_XLFN.".length) : upper;
  return FUNCTION_SPECIFIC_ARG_ENUMS[base]?.[argIndex] ?? null;
}

/**
 * Heuristic for the common `cumulative` boolean argument in distribution functions.
 *
 * Many Excel functions named `*.DIST`/`*DIST` accept a final boolean arg that toggles
 * cumulative vs PDF/PMF mode. The core catalog only exposes "boolean" so we provide
 * a more informative hint without needing per-function curation.
 *
 * @param {string | undefined} fnName
 * @param {number} argIndex
 */
function getCumulativeDistributionBooleanEnum(fnName, argIndex) {
  if (typeof fnName !== "string" || !Number.isInteger(argIndex)) return null;
  const upper = fnName.toUpperCase();
  const base = upper.startsWith("_XLFN.") ? upper.slice("_XLFN.".length) : upper;

  // Match distribution function families:
  // - Modern: NORM.DIST, NORM.S.DIST, LOGNORM.DIST, F.DIST, T.DIST, ...
  // - Legacy: NORMDIST, NORMSDIST, BINOMDIST, ...
  const looksLikeDist = base.includes(".DIST") || base.endsWith("DIST");
  if (!looksLikeDist) return null;

  // Most distributions use the boolean arg as `cumulative`. This is a best-effort
  // hint and should stay conservative (only 2 suggestions).
  return [
    { replacement: "TRUE", displayText: "TRUE (cumulative)", confidence: 0.66 },
    { replacement: "FALSE", displayText: "FALSE (probability)", confidence: 0.64 },
  ];
}

/**
 * Prefer the value that matches the typed prefix without changing overall ordering too much.
 * @param {string} typedPrefix
 * @param {string} replacement
 */
function prefixMatchBoost(typedPrefix, replacement) {
  if (!typedPrefix) return 0;
  if (!replacement) return 0;
  const typed = typedPrefix.trim();
  if (!typed) return 0;
  return startsWithIgnoreCase(replacement, typed) ? 0.05 : 0;
}

/**
 * Dedupe while preserving the first (highest-priority) entry.
 * The overall suggestion list is still globally ranked/deduped later, but keeping
 * this bounded avoids pushing generic suggestions out of the `maxSuggestions` cap.
 *
 * @param {Suggestion[]} suggestions
 */
function dedupeSuggestions(suggestions) {
  /** @type {Set<string>} */
  const seen = new Set();
  /** @type {Suggestion[]} */
  const out = [];
  for (const s of suggestions) {
    const key = s?.text;
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(s);
  }
  return out;
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

function looksLikeCompleteA1RangeArg(text) {
  if (typeof text !== "string") return false;
  const trimmed = text.trim();
  if (!trimmed) return false;
  // Only trigger for fully-specified 2-cell A1 ranges like A1:A10 (possibly with $ markers).
  // This intentionally excludes sheet-qualified refs and structured references.
  return /^(\$?[A-Za-z]{1,3})(\$?\d+):(\$?[A-Za-z]{1,3})(\$?\d+)$/.test(trimmed);
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
      bestByText.set(s.text, {
        ...s,
        displayText: typeof s.displayText === "string" && s.displayText.length > 0 ? s.displayText : s.text,
      });
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
    const aText = typeof a.displayText === "string" ? a.displayText : a.text;
    const bText = typeof b.displayText === "string" ? b.displayText : b.text;
    return aText.localeCompare(bText);
  });
}

function clamp01(v) {
  return Math.max(0, Math.min(1, v));
}

/**
 * Returns true when the replacement looks incomplete and should not trigger
 * auto-closing parens.
 *
 * @param {string} replacement
 */
function isIncompleteRangeReplacement(replacement) {
  if (typeof replacement !== "string" || replacement.length === 0) return false;
  const trimmed = replacement.trimEnd();
  if (trimmed.endsWith("!")) return true;
  // Optional: don't auto-close for obviously incomplete range tokens.
  if (trimmed.endsWith(":")) return true;
  return false;
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
 * @param {(() => void) | undefined} onTimeout
 * @returns {Promise<T>}
 */
function withTimeout(promise, timeoutMs, onTimeout) {
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) return promise;

  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      try {
        onTimeout?.();
      } catch {
        // ignore
      }
      reject(new Error("timeout"));
    }, timeoutMs);

    Promise.resolve(promise).then(
      (value) => {
        clearTimeout(timeout);
        resolve(value);
      },
      (err) => {
        clearTimeout(timeout);
        reject(err);
      },
    );
  });
}

/**
 * Forward an external AbortSignal onto a local AbortController so callers can cancel
 * in-flight backend completions (e.g. when the user keeps typing).
 *
 * @param {AbortSignal | undefined} requestSignal
 * @param {AbortController} controller
 * @returns {(() => void) | null}
 */
function forwardAbortSignal(requestSignal, controller) {
  /** @type {(() => void) | null} */
  let removeListener = null;
  if (!requestSignal) return removeListener;

  if (requestSignal.aborted) {
    controller.abort();
    return removeListener;
  }

  if (typeof requestSignal.addEventListener === "function") {
    const onAbort = () => controller.abort();
    requestSignal.addEventListener("abort", onAbort, { once: true });
    removeListener = () => requestSignal.removeEventListener("abort", onAbort);
  }

  return removeListener;
}

const DEFAULT_CELL_REF = Object.freeze({ row: 0, col: 0 });

/**
 * @param {any} value
 * @returns {string}
 */
function safeToString(value) {
  try {
    if (typeof value === "string") return value;
    return String(value ?? "");
  } catch {
    return "";
  }
}

/**
 * @param {any} cellRef
 * @returns {{row:number,col:number}}
 */
function safeNormalizeCellRef(cellRef) {
  try {
    const normalized = normalizeCellRef(cellRef);
    const row = Number.isInteger(normalized?.row) && normalized.row >= 0 ? normalized.row : 0;
    const col = Number.isInteger(normalized?.col) && normalized.col >= 0 ? normalized.col : 0;
    return { row, col };
  } catch {
    return { row: 0, col: 0 };
  }
}

/**
 * @param {{row:number,col:number}} cellRef
 * @returns {string}
 */
function safeToA1(cellRef) {
  try {
    return toA1(cellRef);
  } catch {
    return "A1";
  }
}

/**
 * @param {any} cache
 * @param {string} key
 */
function safeCacheGet(cache, key) {
  try {
    return cache?.get?.(key);
  } catch {
    return undefined;
  }
}

/**
 * @param {any} cache
 * @param {string} key
 * @param {any} value
 */
function safeCacheSet(cache, key, value) {
  try {
    cache?.set?.(key, value);
  } catch {
    // ignore
  }
}

/**
 * @template T
 * @param {() => T | Promise<T>} fn
 * @returns {Promise<T[]>}
 */
async function safeArrayResult(fn) {
  try {
    const result = await Promise.resolve().then(fn);
    return Array.isArray(result) ? result : [];
  } catch {
    return [];
  }
}

function safeSuggestRanges(params) {
  try {
    const out = suggestRanges(params);
    return Array.isArray(out) ? out : [];
  } catch {
    return [];
  }
}
