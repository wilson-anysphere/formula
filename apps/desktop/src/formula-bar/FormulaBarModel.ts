import { getActiveArgumentFrame, getActiveArgumentSpan } from "./highlight/activeArgument.js";
import { getFunctionSignature, signatureParts } from "./highlight/functionSignatures.js";
import { parseA1Range, rangeToA1, type RangeAddress } from "../spreadsheet/a1.js";
import { formatSheetNameForA1 } from "../sheet/formatSheetNameForA1.js";
import {
  assignFormulaReferenceColors,
  extractFormulaReferences,
  extractFormulaReferencesFromTokens,
  type ColoredFormulaReference,
  type ExtractFormulaReferencesOptions,
  type FormulaReference,
  type FormulaReferenceRange,
} from "@formula/spreadsheet-frontend";
import {
  tokenizeFormula,
  type FormulaToken,
  type FormulaTokenType,
} from "@formula/spreadsheet-frontend/formula/tokenizeFormula";
import type {
  FormulaParseError,
  FormulaPartialLexResult,
  FormulaPartialParseResult,
  FormulaSpan,
  FormulaToken as EngineFormulaToken,
  FunctionContext as EngineFunctionContext,
} from "@formula/engine";
import { normalizeFormulaLocaleId } from "../spreadsheet/formulaLocale.js";

type ActiveCellInfo = {
  address: string;
  input: string;
  value: unknown;
};

type FormulaBarAiSuggestion = {
  text: string;
  preview?: unknown;
};

type ErrorExplanation = {
  code: string;
  title: string;
  description: string;
  suggestions: string[];
};

type FunctionSignature = NonNullable<ReturnType<typeof getFunctionSignature>>;

type FunctionHint = {
  context: { name: string; argIndex: number };
  signature: FunctionSignature;
  parts: Array<{ text: string; kind: "name" | "param" | "paramActive" | "punct" }>;
};

type HighlightSpan = {
  kind: FormulaTokenType;
  text: string;
  start: number;
  end: number;
  /**
   * Optional CSS class applied to the rendered <span>.
   *
   * Used by the WASM-backed editor tooling integration to surface parse errors
   * with an exact span highlight.
   */
  className?: string;
};

type HighlightSpanLike = HighlightSpan | FormulaToken;

const RESOLVED_REFERENCE_TEXT_CACHE_LIMIT = 200;

export class FormulaBarModel {
  #activeCell: ActiveCellInfo = { address: "A1", input: "", value: "" };
  #draft: string = "";
  #draftVersion = 0;
  #isEditing = false;
  #cursorStart = 0;
  #cursorEnd = 0;
  #tokenCache: { draft: string; tokens: FormulaToken[] } | null = null;
  #activeArgumentSpanCache:
    | { draft: string; cursor: number; argSeparator: string; result: ReturnType<typeof getActiveArgumentSpan> }
    | null = null;
  #activeArgumentSpanCache2:
    | { draft: string; cursor: number; argSeparator: string; result: ReturnType<typeof getActiveArgumentSpan> }
    | null = null;
  #activeArgumentFrameCache:
    | { draft: string; cursor: number; argSeparator: string; result: ReturnType<typeof getActiveArgumentFrame> }
    | null = null;
  #activeArgumentFrameCache2:
    | { draft: string; cursor: number; argSeparator: string; result: ReturnType<typeof getActiveArgumentFrame> }
    | null = null;
  #extractFormulaReferencesOptions: ExtractFormulaReferencesOptions | null = null;
  #extractFormulaReferencesOptionsVersion = 0;
  #referenceExtractionCache:
    | {
        draft: string;
        optionsVersion: number;
        references: FormulaReference[];
      }
    | null = null;
  #resolvedReferenceTextCache:
    | { optionsVersion: number; entries: Map<string, RangeAddress | null> }
    | null = null;
  #rangeInsertion: { start: number; end: number } | null = null;
  #hoveredReference: RangeAddress | null = null;
  #hoveredReferenceText: string | null = null;
  #referenceColorByText = new Map<string, string>();
  #coloredReferences: ColoredFormulaReference[] = [];
  #activeReferenceIndex: number | null = null;
  #referenceHighlightsCache:
    | Array<{ range: FormulaReferenceRange; color: string; text: string; index: number; active: boolean }>
    | null = null;
  #referenceHighlightsCacheRefs: ColoredFormulaReference[] | null = null;
  #referenceHighlightsCacheActiveIndex: number | null = null;
  #engineHighlightSpans: HighlightSpan[] | null = null;
  #engineLexTokens: EngineFormulaToken[] | null = null;
  #engineHighlightErrorSpanStart: number | null = null;
  #engineHighlightErrorSpanEnd: number | null = null;
  #engineHighlightRefs: ColoredFormulaReference[] | null = null;
  #engineFunctionContext: EngineFunctionContext | null = null;
  #engineSyntaxError: FormulaParseError | null = null;
  #engineToolingFormula: string | null = null;
  #engineToolingLocaleId: string = "en-US";
  #errorExplanationCache: { value: unknown; result: ErrorExplanation | null } | null = null;
  #functionHintCache:
    | {
        draft: string;
        localeId: string;
        argSeparator: string;
        name: string;
        argIndex: number;
        hint: FunctionHint;
      }
    | null = null;
  /**
   * Full text suggestion for the current draft (not just the "ghost text" tail).
   *
   * Prefer setting this to the full suggested draft text (so we can derive a
   * "ghost text" tail via `aiGhostText()` and apply it with a single replace).
   *
   * For compatibility, `setAiSuggestion()` also accepts just the ghost/insertion
   * text to be inserted at the caret (e.g. "M"), and normalizes it into a full
   * suggestion string.
   */
  #aiSuggestion: string | null = null;
  #aiSuggestionIsFullDraft = false;
  #aiSuggestionPreview: unknown | null = null;
  #aiGhostCache:
    | { suggestion: string; isFullDraft: boolean; draftVersion: number; cursor: number; ghost: string }
    | null = null;

  setActiveCell(info: ActiveCellInfo): void {
    this.#activeCell = { ...info };
    this.#draft = info.input ?? "";
    this.#draftVersion += 1;
    this.#isEditing = false;
    this.#cursorStart = this.#draft.length;
    this.#cursorEnd = this.#draft.length;
    this.#tokenCache = null;
    this.#clearActiveArgumentSpanCache();
    this.#rangeInsertion = null;
    this.#hoveredReference = null;
    this.#hoveredReferenceText = null;
    this.#referenceColorByText.clear();
    this.#coloredReferences = [];
    this.#activeReferenceIndex = null;
    this.#referenceHighlightsCache = null;
    this.#referenceHighlightsCacheRefs = null;
    this.#referenceHighlightsCacheActiveIndex = null;
    this.#clearEditorTooling();
    this.#errorExplanationCache = null;
    this.#functionHintCache = null;
    this.#aiSuggestion = null;
    this.#aiSuggestionIsFullDraft = false;
    this.#aiSuggestionPreview = null;
    this.#aiGhostCache = null;
    this.#resolvedReferenceTextCache = null;
  }

  /**
   * Provide an optional name -> range resolver so formula reference highlights can
   * include named ranges (identifiers that are not A1-style refs).
   */
  setNameResolver(resolver: ((name: string) => FormulaReferenceRange | null) | null): void {
    const prev = this.#extractFormulaReferencesOptions;
    const next: ExtractFormulaReferencesOptions = { ...(prev ?? {}) };
    next.resolveName = resolver ?? undefined;
    this.setExtractFormulaReferencesOptions(
      next.resolveName || next.resolveStructuredRef || next.tables ? next : null
    );
  }

  setExtractFormulaReferencesOptions(opts: ExtractFormulaReferencesOptions | null): void {
    this.#extractFormulaReferencesOptions = opts;
    this.#extractFormulaReferencesOptionsVersion += 1;
    // Clear any cached reference tokenization so we re-extract with the new options.
    this.#referenceExtractionCache = null;
    this.#resolvedReferenceTextCache = null;
    if (this.#isEditing) {
      this.#updateReferenceHighlights();
      this.#updateHoverFromCursor();
    }
  }

  extractFormulaReferencesOptions(): ExtractFormulaReferencesOptions | null {
    return this.#extractFormulaReferencesOptions;
  }

  /**
   * Best-effort named-range resolver for view-mode hover previews.
   *
   * The formula tokenizer used for syntax highlighting represents named ranges as
   * `identifier` tokens, so hover logic cannot rely on `reference` tokens alone.
   */
  resolveNameRange(name: string): FormulaReferenceRange | null {
    const resolver = this.#extractFormulaReferencesOptions?.resolveName;
    if (!resolver) return null;
    const lower = String(name ?? "").toLowerCase();
    if (lower === "true" || lower === "false") return null;
    return resolver(name);
  }

  /**
   * Best-effort resolver for a single reference token text (A1 or structured ref).
   *
   * This is primarily used by view-mode hover previews, which operate on the rendered
   * syntax-highlight spans rather than cursor-aware reference extraction.
   *
   * - A1 refs (including sheet-qualified refs like `Sheet2!A1:B2`) are handled via
   *   `parseSheetQualifiedA1Range`.
   * - Structured refs (e.g. `Table1[Amount]`) are resolved using the configured
   *   `ExtractFormulaReferencesOptions` (tables / resolveStructuredRef), if any.
   */
  resolveReferenceText(text: string): RangeAddress | null {
    const raw = String(text ?? "");
    if (!raw) return null;
    const lastIndex = raw.length - 1;
    const trimmed = isWhitespaceChar(raw[0] ?? "") || isWhitespaceChar(raw[lastIndex] ?? "") ? raw.trim() : raw;
    if (!trimmed) return null;

    const currentOptionsVersion = this.#extractFormulaReferencesOptionsVersion;
    let cache = this.#resolvedReferenceTextCache;
    if (!cache || cache.optionsVersion !== currentOptionsVersion) {
      cache = { optionsVersion: currentOptionsVersion, entries: new Map() };
      this.#resolvedReferenceTextCache = cache;
    }
    if (cache.entries.has(trimmed)) {
      return cache.entries.get(trimmed) ?? null;
    }

    const a1 = parseSheetQualifiedA1Range(trimmed);
    if (a1) {
      cache.entries.set(trimmed, a1);
      if (cache.entries.size > RESOLVED_REFERENCE_TEXT_CACHE_LIMIT) {
        const oldest = cache.entries.keys().next().value;
        if (typeof oldest === "string") cache.entries.delete(oldest);
      }
      return a1;
    }

    const { references } = extractFormulaReferences(trimmed, undefined, undefined, this.#extractFormulaReferencesOptions ?? undefined);
    const first = references[0];
    let resolved: RangeAddress | null = null;
    if (first && first.start === 0 && first.end === trimmed.length) {
      const r = first.range;
      resolved = { start: { row: r.startRow, col: r.startCol }, end: { row: r.endRow, col: r.endCol } };
    }

    cache.entries.set(trimmed, resolved);
    if (cache.entries.size > RESOLVED_REFERENCE_TEXT_CACHE_LIMIT) {
      // Map iteration order reflects insertion order; evict the oldest entry to keep this cache bounded.
      const oldest = cache.entries.keys().next().value;
      if (typeof oldest === "string") cache.entries.delete(oldest);
    }

    return resolved;
  }

  get activeCell(): ActiveCellInfo {
    return this.#activeCell;
  }

  get isEditing(): boolean {
    return this.#isEditing;
  }

  get draft(): string {
    return this.#draft;
  }

  get draftVersion(): number {
    return this.#draftVersion;
  }

  get cursorStart(): number {
    return this.#cursorStart;
  }

  get cursorEnd(): number {
    return this.#cursorEnd;
  }

  beginEdit(): void {
    this.#isEditing = true;
    this.#clearEditorTooling();
    this.#clearActiveArgumentSpanCache();
    this.#updateReferenceHighlights();
    this.#updateHoverFromCursor();
  }

  updateDraft(draft: string, cursorStart: number, cursorEnd: number): void {
    const nextCursorStart = Math.max(0, Math.min(cursorStart, draft.length));
    const nextCursorEnd = Math.max(0, Math.min(cursorEnd, draft.length));
    const draftChanged = draft !== this.#draft;
    const cursorChanged = nextCursorStart !== this.#cursorStart || nextCursorEnd !== this.#cursorEnd;
    if (this.#isEditing && !draftChanged && !cursorChanged) return;

    this.#isEditing = true;
    this.#draft = draft;
    if (draftChanged) {
      this.#draftVersion += 1;
    }
    this.#cursorStart = nextCursorStart;
    this.#cursorEnd = nextCursorEnd;
    if (this.#rangeInsertion != null) {
      this.#rangeInsertion = null;
    }
    if (draftChanged) {
      // Draft text changed; any cached engine tokens/spans are now stale.
      this.#clearEditorTooling();
      this.#tokenCache = null;
      this.#clearActiveArgumentSpanCache();
      this.#functionHintCache = null;
    } else if (cursorChanged) {
      // Cursor moved within the same draft; keep lex-based highlights/syntax errors but
      // clear cursor-dependent parse context so hint rendering can refresh.
      this.#engineFunctionContext = null;
    }
    if (this.#aiSuggestion != null) {
      this.#aiSuggestion = null;
      this.#aiSuggestionIsFullDraft = false;
      this.#aiGhostCache = null;
    }
    if (this.#aiSuggestionPreview != null) {
      this.#aiSuggestionPreview = null;
    }
    this.#updateReferenceHighlights();
    this.#updateHoverFromCursor();
  }

  commit(): string {
    this.#isEditing = false;
    this.#activeCell = { ...this.#activeCell, input: this.#draft };
    this.#rangeInsertion = null;
    this.#clearEditorTooling();
    this.#tokenCache = null;
    this.#clearActiveArgumentSpanCache();
    this.#aiSuggestion = null;
    this.#aiSuggestionIsFullDraft = false;
    this.#aiSuggestionPreview = null;
    this.#aiGhostCache = null;
    this.#hoveredReference = null;
    this.#hoveredReferenceText = null;
    this.#referenceColorByText.clear();
    this.#coloredReferences = [];
    this.#activeReferenceIndex = null;
    this.#referenceHighlightsCache = null;
    this.#referenceHighlightsCacheRefs = null;
    this.#referenceHighlightsCacheActiveIndex = null;
    return this.#draft;
  }

  cancel(): void {
    this.#isEditing = false;
    const nextDraft = this.#activeCell.input;
    if (nextDraft !== this.#draft) {
      this.#draft = nextDraft;
      this.#draftVersion += 1;
    } else {
      this.#draft = nextDraft;
    }
    this.#cursorStart = this.#draft.length;
    this.#cursorEnd = this.#draft.length;
    this.#rangeInsertion = null;
    this.#clearEditorTooling();
    this.#tokenCache = null;
    this.#clearActiveArgumentSpanCache();
    this.#aiSuggestion = null;
    this.#aiSuggestionIsFullDraft = false;
    this.#aiSuggestionPreview = null;
    this.#aiGhostCache = null;
    this.#hoveredReference = null;
    this.#hoveredReferenceText = null;
    this.#referenceColorByText.clear();
    this.#coloredReferences = [];
    this.#activeReferenceIndex = null;
    this.#referenceHighlightsCache = null;
    this.#referenceHighlightsCacheRefs = null;
    this.#referenceHighlightsCacheActiveIndex = null;
  }

  highlightedSpans(): HighlightSpanLike[] {
    if (this.#engineToolingFormula === this.#draft && this.#engineHighlightSpans) {
      return this.#engineHighlightSpans;
    }
    return this.#tokensForDraft();
  }

  #tokensForDraft(): FormulaToken[] {
    const cache = this.#tokenCache;
    if (cache && cache.draft === this.#draft) return cache.tokens;
    const tokens = tokenizeFormula(this.#draft);
    this.#tokenCache = { draft: this.#draft, tokens };
    return tokens;
  }

  functionHint(): FunctionHint | null {
    const hasEngineLocaleForDraft = this.#engineToolingFormula === this.#draft;
    const rawLocaleId = hasEngineLocaleForDraft
      ? this.#engineToolingLocaleId
      : (typeof document !== "undefined" ? document.documentElement?.lang : "")?.trim?.() || "en-US";
    // Function hint semantics (localized function names + argument separators) follow the formula
    // engine's supported locale set. Treat unsupported/variant locale IDs as their canonical
    // formula locale ids to keep caching stable.
    const localeId = normalizeFormulaLocaleId(rawLocaleId) ?? "en-US";
    const argSeparator = inferArgSeparator(localeId);
    const argSeparatorChar = argSeparator[0] === ";" ? ";" : ",";

    let ctxName: string | null = null;
    let ctxArgIndex: number | null = null;

    if (hasEngineLocaleForDraft && this.#engineFunctionContext) {
      const ctx = this.#engineFunctionContext;
      ctxName = ctx.name;
      ctxArgIndex = ctx.argIndex;
    } else {
      let active = this.#activeArgumentFrame(this.#cursorStart, argSeparatorChar);
      ctxName = active?.fnName ?? null;
      ctxArgIndex = active?.argIndex ?? null;

      // Excel UX: keep showing the innermost function hint when the caret is just
      // after a closing paren (e.g. "=ROUND(1,2)|"). The simple tokenizer-based
      // parser considers the call "closed" once it consumes ')', which would
      // otherwise clear the hint.
      if (ctxName == null && this.#cursorStart === this.#cursorEnd && this.#cursorStart > 0) {
        let scan = this.#cursorStart - 1;
        while (scan >= 0 && isWhitespaceChar(this.#draft[scan] ?? "")) scan -= 1;
        if (scan >= 0 && this.#draft[scan] === ")") {
          active = this.#activeArgumentFrame(scan, argSeparatorChar);
          ctxName = active?.fnName ?? null;
          ctxArgIndex = active?.argIndex ?? null;
        }
      }
    }

    if (ctxName == null || ctxArgIndex == null) return null;

    const cache = this.#functionHintCache;
    if (
      cache &&
      cache.draft === this.#draft &&
      cache.localeId === localeId &&
      cache.argSeparator === argSeparator &&
      cache.name === ctxName &&
      cache.argIndex === ctxArgIndex
    ) {
      return cache.hint;
    }

    const signature = getFunctionSignature(ctxName, { localeId });
    if (!signature) return null;
    const hint: FunctionHint = {
      context: { name: ctxName, argIndex: ctxArgIndex },
      signature,
      parts: signatureParts(signature, ctxArgIndex, { argSeparator }),
    };
    this.#functionHintCache = { draft: this.#draft, localeId, argSeparator, name: ctxName, argIndex: ctxArgIndex, hint };
    return hint;
  }

  errorExplanation(): ErrorExplanation | null {
    const value = this.#activeCell.value;
    const cache = this.#errorExplanationCache;
    if (cache && cache.value === value) return cache.result;
    const result = explainFormulaError(value);
    this.#errorExplanationCache = { value, result };
    return result;
  }

  syntaxError(): FormulaParseError | null {
    if (this.#engineToolingFormula !== this.#draft) return null;
    const err = this.#engineSyntaxError;
    if (!err) return null;
    // The partial parser sometimes reports "expected token" errors with a zero-length span at
    // the cursor (incomplete input). Only surface errors that correspond to a highlightable span.
    if (!err.span || err.span.end <= err.span.start) return null;
    return err;
  }

  /**
   * Apply editor tooling results from the WASM engine (`lexFormulaPartial` + `parseFormulaPartial`).
   *
   * The caller (FormulaBarView) is responsible for debouncing and stale response filtering.
   */
  applyEngineToolingResult(args: {
    formula: string;
    localeId: string;
    lexResult: FormulaPartialLexResult;
    parseResult: FormulaPartialParseResult;
  }): void {
    // Only accept tooling results that match the current draft.
    if (args.formula !== this.#draft) return;

    this.#engineToolingFormula = args.formula;
    this.#engineToolingLocaleId = args.localeId || "en-US";
    this.#engineFunctionContext = args.parseResult.context.function;

    // Prefer lexer errors (unexpected characters, unterminated strings) over parse errors.
    // The parser may also report an error in these cases, but the lexer message/span tends
    // to be more precise for editor feedback.
    const error = args.lexResult.error ?? args.parseResult.error;
    this.#engineSyntaxError = error;

    const errorSpan = error?.span ?? null;
    const highlightableErrorSpan = errorSpan && errorSpan.end > errorSpan.start ? errorSpan : null;
    const errorSpanStart = highlightableErrorSpan?.start ?? null;
    const errorSpanEnd = highlightableErrorSpan?.end ?? null;

    const highlightStable =
      this.#engineHighlightSpans != null &&
      this.#engineLexTokens === args.lexResult.tokens &&
      this.#engineHighlightErrorSpanStart === errorSpanStart &&
      this.#engineHighlightErrorSpanEnd === errorSpanEnd &&
      this.#engineHighlightRefs === this.#coloredReferences;

    if (!highlightStable) {
      // Reuse the already-computed reference extraction metadata so applying engine results
      // does not re-tokenize the full formula string on the main thread.
      const referenceTokens = this.#coloredReferences;
      const engineSpans = highlightFromEngineTokens(args.formula, args.lexResult.tokens);
      const withRefs = spliceReferenceSpans(args.formula, engineSpans, referenceTokens);
      const withError = applyErrorSpan(args.formula, withRefs, highlightableErrorSpan);
      this.#engineHighlightSpans = withError;
      this.#engineLexTokens = args.lexResult.tokens;
      this.#engineHighlightErrorSpanStart = errorSpanStart;
      this.#engineHighlightErrorSpanEnd = errorSpanEnd;
      this.#engineHighlightRefs = this.#coloredReferences;
    }
  }

  setHoveredReference(referenceText: string | null): void {
    if (!referenceText) {
      this.#hoveredReference = null;
      this.#hoveredReferenceText = null;
      return;
    }
    this.#hoveredReferenceText = referenceText;
    this.#hoveredReference = this.resolveReferenceText(referenceText);
  }

  hoveredReference(): RangeAddress | null {
    return this.#hoveredReference;
  }

  hoveredReferenceText(): string | null {
    return this.#hoveredReferenceText;
  }

  coloredReferences(): readonly ColoredFormulaReference[] {
    return this.#coloredReferences;
  }

  referenceHighlights(): Array<{ range: FormulaReferenceRange; color: string; text: string; index: number; active: boolean }> {
    const refs = this.#coloredReferences;
    const activeIndex = this.#activeReferenceIndex;

    const cachedRefs = this.#referenceHighlightsCacheRefs;
    let cached = this.#referenceHighlightsCache;

    // When the formula text changes, rebuild the highlight list from the colored reference metadata.
    if (!cached || cachedRefs !== refs) {
      cached = new Array(refs.length);
      for (let i = 0; i < refs.length; i += 1) {
        const ref = refs[i]!;
        cached[i] = {
          range: ref.range,
          color: ref.color,
          text: ref.text,
          index: ref.index,
          active: activeIndex === ref.index,
        };
      }
      this.#referenceHighlightsCache = cached;
      this.#referenceHighlightsCacheRefs = refs;
      this.#referenceHighlightsCacheActiveIndex = activeIndex;
      return cached;
    }

    // Cursor moves within the same formula are common; update only the active flag.
    const prevActive = this.#referenceHighlightsCacheActiveIndex;
    if (prevActive !== activeIndex) {
      if (prevActive != null) {
        const prev = cached[prevActive];
        if (prev) prev.active = false;
      }
      if (activeIndex != null) {
        const next = cached[activeIndex];
        if (next) next.active = true;
      }
      this.#referenceHighlightsCacheActiveIndex = activeIndex;
    }

    return cached;
  }

  activeReferenceIndex(): number | null {
    return this.#activeReferenceIndex;
  }

  beginRangeSelection(range: RangeAddress, sheetId?: string): void {
    if (!this.#isEditing) return;
    const active =
      this.#activeReferenceIndex == null ? null : (this.#coloredReferences[this.#activeReferenceIndex] ?? null);
    this.#insertOrReplaceRange(
      formatRangeReference(range, sheetId),
      true,
      active ? { start: active.start, end: active.end } : null
    );
    this.#aiSuggestion = null;
    this.#aiSuggestionIsFullDraft = false;
    this.#aiSuggestionPreview = null;
    this.#aiGhostCache = null;
    this.#updateReferenceHighlights();
    this.#updateHoverFromCursor();
  }

  updateRangeSelection(range: RangeAddress, sheetId?: string): void {
    if (!this.#isEditing) return;
    const rangeText = formatRangeReference(range, sheetId);
    const insertion = this.#rangeInsertion;
    if (
      insertion &&
      this.#cursorStart === insertion.end &&
      this.#cursorEnd === insertion.end &&
      insertion.end - insertion.start === rangeText.length &&
      this.#draft.startsWith(rangeText, insertion.start)
    ) {
      // Some range selection providers can emit multiple updates with identical ranges.
      // Avoid re-tokenizing / re-highlighting the full draft when nothing changed.
      if (this.#aiSuggestion != null) {
        this.#aiSuggestion = null;
        this.#aiSuggestionIsFullDraft = false;
        this.#aiGhostCache = null;
      }
      if (this.#aiSuggestionPreview != null) this.#aiSuggestionPreview = null;
      return;
    }

    this.#insertOrReplaceRange(rangeText, false);
    this.#aiSuggestion = null;
    this.#aiSuggestionIsFullDraft = false;
    this.#aiSuggestionPreview = null;
    this.#aiGhostCache = null;
    this.#updateReferenceHighlights();
    this.#updateHoverFromCursor();
  }

  endRangeSelection(): void {
    this.#rangeInsertion = null;
    this.#aiSuggestion = null;
    this.#aiSuggestionIsFullDraft = false;
    this.#aiSuggestionPreview = null;
    this.#aiGhostCache = null;
  }

  setAiSuggestion(suggestion: string | FormulaBarAiSuggestion | null): void {
    const suggestionText = typeof suggestion === "string" ? suggestion : suggestion?.text;
    const preview = typeof suggestion === "string" ? null : (suggestion?.preview ?? null);

    if (!suggestionText) {
      this.#aiSuggestion = null;
      this.#aiSuggestionIsFullDraft = false;
      this.#aiSuggestionPreview = null;
      this.#aiGhostCache = null;
      return;
    }

    // Normalize "tail" suggestions ("M") into full-string suggestions ("=SUM")
    // so `aiGhostText()` + `acceptAiSuggestion()` can operate consistently.
    if (!this.#isEditing) {
      this.#aiSuggestion = suggestionText;
      this.#aiSuggestionIsFullDraft = false;
      this.#aiSuggestionPreview = preview;
      this.#aiGhostCache = null;
      return;
    }

    const start = Math.min(this.#cursorStart, this.#cursorEnd);
    const end = Math.max(this.#cursorStart, this.#cursorEnd);
    const draft = this.#draft;
    const looksLikeFullSuggestion = matchesDraftFrame(draft, start, end, suggestionText);
    this.#aiSuggestion = looksLikeFullSuggestion ? suggestionText : draft.slice(0, start) + suggestionText + draft.slice(end);
    this.#aiSuggestionIsFullDraft = true;
    this.#aiSuggestionPreview = preview;
    this.#aiGhostCache = null;
  }

  aiSuggestion(): string | null {
    return this.#aiSuggestion;
  }

  aiSuggestionPreview(): unknown | null {
    return this.#aiSuggestionPreview;
  }

  aiGhostText(): string {
    if (!this.#aiSuggestion) {
      this.#aiGhostCache = null;
      return "";
    }
    if (this.#cursorStart !== this.#cursorEnd) {
      this.#aiGhostCache = null;
      return "";
    }

    const cursor = this.#cursorStart;
    const cache = this.#aiGhostCache;
    if (
      cache &&
      cache.suggestion === this.#aiSuggestion &&
      cache.isFullDraft === this.#aiSuggestionIsFullDraft &&
      cache.draftVersion === this.#draftVersion &&
      cache.cursor === cursor
    ) {
      return cache.ghost;
    }

    let ghost = "";
    if (this.#aiSuggestionIsFullDraft) {
      const suffixLen = this.#draft.length - cursor;
      const end = this.#aiSuggestion.length - suffixLen;
      if (end >= cursor) {
        ghost = this.#aiSuggestion.slice(cursor, end);
      }
    } else {
      const prefix = this.#draft.slice(0, cursor);
      const suffix = this.#draft.slice(cursor);
      if (this.#aiSuggestion.startsWith(prefix) && (!suffix || this.#aiSuggestion.endsWith(suffix))) {
        ghost = this.#aiSuggestion.slice(cursor, this.#aiSuggestion.length - suffix.length);
      }
    }

    this.#aiGhostCache = {
      suggestion: this.#aiSuggestion,
      isFullDraft: this.#aiSuggestionIsFullDraft,
      draftVersion: this.#draftVersion,
      cursor,
      ghost,
    };
    return ghost;
  }

  acceptAiSuggestion(): boolean {
    if (!this.#aiSuggestion) return false;
    if (!this.#isEditing) return false;

    const suggestionText = this.#aiSuggestion;
    const start = Math.min(this.#cursorStart, this.#cursorEnd);
    const end = Math.max(this.#cursorStart, this.#cursorEnd);

    const draft = this.#draft;
    const suffixLen = draft.length - end;

    // The completion engine typically supplies the full suggested text, but some
    // surfaces may pass only the "ghost" tail (text to insert at the caret).
    const looksLikeFullReplacement = this.#aiSuggestionIsFullDraft || matchesDraftFrame(draft, start, end, suggestionText);

    if (looksLikeFullReplacement) {
      const ghost = this.aiGhostText();
      this.#draft = suggestionText;
      this.#draftVersion += 1;
      const newCursor = ghost ? start + ghost.length : suggestionText.length - suffixLen;
      this.#cursorStart = newCursor;
      this.#cursorEnd = newCursor;
    } else {
      this.#draft = draft.slice(0, start) + suggestionText + draft.slice(end);
      this.#draftVersion += 1;
      const newCursor = start + suggestionText.length;
      this.#cursorStart = newCursor;
      this.#cursorEnd = newCursor;
    }
    this.#aiSuggestion = null;
    this.#aiSuggestionIsFullDraft = false;
    this.#aiSuggestionPreview = null;
    this.#aiGhostCache = null;
    this.#rangeInsertion = null;
    this.#updateReferenceHighlights();
    this.#updateHoverFromCursor();
    return true;
  }

  #insertOrReplaceRange(rangeText: string, isBegin: boolean, replaceSpan: { start: number; end: number } | null = null): void {
    if (!this.#rangeInsertion || isBegin) {
      const start = replaceSpan ? replaceSpan.start : Math.min(this.#cursorStart, this.#cursorEnd);
      const end = replaceSpan ? replaceSpan.end : Math.max(this.#cursorStart, this.#cursorEnd);
      this.#draft = this.#draft.slice(0, start) + rangeText + this.#draft.slice(end);
      this.#draftVersion += 1;
      this.#rangeInsertion = { start, end: start + rangeText.length };
      this.#cursorStart = this.#rangeInsertion.end;
      this.#cursorEnd = this.#rangeInsertion.end;
      return;
    }

    const { start, end } = this.#rangeInsertion;
    this.#draft = this.#draft.slice(0, start) + rangeText + this.#draft.slice(end);
    this.#draftVersion += 1;
    this.#rangeInsertion = { start, end: start + rangeText.length };
    this.#cursorStart = this.#rangeInsertion.end;
    this.#cursorEnd = this.#rangeInsertion.end;
  }

  #updateHoverFromCursor(): void {
    if (!this.#isEditing || !isFormulaText(this.#draft)) {
      if (this.#hoveredReference != null) this.#hoveredReference = null;
      if (this.#hoveredReferenceText != null) this.#hoveredReferenceText = null;
      return;
    }

    const activeIndex = this.#activeReferenceIndex;
    if (activeIndex == null) {
      if (this.#hoveredReference != null) this.#hoveredReference = null;
      if (this.#hoveredReferenceText != null) this.#hoveredReferenceText = null;
      return;
    }

    // Reuse the already-computed reference metadata to avoid re-tokenizing the full
    // formula string on every cursor move.
    const active = this.#coloredReferences[activeIndex] ?? null;
    if (!active) {
      if (this.#hoveredReference != null) this.#hoveredReference = null;
      if (this.#hoveredReferenceText != null) this.#hoveredReferenceText = null;
      return;
    }

    const nextText = active.text;
    const nextStartRow = active.range.startRow;
    const nextStartCol = active.range.startCol;
    const nextEndRow = active.range.endRow;
    const nextEndCol = active.range.endCol;

    const prevText = this.#hoveredReferenceText;
    const prevRange = this.#hoveredReference;
    const sameRange =
      prevRange != null &&
      prevRange.start.row === nextStartRow &&
      prevRange.start.col === nextStartCol &&
      prevRange.end.row === nextEndRow &&
      prevRange.end.col === nextEndCol;

    // When only the caret moves within the same active reference, the hover range/text
    // do not change. Avoid allocating new range objects in that hot path.
    if (prevText === nextText && sameRange) return;

    this.#hoveredReferenceText = nextText;
    this.#hoveredReference = { start: { row: nextStartRow, col: nextStartCol }, end: { row: nextEndRow, col: nextEndCol } };

    // NOTE: `active.text` can be a sheet-qualified A1 reference, a named range, or
    // (when configured) a structured reference. Consumers that need sheet gating
    // should use `hoveredReferenceText()` (and resolve names/sheet qualifiers there).
  }

  #updateReferenceHighlights(): void {
    if (!this.#isEditing || !isFormulaText(this.#draft)) {
      // Cursor moves can still call into this method (e.g. while editing plain text).
      // Avoid allocating fresh empty arrays / resetting caches when we're already in the
      // "no reference highlights" state.
      if (this.#coloredReferences.length !== 0) this.#coloredReferences = [];
      if (this.#activeReferenceIndex != null) this.#activeReferenceIndex = null;
      if (this.#referenceHighlightsCache != null) this.#referenceHighlightsCache = null;
      if (this.#referenceHighlightsCacheRefs != null) this.#referenceHighlightsCacheRefs = null;
      if (this.#referenceHighlightsCacheActiveIndex != null) this.#referenceHighlightsCacheActiveIndex = null;
      return;
    }

    const cache = this.#referenceExtractionCache;
    let references: FormulaReference[];
    const cacheHit =
      cache != null && cache.draft === this.#draft && cache.optionsVersion === this.#extractFormulaReferencesOptionsVersion;
    if (cacheHit) {
      references = cache.references;
    } else {
      const tokens = this.#tokensForDraft();
      references = extractFormulaReferencesFromTokens(tokens, this.#draft, this.#extractFormulaReferencesOptions ?? undefined);
      this.#referenceExtractionCache = {
        draft: this.#draft,
        optionsVersion: this.#extractFormulaReferencesOptionsVersion,
        references,
      };
    }

    // Cursor moves within the same reference token are extremely common (especially while
    // the user edits/steps through a ref like `A1:B2`). Avoid the O(log n) binary search in
    // that hot path by checking the previously-active reference first.
    const prevActiveIndex = this.#activeReferenceIndex;
    let activeIndex: number | null | undefined = undefined;
    if (cacheHit && prevActiveIndex != null && this.#coloredReferences.length === references.length) {
      const prevRef = this.#coloredReferences[prevActiveIndex] ?? null;
      if (prevRef) {
        const start = Math.min(this.#cursorStart, this.#cursorEnd);
        const end = Math.max(this.#cursorStart, this.#cursorEnd);
        if (start !== end) {
          if (prevRef.start <= start && end <= prevRef.end) activeIndex = prevRef.index;
        } else {
          // Caret: treat the reference as active if the caret is inside (including at end).
          if (prevRef.start <= start && start <= prevRef.end) activeIndex = prevRef.index;
        }
      }
    }
    if (activeIndex === undefined) {
      activeIndex = findActiveReferenceIndex(references, this.#cursorStart, this.#cursorEnd);
    }

    // Reference colors and ranges are determined by the formula text. When only the cursor/selection
    // moves, we can keep the existing colored reference list and update just the active index.
    if (!cacheHit || this.#coloredReferences.length !== references.length) {
      const { colored, nextByText } = assignFormulaReferenceColors(references, this.#referenceColorByText);
      this.#referenceColorByText = nextByText;
      this.#coloredReferences = colored;
    }

    this.#activeReferenceIndex = activeIndex;
  }

  #clearEditorTooling(): void {
    this.#engineHighlightSpans = null;
    this.#engineLexTokens = null;
    this.#engineHighlightErrorSpanStart = null;
    this.#engineHighlightErrorSpanEnd = null;
    this.#engineHighlightRefs = null;
    this.#engineFunctionContext = null;
    this.#engineSyntaxError = null;
    this.#engineToolingFormula = null;
  }

  #activeArgumentFrame(cursorIndex: number, argSeparator: string): ReturnType<typeof getActiveArgumentFrame> {
    const draft = this.#draft;
    const cursor = Math.max(0, Math.min(cursorIndex, draft.length));

    const c1 = this.#activeArgumentFrameCache;
    if (c1 && c1.draft === draft && c1.argSeparator === argSeparator) {
      if (c1.cursor === cursor) return c1.result;
      const cached = c1.result;
      if (cached && cursor >= cached.span.start && cursor <= cached.span.end) {
        c1.cursor = cursor;
        return cached;
      }
    }
    const c2 = this.#activeArgumentFrameCache2;
    if (c2 && c2.draft === draft && c2.argSeparator === argSeparator) {
      if (c2.cursor === cursor) {
        this.#activeArgumentFrameCache2 = c1;
        this.#activeArgumentFrameCache = c2;
        return c2.result;
      }
      const cached = c2.result;
      if (cached && cursor >= cached.span.start && cursor <= cached.span.end) {
        c2.cursor = cursor;
        this.#activeArgumentFrameCache2 = c1;
        this.#activeArgumentFrameCache = c2;
        return cached;
      }
    }

    const result = getActiveArgumentFrame(draft, cursor, { argSeparators: argSeparator });
    this.#activeArgumentFrameCache2 = c1;
    this.#activeArgumentFrameCache = { draft, cursor, argSeparator, result };
    return result;
  }

  activeArgumentSpan(cursorIndex: number = this.#cursorStart): ReturnType<typeof getActiveArgumentSpan> {
    const draft = this.#draft;
    const cursor = Math.max(0, Math.min(cursorIndex, draft.length));

    const hasEngineLocaleForDraft = this.#engineToolingFormula === draft;
    const localeId = hasEngineLocaleForDraft
      ? this.#engineToolingLocaleId
      : (typeof document !== "undefined" ? document.documentElement?.lang : "")?.trim?.() || "en-US";
    const argSeparatorText = inferArgSeparator(localeId);
    // `inferArgSeparator` returns strings like "; " / ", " (with trailing space).
    // Avoid calling `.trim()` (allocates) in this hot path; the separator is always the first character.
    const argSeparator = argSeparatorText[0] === ";" ? ";" : ",";

    const c1 = this.#activeArgumentSpanCache;
    if (c1 && c1.draft === draft && c1.argSeparator === argSeparator) {
      if (c1.cursor === cursor) return c1.result;
      // Cursor moves within the same argument span are extremely common while editing.
      // `getActiveArgumentSpan` is O(cursor) in the worst case, so reuse the prior result
      // when the cursor is still within that span.
      const cached = c1.result;
      if (cached && cursor >= cached.span.start && cursor <= cached.span.end) {
        c1.cursor = cursor;
        return cached;
      }
    }
    const c2 = this.#activeArgumentSpanCache2;
    if (c2 && c2.draft === draft && c2.argSeparator === argSeparator) {
      if (c2.cursor === cursor) {
        // Promote to the front of the small LRU.
        this.#activeArgumentSpanCache2 = c1;
        this.#activeArgumentSpanCache = c2;
        return c2.result;
      }
      const cached = c2.result;
      if (cached && cursor >= cached.span.start && cursor <= cached.span.end) {
        c2.cursor = cursor;
        this.#activeArgumentSpanCache2 = c1;
        this.#activeArgumentSpanCache = c2;
        return cached;
      }
    }

    const result = getActiveArgumentSpan(draft, cursor, { argSeparators: argSeparator });
    this.#activeArgumentSpanCache2 = c1;
    this.#activeArgumentSpanCache = { draft, cursor, argSeparator, result };
    return result;
  }

  #clearActiveArgumentSpanCache(): void {
    this.#activeArgumentSpanCache = null;
    this.#activeArgumentSpanCache2 = null;
    this.#activeArgumentFrameCache = null;
    this.#activeArgumentFrameCache2 = null;
  }
}

function isWhitespaceChar(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

function matchesDraftFrame(draft: string, start: number, end: number, suggestion: string): boolean {
  const startIndex = Math.max(0, Math.min(start, draft.length));
  const endIndex = Math.max(startIndex, Math.min(end, draft.length));
  if (suggestion.length < startIndex) return false;

  for (let i = 0; i < startIndex; i += 1) {
    if (draft.charCodeAt(i) !== suggestion.charCodeAt(i)) return false;
  }

  const suffixLen = draft.length - endIndex;
  if (suffixLen === 0) return true;

  const suggestionSuffixStart = suggestion.length - suffixLen;
  if (suggestionSuffixStart < startIndex) return false;

  for (let i = 0; i < suffixLen; i += 1) {
    if (draft.charCodeAt(endIndex + i) !== suggestion.charCodeAt(suggestionSuffixStart + i)) return false;
  }

  return true;
}

function isFormulaText(text: string): boolean {
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i] ?? "";
    if (isWhitespaceChar(ch)) continue;
    return ch === "=";
  }
  return false;
}

function findActiveReferenceIndex(
  references: readonly Pick<FormulaReference, "start" | "end" | "index">[],
  cursorStart: number,
  cursorEnd: number
): number | null {
  const start = Math.min(cursorStart, cursorEnd);
  const end = Math.max(cursorStart, cursorEnd);

  if (references.length === 0) return null;

  // If text is selected, treat a reference as active only when the selection is contained
  // within that reference token.
  if (start !== end) {
    return findContainingReference(references, start, end);
  }

  // Caret: treat the reference containing either the character at the caret or
  // immediately before it as active. This matches typical editor behavior where
  // being at the end of a token still counts as "in" the token.
  let active = findContainingReference(references, start, start + 1);
  if (active != null) return active;
  if (start > 0) {
    active = findContainingReference(references, start - 1, start);
    if (active != null) return active;
  }
  return null;
}

function findContainingReference(
  references: readonly Pick<FormulaReference, "start" | "end" | "index">[],
  needleStart: number,
  needleEnd: number
): number | null {
  // `extractFormulaReferences` returns references ordered by appearance in the formula,
  // so we can locate the active ref in O(log n) time using binary search.
  let lo = 0;
  let hi = references.length - 1;
  let candidate = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const ref = references[mid]!;
    if (ref.start <= needleStart) {
      candidate = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  if (candidate < 0) return null;
  const ref = references[candidate]!;
  if (ref.start <= needleStart && needleEnd <= ref.end) return ref.index;
  return null;
}

function formatRangeReference(range: RangeAddress, sheetId?: string): string {
  const a1 = rangeToA1(range);
  if (!sheetId) return a1;
  return `${formatSheetPrefix(sheetId)}${a1}`;
}

function formatSheetPrefix(id: string): string {
  const name = formatSheetNameForA1(id);
  return name ? `${name}!` : "";
}

function parseSheetQualifiedA1Range(text: string): RangeAddress | null {
  const { ref } = splitSheetQualifier(text);
  return parseA1Range(ref);
}

function splitSheetQualifier(input: string): { sheetName: string | null; ref: string } {
  // Keep this local to avoid importing the full `packages/search` barrel (which re-exports the
  // entire search session implementation) just to parse `Sheet1!A1`-style references.
  const raw = String(input ?? "");
  if (!raw) return { sheetName: null, ref: "" };
  const lastIndex = raw.length - 1;
  const s = isWhitespaceChar(raw[0] ?? "") || isWhitespaceChar(raw[lastIndex] ?? "") ? raw.trim() : raw;

  const quoted = s.match(/^'((?:[^']|'')+)'!(.+)$/);
  if (quoted) {
    return {
      sheetName: quoted[1].replace(/''/g, "'"),
      ref: quoted[2],
    };
  }

  const unquoted = s.match(/^([^!]+)!(.+)$/);
  if (unquoted) {
    return { sheetName: unquoted[1] ?? null, ref: unquoted[2] ?? "" };
  }

  return { sheetName: null, ref: s };
}

const ARG_SEPARATOR_CACHE = new Map<string, string>();

function inferArgSeparator(localeId: string): string {
  // Prefer the formula engine's normalized locale IDs so UI arg separators stay consistent
  // with parsing semantics (e.g. `de-CH` is currently treated as `de-DE` by the engine).
  const locale = normalizeFormulaLocaleId(localeId) ?? "en-US";
  const cached = ARG_SEPARATOR_CACHE.get(locale);
  if (cached) return cached;

  try {
    // Excel typically uses `;` as the list/arg separator when the decimal separator is `,`.
    // Infer this using Intl rather than hardcoding locale tables.
    const parts = new Intl.NumberFormat(locale).formatToParts(1.1);
    const decimal = parts.find((p) => p.type === "decimal")?.value ?? ".";
    const sep = decimal === "," ? "; " : ", ";
    ARG_SEPARATOR_CACHE.set(locale, sep);
    return sep;
  } catch {
    return ", ";
  }
}

const ERROR_EXPLANATIONS: Record<string, Omit<ErrorExplanation, "code">> = {
  "#GETTING_DATA": {
    title: "Loading",
    description: "This cell is waiting for an async result (for example, an AI function response).",
    suggestions: ["Wait a moment for the result to arrive.", "If it never resolves, check your AI settings or network connection."],
  },
  "#DLP!": {
    title: "Blocked by data loss prevention",
    description: "This AI function call was blocked by your organization's DLP policy.",
    suggestions: [
      "Remove or change references to restricted cells/ranges.",
      "If this should be allowed, ask an admin to adjust the document/org DLP policy.",
    ],
  },
  "#AI!": {
    title: "AI error",
    description: "The AI function failed to run (model unavailable, network error, or unexpected response).",
    suggestions: ["Check your AI provider/model settings.", "Try again in a moment."],
  },
  "#NULL!": {
    title: "Null intersection",
    description: "The formula tried to reference an intersection that doesn't exist.",
    suggestions: ["Check that the referenced ranges actually intersect.", "Verify the formula’s range operators and separators."],
  },
  "#CALC!": {
    title: "Calculation error",
    description: "The formula couldn’t be calculated (often due to an unsupported or invalid operation).",
    suggestions: ["Check inputs for invalid values.", "Simplify the formula to isolate the failing part."],
  },
  "#FIELD!": {
    title: "Invalid field",
    description: "The formula referenced a field that doesn’t exist (often in data types or external data).",
    suggestions: ["Verify the field name exists.", "Refresh or re-import the underlying data."],
  },
  "#CONNECT!": {
    title: "Connection error",
    description: "The formula depends on external data that couldn’t be reached.",
    suggestions: ["Check your network connection.", "Try refreshing the data source."],
  },
  "#BLOCKED!": {
    title: "Blocked",
    description: "The formula result was blocked (for example, by a permission or data restriction).",
    suggestions: ["Check document permissions and data restrictions.", "Try moving the formula or adjusting inputs."],
  },
  "#UNKNOWN!": {
    title: "Unknown error",
    description: "The formula returned an unknown error.",
    suggestions: ["Try recalculating or re-entering the formula.", "If it persists, report a bug with the workbook."],
  },
  "#DIV/0!": {
    title: "Division by zero",
    description: "The formula tried to divide by zero (or an empty cell).",
    suggestions: ["Check the divisor cell for a 0 or blank value.", "Wrap the division in IFERROR to provide a fallback value."],
  },
  "#NAME?": {
    title: "Unknown name",
    description: "The formula contains a function name or named range that isn’t recognized.",
    suggestions: ["Check the spelling of function names.", "Verify that referenced named ranges exist."],
  },
  "#REF!": {
    title: "Invalid reference",
    description: "The formula refers to a cell or range that no longer exists.",
    suggestions: ["Check for deleted rows/columns in referenced ranges.", "Update the formula to point to valid cells."],
  },
  "#VALUE!": {
    title: "Wrong type of value",
    description: "The formula used a value of the wrong type (e.g. text where a number was expected).",
    suggestions: ["Check referenced cells for unexpected text values.", "Use VALUE or other coercion helpers if needed."],
  },
  "#N/A": {
    title: "Value not available",
    description: "A lookup didn’t find a matching value (or data is missing).",
    suggestions: ["Verify the lookup value exists in the lookup range.", "Consider IFNA/IFERROR to handle missing values."],
  },
  "#NUM!": {
    title: "Invalid number",
    description: "The formula produced an invalid numeric result (too large/small or not representable).",
    suggestions: ["Check for invalid inputs (like negative numbers where not allowed).", "Simplify the calculation to avoid overflow."],
  },
  "#SPILL!": {
    title: "Spill range blocked",
    description: "A dynamic array formula can’t spill because cells in the spill area are not empty.",
    suggestions: ["Clear the cells where the formula needs to spill.", "Move the formula to an empty area."],
  },
};

function explainFormulaError(value: unknown): ErrorExplanation | null {
  if (typeof value !== "string") return null;
  const explanation = ERROR_EXPLANATIONS[value];
  if (!explanation) return null;
  return { code: value, ...explanation };
}

function highlightFormula(input: string): HighlightSpan[] {
  const tokens = tokenizeFormula(input);
  const spans = new Array<HighlightSpan>(tokens.length);
  for (let i = 0; i < tokens.length; i += 1) {
    const token = tokens[i]!;
    spans[i] = { kind: token.type, text: token.text, start: token.start, end: token.end };
  }
  return spans;
}

function highlightFromEngineTokens(formula: string, tokens: EngineFormulaToken[]): HighlightSpan[] {
  // In the steady state, engine tokens are already sorted by span start and include a single trailing `Eof`.
  // Avoid allocating a filtered/copy array on every keystroke; only fall back to sorting when we detect
  // out-of-order spans.
  const optimistic = highlightFromEngineTokensSorted(formula, tokens, true);
  if (optimistic) return optimistic;

  const ordered = tokens.filter((t) => t.kind !== "Eof").sort((a, b) => a.span.start - b.span.start);
  return highlightFromEngineTokensSorted(formula, ordered, false) ?? [];
}

function highlightFromEngineTokensSorted(
  formula: string,
  tokens: readonly EngineFormulaToken[],
  abortOnUnsorted: boolean
): HighlightSpan[] | null {
  const spans: HighlightSpan[] = [];
  let pos = 0;
  // Best-effort: treat an identifier immediately followed by "(" (ignoring whitespace) as a function name.
  // Instead of building a separate index (Set) with a second pass, mutate the already-emitted identifier
  // span when we later observe the "(" token.
  let lastNonWhitespaceSpanIndex: number | null = null;
  let lastTokenStart = -Infinity;

  for (const token of tokens) {
    if (token.kind === "Eof") continue;
    if (abortOnUnsorted) {
      const start = token.span.start;
      if (start < lastTokenStart) return null;
      lastTokenStart = start;
    }
    const start = Math.max(0, Math.min(token.span.start, formula.length));
    const end = Math.max(0, Math.min(token.span.end, formula.length));
    if (end < start) continue;

    if (start > pos) {
      // The engine lexer intentionally omits some characters (e.g. leading `=` in formula inputs).
      // Highlight any uncovered gaps using the local tokenizer as a best-effort fallback.
      const gapText = formula.slice(pos, start);
      for (const gapSpan of highlightFormula(gapText)) {
        const gapStart = pos + gapSpan.start;
        const gapEnd = pos + gapSpan.end;
        if (gapEnd <= gapStart) continue;
        spans.push({
          kind: gapSpan.kind,
          start: gapStart,
          end: gapEnd,
          text: formula.slice(gapStart, gapEnd),
        });
        if (gapSpan.kind !== "whitespace") {
          lastNonWhitespaceSpanIndex = spans.length - 1;
        }
      }
    }

    if (token.kind === "LParen" && lastNonWhitespaceSpanIndex != null) {
      const prev = spans[lastNonWhitespaceSpanIndex];
      if (prev?.kind === "identifier") {
        prev.kind = "function";
      }
    }

    if (end > start) {
      spans.push({
        kind: engineTokenKindToHighlightKind(token),
        start,
        end,
        text: formula.slice(start, end),
      });
      const spanKind = spans[spans.length - 1]!.kind;
      if (spanKind !== "whitespace") {
        lastNonWhitespaceSpanIndex = spans.length - 1;
      }
    }
    pos = Math.max(pos, end);
  }

  if (pos < formula.length) {
    const gapText = formula.slice(pos);
    for (const gapSpan of highlightFormula(gapText)) {
      const gapStart = pos + gapSpan.start;
      const gapEnd = pos + gapSpan.end;
      if (gapEnd <= gapStart) continue;
      spans.push({
        kind: gapSpan.kind,
        start: gapStart,
        end: gapEnd,
        text: formula.slice(gapStart, gapEnd),
      });
    }
  }

  return spans;
}

function engineTokenKindToHighlightKind(token: EngineFormulaToken): HighlightSpan["kind"] {
  switch (token.kind) {
    case "Whitespace":
      return "whitespace";
    case "Number":
      return "number";
    case "String":
      return "string";
    case "Boolean":
      return "identifier";
    case "Error":
      return "error";
    case "Cell":
    case "R1C1Cell":
    case "R1C1Row":
    case "R1C1Col":
      return "reference";
    case "Ident":
    case "QuotedIdent":
      return "identifier";
    case "Plus":
    case "Minus":
    case "Star":
    case "Slash":
    case "Caret":
    case "Amp":
    case "Percent":
    case "Hash":
    case "Eq":
    case "Ne":
    case "Lt":
    case "Gt":
    case "Le":
    case "Ge":
    case "At":
    case "Union":
    case "Intersect":
      return "operator";
    case "LParen":
    case "RParen":
    case "LBrace":
    case "RBrace":
    case "LBracket":
    case "RBracket":
    case "Bang":
    case "Colon":
    case "Dot":
    case "ArgSep":
    case "ArrayRowSep":
    case "ArrayColSep":
      return "punctuation";
    default:
      return "unknown";
  }
}

function spliceReferenceSpans(
  formula: string,
  spans: HighlightSpan[],
  referenceTokens: Array<{ start: number; end: number }>
): HighlightSpan[] {
  if (referenceTokens.length === 0) return spans;

  // In FormulaBarModel, `referenceTokens` comes from `extractFormulaReferencesFromTokens`, which
  // produces spans in appearance order. Avoid a redundant sortedness scan/copy on the hot path.
  const refs: ReadonlyArray<{ start: number; end: number }> = referenceTokens;
  const out: HighlightSpan[] = [];
  let pos = 0;
  let spanIndex = 0;

  const pushSpan = (span: HighlightSpan): void => {
    const prev = out[out.length - 1];
    if (prev && prev.end === span.start && prev.kind === span.kind && (prev.className ?? "") === (span.className ?? "")) {
      prev.end = span.end;
      prev.text += span.text;
      return;
    }
    out.push(span);
  };

  const emitSlice = (sliceStart: number, sliceEnd: number): void => {
    if (sliceEnd <= sliceStart) return;
    while (spanIndex < spans.length && spans[spanIndex]!.end <= sliceStart) spanIndex += 1;
    while (spanIndex < spans.length) {
      const span = spans[spanIndex]!;
      if (span.start >= sliceEnd) break;
      const s = Math.max(span.start, sliceStart);
      const e = Math.min(span.end, sliceEnd);
      if (e > s) {
        pushSpan({
          kind: span.kind,
          start: s,
          end: e,
          text: formula.slice(s, e),
          className: span.className,
        });
      }
      if (span.end <= sliceEnd) {
        spanIndex += 1;
        continue;
      }
      break;
    }
  };

  for (const ref of refs) {
    let start = Math.max(0, Math.min(ref.start, formula.length));
    const end = Math.max(0, Math.min(ref.end, formula.length));
    if (end <= start) continue;

    // Defensive: `referenceTokens` should be non-overlapping and sorted, but clamp to
    // a monotonic position so we never emit out-of-order spans.
    if (start < pos) start = pos;

    if (start > pos) {
      emitSlice(pos, start);
    }
    pushSpan({ kind: "reference", start, end, text: formula.slice(start, end) });
    pos = Math.max(pos, end);
  }

  if (pos < formula.length) {
    emitSlice(pos, formula.length);
  }

  return out;
}

function applyErrorSpan(formula: string, spans: HighlightSpan[], errorSpan: FormulaSpan | null): HighlightSpan[] {
  if (!errorSpan) return spans;
  const start = Math.max(0, Math.min(errorSpan.start, formula.length));
  const end = Math.max(0, Math.min(errorSpan.end, formula.length));
  if (end <= start) return spans;

  const out: HighlightSpan[] = [];

  const pushSpan = (span: HighlightSpan): void => {
    const prev = out[out.length - 1];
    if (prev && prev.end === span.start && prev.kind === span.kind && (prev.className ?? "") === (span.className ?? "")) {
      prev.end = span.end;
      prev.text += span.text;
      return;
    }
    out.push(span);
  };
  for (const span of spans) {
    if (span.end <= start || span.start >= end) {
      pushSpan(span);
      continue;
    }

    if (span.start < start) {
      pushSpan({ ...span, end: start, text: formula.slice(span.start, start) });
    }

    const overlapStart = Math.max(span.start, start);
    const overlapEnd = Math.min(span.end, end);
    const errorClassName = span.className ? `${span.className} formula-bar-token--error` : "formula-bar-token--error";
    pushSpan({
      ...span,
      start: overlapStart,
      end: overlapEnd,
      text: formula.slice(overlapStart, overlapEnd),
      className: errorClassName,
    });

    if (span.end > end) {
      pushSpan({ ...span, start: end, text: formula.slice(end, span.end) });
    }
  }

  return out;
}
