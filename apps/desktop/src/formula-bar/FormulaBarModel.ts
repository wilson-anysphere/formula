import { explainFormulaError, type ErrorExplanation } from "./errors.js";
import { getFunctionCallContext, type FunctionHint } from "./highlight/functionContext.js";
import { getFunctionSignature, signatureParts } from "./highlight/functionSignatures.js";
import { highlightFormula, type HighlightSpan } from "./highlight/highlightFormula.js";
import { tokenizeFormula } from "./highlight/tokenizeFormula.js";
import { rangeToA1, type RangeAddress } from "../spreadsheet/a1.js";
import { parseSheetQualifiedA1Range } from "./parseSheetQualifiedA1Range.js";
import { formatSheetNameForA1 } from "../sheet/formatSheetNameForA1.js";
import {
  assignFormulaReferenceColors,
  extractFormulaReferences,
  type ColoredFormulaReference,
  type ExtractFormulaReferencesOptions,
  type FormulaReference,
  type FormulaReferenceRange,
} from "@formula/spreadsheet-frontend";
import type {
  FormulaParseError,
  FormulaPartialLexResult,
  FormulaPartialParseResult,
  FormulaSpan,
  FormulaToken as EngineFormulaToken,
  FunctionContext as EngineFunctionContext,
} from "@formula/engine";

type ActiveCellInfo = {
  address: string;
  input: string;
  value: unknown;
};

export type FormulaBarAiSuggestion = {
  text: string;
  preview?: unknown;
};

export class FormulaBarModel {
  #activeCell: ActiveCellInfo = { address: "A1", input: "", value: "" };
  #draft: string = "";
  #isEditing = false;
  #cursorStart = 0;
  #cursorEnd = 0;
  #extractFormulaReferencesOptions: ExtractFormulaReferencesOptions | null = null;
  #extractFormulaReferencesOptionsVersion = 0;
  #referenceExtractionCache:
    | {
        draft: string;
        optionsVersion: number;
        references: FormulaReference[];
      }
    | null = null;
  #highlightedSpansCache: HighlightSpan[] | null = null;
  #highlightedSpansCacheDraft: string | null = null;
  #rangeInsertion: { start: number; end: number } | null = null;
  #hoveredReference: RangeAddress | null = null;
  #hoveredReferenceText: string | null = null;
  #referenceColorByText = new Map<string, string>();
  #coloredReferences: ColoredFormulaReference[] = [];
  #activeReferenceIndex: number | null = null;
  #engineHighlightSpans: HighlightSpan[] | null = null;
  #engineLexTokens: EngineFormulaToken[] | null = null;
  #engineHighlightErrorSpanKey: string | null = null;
  #engineFunctionContext: EngineFunctionContext | null = null;
  #engineSyntaxError: FormulaParseError | null = null;
  #engineToolingFormula: string | null = null;
  #engineToolingLocaleId: string = "en-US";
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
  #aiSuggestionPreview: unknown | null = null;

  setActiveCell(info: ActiveCellInfo): void {
    this.#activeCell = { ...info };
    this.#draft = info.input ?? "";
    this.#isEditing = false;
    this.#cursorStart = this.#draft.length;
    this.#cursorEnd = this.#draft.length;
    this.#rangeInsertion = null;
    this.#hoveredReference = null;
    this.#hoveredReferenceText = null;
    this.#referenceColorByText.clear();
    this.#coloredReferences = [];
    this.#activeReferenceIndex = null;
    this.#clearEditorTooling();
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
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
    const trimmed = String(text ?? "").trim();
    if (!trimmed) return null;

    const a1 = parseSheetQualifiedA1Range(trimmed);
    if (a1) return a1;

    const { references } = extractFormulaReferences(trimmed, undefined, undefined, this.#extractFormulaReferencesOptions ?? undefined);
    const first = references[0];
    if (!first) return null;
    // Avoid accidentally claiming a range for a more complex expression by requiring the
    // extracted reference to cover the full input string.
    if (first.start !== 0 || first.end !== trimmed.length) return null;

    const r = first.range;
    return { start: { row: r.startRow, col: r.startCol }, end: { row: r.endRow, col: r.endCol } };
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

  get cursorStart(): number {
    return this.#cursorStart;
  }

  get cursorEnd(): number {
    return this.#cursorEnd;
  }

  beginEdit(): void {
    this.#isEditing = true;
    this.#clearEditorTooling();
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
    this.#cursorStart = nextCursorStart;
    this.#cursorEnd = nextCursorEnd;
    this.#rangeInsertion = null;
    if (draftChanged) {
      // Draft text changed; any cached engine tokens/spans are now stale.
      this.#clearEditorTooling();
    } else if (cursorChanged) {
      // Cursor moved within the same draft; keep lex-based highlights/syntax errors but
      // clear cursor-dependent parse context so hint rendering can refresh.
      this.#engineFunctionContext = null;
    }
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
    this.#updateReferenceHighlights();
    this.#updateHoverFromCursor();
  }

  commit(): string {
    this.#isEditing = false;
    this.#activeCell = { ...this.#activeCell, input: this.#draft };
    this.#rangeInsertion = null;
    this.#clearEditorTooling();
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
    this.#hoveredReference = null;
    this.#hoveredReferenceText = null;
    this.#referenceColorByText.clear();
    this.#coloredReferences = [];
    this.#activeReferenceIndex = null;
    return this.#draft;
  }

  cancel(): void {
    this.#isEditing = false;
    this.#draft = this.#activeCell.input;
    this.#cursorStart = this.#draft.length;
    this.#cursorEnd = this.#draft.length;
    this.#rangeInsertion = null;
    this.#clearEditorTooling();
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
    this.#hoveredReference = null;
    this.#hoveredReferenceText = null;
    this.#referenceColorByText.clear();
    this.#coloredReferences = [];
    this.#activeReferenceIndex = null;
  }

  highlightedSpans(): HighlightSpan[] {
    if (this.#engineToolingFormula === this.#draft && this.#engineHighlightSpans) {
      return this.#engineHighlightSpans;
    }
    if (this.#highlightedSpansCache && this.#highlightedSpansCacheDraft === this.#draft) {
      return this.#highlightedSpansCache;
    }
    const spans = highlightFormula(this.#draft);
    this.#highlightedSpansCache = spans;
    this.#highlightedSpansCacheDraft = this.#draft;
    return spans;
  }

  functionHint(): FunctionHint | null {
    const hasEngineLocaleForDraft = this.#engineToolingFormula === this.#draft;
    const localeId = hasEngineLocaleForDraft
      ? this.#engineToolingLocaleId
      : (typeof document !== "undefined" ? document.documentElement?.lang : "")?.trim?.() || "en-US";
    const argSeparator = inferArgSeparator(localeId);

    if (hasEngineLocaleForDraft && this.#engineFunctionContext) {
      const ctx = this.#engineFunctionContext;
      const signature = getFunctionSignature(ctx.name);
      if (!signature) return null;

      return {
        context: { name: ctx.name, argIndex: ctx.argIndex },
        signature,
        parts: signatureParts(signature, ctx.argIndex, { argSeparator }),
      };
    }

    let context = getFunctionCallContext(this.#draft, this.#cursorStart);

    // Excel UX: keep showing the innermost function hint when the caret is just
    // after a closing paren (e.g. "=ROUND(1,2)|"). The simple tokenizer-based
    // parser considers the call "closed" once it consumes ')', which would
    // otherwise clear the hint.
    if (
      !context &&
      this.#cursorStart === this.#cursorEnd &&
      this.#cursorStart > 0 &&
      this.#draft[this.#cursorStart - 1] === ")"
    ) {
      context = getFunctionCallContext(this.#draft, this.#cursorStart - 1);
    }

    if (!context) return null;
    const signature = getFunctionSignature(context.name);
    if (!signature) return null;
    return {
      context,
      signature,
      parts: signatureParts(signature, context.argIndex, { argSeparator }),
    };
  }

  errorExplanation(): ErrorExplanation | null {
    return explainFormulaError(this.#activeCell.value);
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
    const errorSpanKey =
      errorSpan && errorSpan.end > errorSpan.start ? `${errorSpan.start}:${errorSpan.end}` : null;

    const highlightStable =
      this.#engineHighlightSpans != null &&
      this.#engineToolingFormula === args.formula &&
      this.#engineLexTokens === args.lexResult.tokens &&
      this.#engineHighlightErrorSpanKey === errorSpanKey;

    if (!highlightStable) {
      const referenceTokens = tokenizeFormula(args.formula).filter((t) => t.type === "reference");
      const engineSpans = highlightFromEngineTokens(args.formula, args.lexResult.tokens);
      const withRefs = spliceReferenceSpans(args.formula, engineSpans, referenceTokens);
      const withError = applyErrorSpan(args.formula, withRefs, errorSpan && errorSpan.end > errorSpan.start ? errorSpan : null);
      this.#engineHighlightSpans = withError;
      this.#engineLexTokens = args.lexResult.tokens;
      this.#engineHighlightErrorSpanKey = errorSpanKey;
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
    return this.#coloredReferences.map((ref) => ({
      range: ref.range,
      color: ref.color,
      text: ref.text,
      index: ref.index,
      active: this.#activeReferenceIndex === ref.index,
    }));
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
    this.#aiSuggestionPreview = null;
    this.#updateReferenceHighlights();
    this.#updateHoverFromCursor();
  }

  updateRangeSelection(range: RangeAddress, sheetId?: string): void {
    if (!this.#isEditing) return;
    this.#insertOrReplaceRange(formatRangeReference(range, sheetId), false);
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
    this.#updateReferenceHighlights();
    this.#updateHoverFromCursor();
  }

  endRangeSelection(): void {
    this.#rangeInsertion = null;
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
  }

  setAiSuggestion(suggestion: string | FormulaBarAiSuggestion | null): void {
    const suggestionText = typeof suggestion === "string" ? suggestion : suggestion?.text;
    const preview = typeof suggestion === "string" ? null : (suggestion?.preview ?? null);

    if (!suggestionText) {
      this.#aiSuggestion = null;
      this.#aiSuggestionPreview = null;
      return;
    }

    // Normalize "tail" suggestions ("M") into full-string suggestions ("=SUM")
    // so `aiGhostText()` + `acceptAiSuggestion()` can operate consistently.
    if (!this.#isEditing) {
      this.#aiSuggestion = suggestionText;
      this.#aiSuggestionPreview = preview;
      return;
    }

    const start = Math.min(this.#cursorStart, this.#cursorEnd);
    const end = Math.max(this.#cursorStart, this.#cursorEnd);
    const prefix = this.#draft.slice(0, start);
    const suffix = this.#draft.slice(end);
    const looksLikeFullSuggestion =
      suggestionText.startsWith(prefix) && (!suffix || suggestionText.endsWith(suffix));
    this.#aiSuggestion = looksLikeFullSuggestion ? suggestionText : prefix + suggestionText + suffix;
    this.#aiSuggestionPreview = preview;
  }

  aiSuggestion(): string | null {
    return this.#aiSuggestion;
  }

  aiSuggestionPreview(): unknown | null {
    return this.#aiSuggestionPreview;
  }

  aiGhostText(): string {
    if (!this.#aiSuggestion) return "";
    if (this.#cursorStart !== this.#cursorEnd) return "";

    const cursor = this.#cursorStart;
    const prefix = this.#draft.slice(0, cursor);
    const suffix = this.#draft.slice(cursor);
    if (!this.#aiSuggestion.startsWith(prefix)) return "";
    if (suffix && !this.#aiSuggestion.endsWith(suffix)) return "";
    return this.#aiSuggestion.slice(cursor, this.#aiSuggestion.length - suffix.length);
  }

  acceptAiSuggestion(): boolean {
    if (!this.#aiSuggestion) return false;
    if (!this.#isEditing) return false;

    const suggestionText = this.#aiSuggestion;
    const start = Math.min(this.#cursorStart, this.#cursorEnd);
    const end = Math.max(this.#cursorStart, this.#cursorEnd);

    const prefix = this.#draft.slice(0, start);
    const suffix = this.#draft.slice(end);

    // The completion engine typically supplies the full suggested text, but some
    // surfaces may pass only the "ghost" tail (text to insert at the caret).
    const looksLikeFullReplacement =
      suggestionText.startsWith(prefix) && (suffix.length === 0 || suggestionText.endsWith(suffix));

    if (looksLikeFullReplacement) {
      const ghost = this.aiGhostText();
      this.#draft = suggestionText;
      const newCursor = ghost ? start + ghost.length : suggestionText.length - suffix.length;
      this.#cursorStart = newCursor;
      this.#cursorEnd = newCursor;
    } else {
      this.#draft = this.#draft.slice(0, start) + suggestionText + this.#draft.slice(end);
      const newCursor = start + suggestionText.length;
      this.#cursorStart = newCursor;
      this.#cursorEnd = newCursor;
    }
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
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
      this.#rangeInsertion = { start, end: start + rangeText.length };
      this.#cursorStart = this.#rangeInsertion.end;
      this.#cursorEnd = this.#rangeInsertion.end;
      return;
    }

    const { start, end } = this.#rangeInsertion;
    this.#draft = this.#draft.slice(0, start) + rangeText + this.#draft.slice(end);
    this.#rangeInsertion = { start, end: start + rangeText.length };
    this.#cursorStart = this.#rangeInsertion.end;
    this.#cursorEnd = this.#rangeInsertion.end;
  }

  #updateHoverFromCursor(): void {
    if (!this.#isEditing || !this.#draft.trim().startsWith("=")) {
      this.#hoveredReference = null;
      this.#hoveredReferenceText = null;
      return;
    }

    const activeIndex = this.#activeReferenceIndex;
    if (activeIndex == null) {
      this.#hoveredReference = null;
      this.#hoveredReferenceText = null;
      return;
    }

    // Reuse the already-computed reference metadata to avoid re-tokenizing the full
    // formula string on every cursor move.
    const active = this.#coloredReferences[activeIndex] ?? null;
    if (!active) {
      this.#hoveredReference = null;
      this.#hoveredReferenceText = null;
      return;
    }

    this.#hoveredReferenceText = active.text;
    this.#hoveredReference = {
      start: { row: active.range.startRow, col: active.range.startCol },
      end: { row: active.range.endRow, col: active.range.endCol },
    };

    // NOTE: `active.text` can be a sheet-qualified A1 reference, a named range, or
    // (when configured) a structured reference. Consumers that need sheet gating
    // should use `hoveredReferenceText()` (and resolve names/sheet qualifiers there).
  }

  #updateReferenceHighlights(): void {
    if (!this.#isEditing || !this.#draft.trim().startsWith("=")) {
      this.#coloredReferences = [];
      this.#activeReferenceIndex = null;
      return;
    }

    const cache = this.#referenceExtractionCache;
    let references: FormulaReference[];
    const cacheHit =
      cache != null && cache.draft === this.#draft && cache.optionsVersion === this.#extractFormulaReferencesOptionsVersion;
    if (cacheHit) {
      references = cache.references;
    } else {
      references = extractFormulaReferences(
        this.#draft,
        undefined,
        undefined,
        this.#extractFormulaReferencesOptions ?? undefined
      ).references;
      this.#referenceExtractionCache = {
        draft: this.#draft,
        optionsVersion: this.#extractFormulaReferencesOptionsVersion,
        references,
      };
    }

    const activeIndex = findActiveReferenceIndex(references, this.#cursorStart, this.#cursorEnd);

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
    this.#engineHighlightErrorSpanKey = null;
    this.#engineFunctionContext = null;
    this.#engineSyntaxError = null;
    this.#engineToolingFormula = null;
  }
}

function findActiveReferenceIndex(
  references: readonly Pick<FormulaReference, "start" | "end" | "index">[],
  cursorStart: number,
  cursorEnd: number
): number | null {
  const start = Math.min(cursorStart, cursorEnd);
  const end = Math.max(cursorStart, cursorEnd);

  // If text is selected, treat a reference as active only when the selection is contained
  // within that reference token.
  if (start !== end) {
    const active = references.find((ref) => start >= ref.start && end <= ref.end);
    return active ? active.index : null;
  }

  // Caret: treat the reference containing either the character at the caret or
  // immediately before it as active. This matches typical editor behavior where
  // being at the end of a token still counts as "in" the token.
  const positions = start === 0 ? [0] : [start, start - 1];
  for (const pos of positions) {
    const active = references.find((ref) => ref.start <= pos && pos < ref.end);
    if (active) return active.index;
  }
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

const ARG_SEPARATOR_CACHE = new Map<string, string>();

function inferArgSeparator(localeId: string): string {
  const locale = localeId?.trim?.() || "en-US";
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

function highlightFromEngineTokens(formula: string, tokens: EngineFormulaToken[]): HighlightSpan[] {
  const filtered = tokens.filter((t) => t.kind !== "Eof");
  filtered.sort((a, b) => a.span.start - b.span.start);

  // Best-effort: treat an identifier immediately followed by "(" as a function name.
  const functionIdentStarts = new Set<number>();
  for (let i = 0; i < filtered.length; i += 1) {
    const token = filtered[i]!;
    if (token.kind !== "Ident" && token.kind !== "QuotedIdent") continue;
    for (let j = i + 1; j < filtered.length; j += 1) {
      const next = filtered[j]!;
      if (next.kind === "Whitespace") continue;
      if (next.kind === "LParen") functionIdentStarts.add(token.span.start);
      break;
    }
  }

  const spans: HighlightSpan[] = [];
  let pos = 0;
  for (const token of filtered) {
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
      }
    }

    if (end > start) {
      spans.push({
        kind: engineTokenKindToHighlightKind(token, functionIdentStarts.has(token.span.start)),
        start,
        end,
        text: formula.slice(start, end),
      });
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

function engineTokenKindToHighlightKind(token: EngineFormulaToken, isFunctionIdent: boolean): HighlightSpan["kind"] {
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
      return isFunctionIdent ? "function" : "identifier";
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

  const refs = [...referenceTokens].sort((a, b) => a.start - b.start);
  const out: HighlightSpan[] = [];
  let pos = 0;

  for (const ref of refs) {
    const start = Math.max(0, Math.min(ref.start, formula.length));
    const end = Math.max(0, Math.min(ref.end, formula.length));
    if (end <= start) continue;
    if (start > pos) {
      out.push(...sliceSpans(formula, spans, pos, start));
    }
    out.push({ kind: "reference", start, end, text: formula.slice(start, end) });
    pos = end;
  }

  if (pos < formula.length) {
    out.push(...sliceSpans(formula, spans, pos, formula.length));
  }

  return mergeAdjacent(out);
}

function sliceSpans(formula: string, spans: HighlightSpan[], start: number, end: number): HighlightSpan[] {
  const out: HighlightSpan[] = [];
  for (const span of spans) {
    if (span.end <= start) continue;
    if (span.start >= end) break;
    const s = Math.max(span.start, start);
    const e = Math.min(span.end, end);
    if (e <= s) continue;
    out.push({ ...span, start: s, end: e, text: formula.slice(s, e) });
  }
  return out;
}

function mergeAdjacent(spans: HighlightSpan[]): HighlightSpan[] {
  if (spans.length === 0) return spans;
  const out: HighlightSpan[] = [spans[0]!];
  for (let i = 1; i < spans.length; i += 1) {
    const prev = out[out.length - 1]!;
    const next = spans[i]!;
    if (prev.end === next.start && prev.kind === next.kind && (prev.className ?? "") === (next.className ?? "")) {
      prev.end = next.end;
      prev.text += next.text;
      continue;
    }
    out.push(next);
  }
  return out;
}

function applyErrorSpan(formula: string, spans: HighlightSpan[], errorSpan: FormulaSpan | null): HighlightSpan[] {
  if (!errorSpan) return spans;
  const start = Math.max(0, Math.min(errorSpan.start, formula.length));
  const end = Math.max(0, Math.min(errorSpan.end, formula.length));
  if (end <= start) return spans;

  const out: HighlightSpan[] = [];
  for (const span of spans) {
    if (span.end <= start || span.start >= end) {
      out.push(span);
      continue;
    }

    if (span.start < start) {
      out.push({ ...span, end: start, text: formula.slice(span.start, start) });
    }

    const overlapStart = Math.max(span.start, start);
    const overlapEnd = Math.min(span.end, end);
    out.push({
      ...span,
      start: overlapStart,
      end: overlapEnd,
      text: formula.slice(overlapStart, overlapEnd),
      className: [span.className, "formula-bar-token--error"].filter(Boolean).join(" "),
    });

    if (span.end > end) {
      out.push({ ...span, start: end, text: formula.slice(end, span.end) });
    }
  }

  return mergeAdjacent(out);
}
