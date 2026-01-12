import { explainFormulaError, type ErrorExplanation } from "./errors.js";
import { getFunctionHint, type FunctionHint } from "./highlight/functionContext.js";
import { highlightFormula, type HighlightSpan } from "./highlight/highlightFormula.js";
import { tokenizeFormula } from "./highlight/tokenizeFormula.js";
import { rangeToA1, type RangeAddress } from "../spreadsheet/a1.js";
import { parseSheetQualifiedA1Range } from "./parseSheetQualifiedA1Range.js";
import { formatSheetNameForA1 } from "../sheet/formatSheetNameForA1.js";
import {
  assignFormulaReferenceColors,
  extractFormulaReferences,
  type ColoredFormulaReference,
  type FormulaReferenceRange,
} from "@formula/spreadsheet-frontend";

export type ActiveCellInfo = {
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
  #rangeInsertion: { start: number; end: number } | null = null;
  #hoveredReference: RangeAddress | null = null;
  #referenceColorByText = new Map<string, string>();
  #coloredReferences: ColoredFormulaReference[] = [];
  #activeReferenceIndex: number | null = null;
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
    this.#referenceColorByText.clear();
    this.#coloredReferences = [];
    this.#activeReferenceIndex = null;
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
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
    this.#updateReferenceHighlights();
    this.#updateHoverFromCursor();
  }

  updateDraft(draft: string, cursorStart: number, cursorEnd: number): void {
    this.#isEditing = true;
    this.#draft = draft;
    this.#cursorStart = Math.max(0, Math.min(cursorStart, draft.length));
    this.#cursorEnd = Math.max(0, Math.min(cursorEnd, draft.length));
    this.#rangeInsertion = null;
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
    this.#updateReferenceHighlights();
    this.#updateHoverFromCursor();
  }

  commit(): string {
    this.#isEditing = false;
    this.#activeCell = { ...this.#activeCell, input: this.#draft };
    this.#rangeInsertion = null;
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
    this.#hoveredReference = null;
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
    this.#aiSuggestion = null;
    this.#aiSuggestionPreview = null;
    this.#hoveredReference = null;
    this.#referenceColorByText.clear();
    this.#coloredReferences = [];
    this.#activeReferenceIndex = null;
  }

  highlightedSpans(): HighlightSpan[] {
    return highlightFormula(this.#draft);
  }

  functionHint(): FunctionHint | null {
    return getFunctionHint(this.#draft, this.#cursorStart);
  }

  errorExplanation(): ErrorExplanation | null {
    return explainFormulaError(this.#activeCell.value);
  }

  setHoveredReference(referenceText: string | null): void {
    if (!referenceText) {
      this.#hoveredReference = null;
      return;
    }
    this.#hoveredReference = parseSheetQualifiedA1Range(referenceText);
  }

  hoveredReference(): RangeAddress | null {
    return this.#hoveredReference;
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
    if (this.#cursorStart !== this.#cursorEnd) {
      this.#hoveredReference = null;
      return;
    }

    const cursor = this.#cursorStart;
    const probe = cursor > 0 ? cursor - 1 : 0;
    const token = tokenizeFormula(this.#draft).find((t) => t.start <= probe && probe < t.end);
    if (!token || token.type !== "reference") {
      this.#hoveredReference = null;
      return;
    }

    this.#hoveredReference = parseSheetQualifiedA1Range(token.text);
  }

  #updateReferenceHighlights(): void {
    if (!this.#isEditing || !this.#draft.trim().startsWith("=")) {
      this.#coloredReferences = [];
      this.#activeReferenceIndex = null;
      return;
    }

    const { references, activeIndex } = extractFormulaReferences(this.#draft, this.#cursorStart, this.#cursorEnd);
    const { colored, nextByText } = assignFormulaReferenceColors(references, this.#referenceColorByText);
    this.#referenceColorByText = nextByText;
    this.#coloredReferences = colored;
    this.#activeReferenceIndex = activeIndex;
  }
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
