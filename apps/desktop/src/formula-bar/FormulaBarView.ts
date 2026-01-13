import { FormulaBarModel, type FormulaBarAiSuggestion } from "./FormulaBarModel.js";
import { type RangeAddress } from "../spreadsheet/a1.js";
import { parseSheetQualifiedA1Range } from "./parseSheetQualifiedA1Range.js";
import { toggleA1AbsoluteAtCursor, type FormulaReferenceRange } from "@formula/spreadsheet-frontend";
import { searchFunctionResults } from "../command-palette/commandPaletteSearch.js";
import FUNCTION_CATALOG from "../../../../shared/functionCatalog.mjs";
import { getFunctionSignature, type FunctionSignature } from "./highlight/functionSignatures.js";

type FunctionPickerItem = {
  name: string;
  signature?: string;
  summary?: string;
};

type CatalogFunction = { name?: string | null };

const ALL_FUNCTION_NAMES_SORTED: string[] = ((FUNCTION_CATALOG as { functions?: CatalogFunction[] } | null)?.functions ?? [])
  .map((fn) => String(fn?.name ?? "").trim())
  .filter((name) => name.length > 0)
  .sort((a, b) => a.localeCompare(b));

const COMMON_FUNCTION_NAMES = [
  "SUM",
  "AVERAGE",
  "COUNT",
  "MIN",
  "MAX",
  "IF",
  "IFERROR",
  "XLOOKUP",
  "VLOOKUP",
  "INDEX",
  "MATCH",
  "TODAY",
  "NOW",
  "ROUND",
  "CONCAT",
  "SEQUENCE",
  "FILTER",
  "TEXTSPLIT",
];

const DEFAULT_FUNCTION_NAMES: string[] = (() => {
  const byName = new Map(ALL_FUNCTION_NAMES_SORTED.map((name) => [name.toUpperCase(), name]));
  const out: string[] = [];
  const seen = new Set<string>();

  for (const name of COMMON_FUNCTION_NAMES) {
    const resolved = byName.get(name.toUpperCase());
    if (!resolved) continue;
    out.push(resolved);
    seen.add(resolved.toUpperCase());
  }
  for (const name of ALL_FUNCTION_NAMES_SORTED) {
    if (out.length >= 50) break;
    const key = name.toUpperCase();
    if (seen.has(key)) continue;
    out.push(name);
    seen.add(key);
  }
  return out;
})();

export interface FormulaBarViewCallbacks {
  onBeginEdit?: (activeCellAddress: string) => void;
  onCommit: (text: string, commit: FormulaBarCommit) => void;
  onCancel?: () => void;
  onGoTo?: (reference: string) => void;
  onOpenNameBoxMenu?: () => void | Promise<void>;
  onHoverRange?: (range: RangeAddress | null) => void;
  onReferenceHighlights?: (
    highlights: Array<{ range: FormulaReferenceRange; color: string; text: string; index: number; active?: boolean }>
  ) => void;
}

export type FormulaBarCommitReason = "enter" | "tab" | "command";

export interface FormulaBarCommit {
  reason: FormulaBarCommitReason;
  /**
   * Shift modifier for enter/tab (Shift+Enter moves up, Shift+Tab moves left).
   */
  shift: boolean;
}

export class FormulaBarView {
  readonly model = new FormulaBarModel();

  readonly root: HTMLElement;
  readonly textarea: HTMLTextAreaElement;

  #scheduledRender:
    | { id: number; kind: "raf" }
    | { id: ReturnType<typeof setTimeout>; kind: "timeout" }
    | null = null;
  #pendingRender: { preserveTextareaValue: boolean } | null = null;
  #lastHighlightHtml: string | null = null;

  #nameBoxDropdownEl: HTMLButtonElement;
  #cancelButtonEl: HTMLButtonElement;
  #commitButtonEl: HTMLButtonElement;
  #fxButtonEl: HTMLButtonElement;
  #addressEl: HTMLInputElement;
  #highlightEl: HTMLElement;
  #hintEl: HTMLElement;
  #errorButton: HTMLButtonElement;
  #errorPanel: HTMLElement;
  #isErrorPanelOpen = false;
  #hoverOverride: RangeAddress | null = null;
  #selectedReferenceIndex: number | null = null;
  #callbacks: FormulaBarViewCallbacks;

  #functionPickerEl: HTMLDivElement;
  #functionPickerInputEl: HTMLInputElement;
  #functionPickerListEl: HTMLUListElement;
  #functionPickerOpen = false;
  #functionPickerItems: FunctionPickerItem[] = [];
  #functionPickerItemEls: HTMLLIElement[] = [];
  #functionPickerSelectedIndex = 0;
  #functionPickerAnchorSelection: { start: number; end: number } | null = null;
  #functionPickerDocMouseDown = (e: MouseEvent) => this.#onFunctionPickerDocMouseDown(e);

  constructor(root: HTMLElement, callbacks: FormulaBarViewCallbacks) {
    this.root = root;
    this.#callbacks = callbacks;

    root.classList.add("formula-bar");

    const row = document.createElement("div");
    row.className = "formula-bar-row";

    const address = document.createElement("input");
    address.className = "formula-bar-address";
    address.dataset.testid = "formula-address";
    address.setAttribute("aria-label", "Name box");
    address.autocomplete = "off";
    address.spellcheck = false;
    address.value = "A1";

    const nameBox = document.createElement("div");
    nameBox.className = "formula-bar-name-box";

    const nameBoxDropdown = document.createElement("button");
    nameBoxDropdown.className = "formula-bar-name-box-dropdown";
    nameBoxDropdown.dataset.testid = "name-box-dropdown";
    nameBoxDropdown.type = "button";
    nameBoxDropdown.textContent = "▾";
    nameBoxDropdown.title = "Name box menu";
    nameBoxDropdown.setAttribute("aria-label", "Open name box menu");

    nameBox.appendChild(address);
    nameBox.appendChild(nameBoxDropdown);

    const actions = document.createElement("div");
    actions.className = "formula-bar-actions";

    const cancelButton = document.createElement("button");
    cancelButton.className = "formula-bar-action-button formula-bar-action-button--cancel";
    cancelButton.type = "button";
    cancelButton.textContent = "✕";
    cancelButton.title = "Cancel (Esc)";
    cancelButton.setAttribute("aria-label", "Cancel edit");

    const commitButton = document.createElement("button");
    commitButton.className = "formula-bar-action-button formula-bar-action-button--commit";
    commitButton.type = "button";
    commitButton.textContent = "✓";
    commitButton.title = "Enter (↵)";
    commitButton.setAttribute("aria-label", "Commit edit");

    const fxButton = document.createElement("button");
    fxButton.className = "formula-bar-action-button formula-bar-action-button--fx";
    fxButton.type = "button";
    fxButton.textContent = "fx";
    fxButton.title = "Insert function";
    fxButton.setAttribute("aria-label", "Insert function");
    fxButton.dataset.testid = "formula-fx-button";

    actions.appendChild(cancelButton);
    actions.appendChild(commitButton);
    actions.appendChild(fxButton);

    const editor = document.createElement("div");
    editor.className = "formula-bar-editor";

    const highlight = document.createElement("pre");
    highlight.className = "formula-bar-highlight";
    highlight.dataset.testid = "formula-highlight";
    highlight.setAttribute("aria-hidden", "true");

    const textarea = document.createElement("textarea");
    textarea.className = "formula-bar-input";
    textarea.dataset.testid = "formula-input";
    textarea.spellcheck = false;
    textarea.autocapitalize = "off";
    textarea.autocomplete = "off";
    textarea.wrap = "off";
    textarea.rows = 1;

    editor.appendChild(highlight);
    editor.appendChild(textarea);

    const errorButton = document.createElement("button");
    errorButton.className = "formula-bar-error-button";
    errorButton.type = "button";
    errorButton.textContent = "!";
    errorButton.title = "Show error details";
    errorButton.setAttribute("aria-label", "Show formula error");
    errorButton.setAttribute("aria-expanded", "false");
    errorButton.dataset.testid = "formula-error-button";

    const errorPanel = document.createElement("div");
    errorPanel.className = "formula-bar-error-panel";
    errorPanel.dataset.testid = "formula-error-panel";

    row.appendChild(nameBox);
    row.appendChild(actions);
    row.appendChild(editor);
    row.appendChild(errorButton);

    const hint = document.createElement("div");
    hint.className = "formula-bar-hint";
    hint.dataset.testid = "formula-hint";

    const functionPicker = document.createElement("div");
    functionPicker.className = "formula-function-picker";
    functionPicker.dataset.testid = "formula-function-picker";
    functionPicker.hidden = true;
    functionPicker.setAttribute("role", "dialog");
    functionPicker.setAttribute("aria-label", "Insert function");

    const functionPickerPanel = document.createElement("div");
    functionPickerPanel.className = "command-palette";
    functionPickerPanel.dataset.testid = "formula-function-picker-panel";

    const functionPickerInput = document.createElement("input");
    functionPickerInput.className = "command-palette__input";
    functionPickerInput.dataset.testid = "formula-function-picker-input";
    functionPickerInput.placeholder = "Search functions";
    functionPickerInput.setAttribute("aria-label", "Search functions");

    const functionPickerList = document.createElement("ul");
    functionPickerList.className = "command-palette__list";
    functionPickerList.dataset.testid = "formula-function-picker-list";
    functionPickerList.setAttribute("role", "listbox");
    // Ensure there is at least one tabbable element besides the input so Tab doesn't escape.
    functionPickerList.tabIndex = 0;

    functionPickerPanel.appendChild(functionPickerInput);
    functionPickerPanel.appendChild(functionPickerList);
    functionPicker.appendChild(functionPickerPanel);

    root.appendChild(row);
    root.appendChild(hint);
    root.appendChild(errorPanel);
    root.appendChild(functionPicker);

    this.textarea = textarea;
    this.#nameBoxDropdownEl = nameBoxDropdown;
    this.#cancelButtonEl = cancelButton;
    this.#commitButtonEl = commitButton;
    this.#fxButtonEl = fxButton;
    this.#addressEl = address;
    this.#highlightEl = highlight;
    this.#hintEl = hint;
    this.#errorButton = errorButton;
    this.#errorPanel = errorPanel;
    this.#functionPickerEl = functionPicker;
    this.#functionPickerInputEl = functionPickerInput;
    this.#functionPickerListEl = functionPickerList;

    address.addEventListener("focus", () => {
      address.select();
    });

    nameBoxDropdown.addEventListener("click", () => {
      if (this.#callbacks.onOpenNameBoxMenu) {
        void this.#callbacks.onOpenNameBoxMenu();
        return;
      }

      // Fallback affordance: focus the address input so keyboard "Go To" still feels natural.
      address.focus();
    });

    address.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        const ref = address.value;
        // Blur before navigating so subsequent renders can update the value.
        address.blur();
        this.#callbacks.onGoTo?.(ref);
        return;
      }

      if (e.key === "Escape") {
        e.preventDefault();
        address.value = this.model.activeCell.address;
        address.blur();
      }
    });

    textarea.addEventListener("focus", () => this.#beginEditFromFocus());
    textarea.addEventListener("input", () => this.#onInputOrSelection());
    textarea.addEventListener("click", () => this.#onTextareaClick());
    textarea.addEventListener("keyup", () => this.#onInputOrSelection());
    textarea.addEventListener("select", () => this.#onInputOrSelection());
    textarea.addEventListener("scroll", () => this.#syncScroll());
    textarea.addEventListener("keydown", (e) => this.#onKeyDown(e));

    // When not editing, allow hover previews using the highlighted spans.
    highlight.addEventListener("mousemove", (e) => this.#onHighlightHover(e));
    highlight.addEventListener("mouseleave", () => this.#clearHoverOverride());
    highlight.addEventListener("mousedown", (e) => {
      // Prevent selecting text in <pre> and instead focus the textarea.
      e.preventDefault();
      this.focus({ cursor: "end" });
    });

    errorButton.addEventListener("click", () => {
      if (!this.root.classList.contains("formula-bar--has-error")) return;
      this.#setErrorPanelOpen(!this.#isErrorPanelOpen);
    });

    cancelButton.addEventListener("click", () => this.#cancel());
    commitButton.addEventListener("click", () => this.#commit({ reason: "command", shift: false }));
    fxButton.addEventListener("click", () => this.#focusFx());
    fxButton.addEventListener("mousedown", (e) => {
      // Preserve the caret/selection in the textarea when clicking the fx button.
      e.preventDefault();
    });

    functionPickerInput.addEventListener("input", () => this.#onFunctionPickerInput());
    const pickerKeyDown = (e: KeyboardEvent) => this.#onFunctionPickerKeyDown(e);
    functionPickerInput.addEventListener("keydown", pickerKeyDown);
    functionPickerList.addEventListener("keydown", pickerKeyDown);

    // Initial render.
    this.model.setActiveCell({ address: "A1", input: "", value: "" });
    this.#render({ preserveTextareaValue: false });
  }

  setAiSuggestion(suggestion: string | FormulaBarAiSuggestion | null): void {
    this.model.setAiSuggestion(suggestion);
    this.#render({ preserveTextareaValue: true });
  }

  focus(opts: { cursor?: "end" | "all" } = {}): void {
    // Ensure the textarea is visible so `.focus()` works even when the formula bar is not currently editing.
    // `#render()` keeps this class in sync with `model.isEditing`, but `focus()` is called while
    // still in view mode, so we need to allow the textarea to become focusable first.
    this.root.classList.add("formula-bar--editing");
    // Prevent browser focus handling from scrolling the desktop shell horizontally.
    // (The app uses its own scroll containers; window scrolling is accidental and
    // breaks pointer-coordinate based interactions like range-drag insertion.)
    try {
      this.textarea.focus({ preventScroll: true });
    } catch {
      // Older browsers may not support FocusOptions; fall back to default focus behavior.
      this.textarea.focus();
    }
    if (opts.cursor === "all") {
      this.textarea.setSelectionRange(0, this.textarea.value.length);
    } else if (opts.cursor === "end") {
      const end = this.textarea.value.length;
      this.textarea.setSelectionRange(end, end);
    }
    this.#onInputOrSelection();
  }

  setActiveCell(info: { address: string; input: string; value: unknown }): void {
    if (this.model.isEditing) return;
    this.model.setActiveCell(info);
    this.#hoverOverride = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
  }

  isEditing(): boolean {
    return this.model.isEditing;
  }

  commitEdit(reason: FormulaBarCommitReason = "command", shift = false): void {
    this.#commit({ reason, shift });
  }

  cancelEdit(): void {
    this.#cancel();
  }

  isFormulaEditing(): boolean {
    return this.model.isEditing && this.model.draft.trim().startsWith("=");
  }

  beginRangeSelection(range: RangeAddress, sheetId?: string): void {
    this.model.beginEdit();
    this.model.beginRangeSelection(range, sheetId);
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    this.#setTextareaSelectionFromModel();
    this.#emitOverlays();
  }

  updateRangeSelection(range: RangeAddress, sheetId?: string): void {
    this.model.updateRangeSelection(range, sheetId);
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    this.#setTextareaSelectionFromModel();
    this.#emitOverlays();
  }

  endRangeSelection(): void {
    this.model.endRangeSelection();
  }

  #beginEditFromFocus(): void {
    if (this.model.isEditing) return;
    this.model.beginEdit();
    this.#callbacks.onBeginEdit?.(this.model.activeCell.address);
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: true });
    this.#emitOverlays();
  }

  #onInputOrSelection(): void {
    if (!this.model.isEditing) return;

    const value = this.textarea.value;
    const start = this.textarea.selectionStart ?? value.length;
    const end = this.textarea.selectionEnd ?? value.length;

    // "keyup" and "select" events can fire without changing the textarea value/selection.
    // Skip redundant model updates/renders in that case.
    if (this.model.draft === value && this.model.cursorStart === start && this.model.cursorEnd === end) {
      return;
    }

    this.model.updateDraft(value, start, end);
    this.#selectedReferenceIndex = this.#inferSelectedReferenceIndex(start, end);
    this.#requestRender({ preserveTextareaValue: true });
    this.#emitOverlays();
  }

  #onTextareaClick(): void {
    if (!this.model.isEditing) return;

    const prevSelectedReferenceIndex = this.#selectedReferenceIndex;
    const value = this.textarea.value;
    const start = this.textarea.selectionStart ?? value.length;
    const end = this.textarea.selectionEnd ?? value.length;
    this.model.updateDraft(value, start, end);
    this.#selectedReferenceIndex = this.#inferSelectedReferenceIndex(start, end);

    const isFormulaEditing = this.model.draft.trim().startsWith("=");
    if (isFormulaEditing && start === end) {
      const activeIndex = this.model.activeReferenceIndex();
      const active = activeIndex == null ? null : this.model.coloredReferences()[activeIndex] ?? null;

      if (active) {
        // Excel UX: clicking a reference selects the entire reference token so
        // range dragging replaces it. A subsequent click on the same reference
        // toggles back to a caret, allowing manual edits within the reference.
        if (prevSelectedReferenceIndex === activeIndex) {
          this.#selectedReferenceIndex = null;
        } else {
          this.textarea.setSelectionRange(active.start, active.end);
          this.model.updateDraft(this.textarea.value, active.start, active.end);
          this.#selectedReferenceIndex = activeIndex;
        }
      }
    }

    this.#requestRender({ preserveTextareaValue: true });
    this.#emitOverlays();
  }

  #requestRender(opts: { preserveTextareaValue: boolean }): void {
    // Merge pending render options; if any caller needs to overwrite the textarea
    // value, the combined render must also overwrite it.
    if (this.#pendingRender) {
      this.#pendingRender = {
        preserveTextareaValue: this.#pendingRender.preserveTextareaValue && opts.preserveTextareaValue,
      };
    } else {
      this.#pendingRender = opts;
    }

    // Coalesce multiple rapid input/keyup/select events into a single render per frame.
    if (this.#scheduledRender) return;

    const flush = (): void => {
      // Clear the scheduled handle before rendering so a render can schedule another frame.
      this.#scheduledRender = null;
      const pending = this.#pendingRender ?? { preserveTextareaValue: true };
      this.#pendingRender = null;
      this.#render(pending);
    };

    if (typeof requestAnimationFrame === "function") {
      const id = requestAnimationFrame(() => flush());
      this.#scheduledRender = { id, kind: "raf" };
    } else {
      const id = setTimeout(() => flush(), 0);
      this.#scheduledRender = { id, kind: "timeout" };
    }
  }

  #cancelPendingRender(): void {
    if (!this.#scheduledRender) return;
    const { id, kind } = this.#scheduledRender;
    if (kind === "raf") {
      // `cancelAnimationFrame` isn't present in all JS environments; fall back safely.
      try {
        cancelAnimationFrame(id);
      } catch {
        // ignore
      }
    } else {
      clearTimeout(id);
    }
    this.#scheduledRender = null;
    this.#pendingRender = null;
  }

  #onKeyDown(e: KeyboardEvent): void {
    if (!this.model.isEditing) return;

    if (e.key === "F4" && !e.altKey && !e.ctrlKey && !e.metaKey && this.model.draft.trim().startsWith("=")) {
      e.preventDefault();

      const prevText = this.textarea.value;
      const cursorStart = this.textarea.selectionStart ?? prevText.length;
      const cursorEnd = this.textarea.selectionEnd ?? prevText.length;

      // Ensure model-derived reference metadata is current for the F4 operation
      // (the selection may have changed without triggering our keyup/select listeners yet).
      if (
        this.model.draft !== prevText ||
        this.model.cursorStart !== cursorStart ||
        this.model.cursorEnd !== cursorEnd
      ) {
        this.model.updateDraft(prevText, cursorStart, cursorEnd);
      }

      const toggled = toggleA1AbsoluteAtCursor(prevText, cursorStart, cursorEnd);
      if (!toggled) return;

      this.textarea.value = toggled.text;
      this.textarea.setSelectionRange(toggled.cursorStart, toggled.cursorEnd);
      this.model.updateDraft(toggled.text, toggled.cursorStart, toggled.cursorEnd);
      this.#selectedReferenceIndex = this.#inferSelectedReferenceIndex(toggled.cursorStart, toggled.cursorEnd);
      this.#render({ preserveTextareaValue: true });
      this.#emitOverlays();
      return;
    }

    if (e.key === "Tab") {
      // Excel-like behavior: Tab/Shift+Tab commits the edit (and the app navigates selection).
      // Exception: plain Tab accepts an AI suggestion if one is available.
      //
      // Never allow default browser focus traversal while editing.
      if (!e.shiftKey) {
        const accepted = this.model.acceptAiSuggestion();
        if (accepted) {
          e.preventDefault();
          this.#selectedReferenceIndex = null;
          this.#render({ preserveTextareaValue: false });
          this.#setTextareaSelectionFromModel();
          this.#emitOverlays();
          return;
        }
      }

      e.preventDefault();
      this.#commit({ reason: "tab", shift: e.shiftKey });
      return;
    }

    if (e.key === "Escape") {
      e.preventDefault();
      this.#cancel();
      return;
    }

    // Excel behavior: Enter commits, Alt+Enter inserts newline.
    if (e.key === "Enter" && !e.altKey) {
      e.preventDefault();
      this.#commit({ reason: "enter", shift: e.shiftKey });
      return;
    }
  }

  #cancel(): void {
    if (!this.model.isEditing) return;
    this.#closeFunctionPicker({ restoreFocus: false });
    this.textarea.blur();
    this.model.cancel();
    this.#hoverOverride = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    this.#callbacks.onCancel?.();
    this.#emitOverlays();
  }

  #commit(commit: FormulaBarCommit): void {
    if (!this.model.isEditing) return;
    this.#closeFunctionPicker({ restoreFocus: false });
    this.textarea.blur();
    const committed = this.model.commit();
    this.#hoverOverride = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    this.#callbacks.onCommit(committed, commit);
    this.#emitOverlays();
  }

  #focusFx(): void {
    // If the formula bar isn't mounted, avoid stealing focus (and avoid creating global pickers).
    if (!this.root.isConnected) return;

    // Excel-style: clicking fx focuses the formula input and commonly starts a formula.
    if (this.model.isEditing) this.focus();
    else this.focus({ cursor: "end" });

    if (!this.model.isEditing) return;

    if (this.textarea.value.trim() === "") {
      this.textarea.value = "=";
      this.textarea.setSelectionRange(1, 1);
      this.model.updateDraft(this.textarea.value, 1, 1);
      this.#selectedReferenceIndex = null;
      this.#render({ preserveTextareaValue: true });
      this.#emitOverlays();
    }

    this.#openFunctionPicker();
  }

  #openFunctionPicker(): void {
    if (this.#functionPickerOpen) {
      this.#functionPickerInputEl.focus();
      this.#functionPickerInputEl.select();
      return;
    }
    if (!this.root.isConnected) return;
    if (!this.model.isEditing) return;

    const start = this.textarea.selectionStart ?? this.textarea.value.length;
    const end = this.textarea.selectionEnd ?? this.textarea.value.length;
    this.#functionPickerAnchorSelection = { start, end };

    this.#functionPickerOpen = true;
    this.#functionPickerEl.hidden = false;
    this.#functionPickerInputEl.value = "";
    this.#functionPickerSelectedIndex = 0;

    this.#positionFunctionPicker();
    this.#renderFunctionPickerResults();

    document.addEventListener("mousedown", this.#functionPickerDocMouseDown, true);

    this.#functionPickerInputEl.focus();
    this.#functionPickerInputEl.select();
  }

  #closeFunctionPicker(opts: { restoreFocus: boolean } = { restoreFocus: true }): void {
    if (!this.#functionPickerOpen) return;
    this.#functionPickerOpen = false;
    this.#functionPickerEl.hidden = true;
    this.#functionPickerItems = [];
    this.#functionPickerItemEls = [];
    this.#functionPickerSelectedIndex = 0;
    const anchor = this.#functionPickerAnchorSelection;
    this.#functionPickerAnchorSelection = null;

    document.removeEventListener("mousedown", this.#functionPickerDocMouseDown, true);

    if (!opts.restoreFocus) return;
    if (!this.root.isConnected) return;

    try {
      this.textarea.focus({ preventScroll: true });
    } catch {
      this.textarea.focus();
    }

    if (anchor) {
      this.textarea.setSelectionRange(anchor.start, anchor.end);
      this.model.updateDraft(this.textarea.value, anchor.start, anchor.end);
      this.#selectedReferenceIndex = this.#inferSelectedReferenceIndex(anchor.start, anchor.end);
      this.#render({ preserveTextareaValue: true });
      this.#emitOverlays();
    }
  }

  #onFunctionPickerDocMouseDown(e: MouseEvent): void {
    if (!this.#functionPickerOpen) return;
    const target = e.target as Node | null;
    if (!target) return;
    if (this.#functionPickerEl.contains(target)) return;
    if (this.#fxButtonEl.contains(target)) return;
    // Clicking outside should close the popover without stealing focus from the clicked surface.
    this.#closeFunctionPicker({ restoreFocus: false });
  }

  #positionFunctionPicker(): void {
    // Anchor below the fx button by default.
    const rootRect = this.root.getBoundingClientRect();
    const fxRect = this.#fxButtonEl.getBoundingClientRect();
    const top = fxRect.bottom - rootRect.top + 6;
    const left = fxRect.left - rootRect.left;
    this.#functionPickerEl.style.top = `${Math.max(0, Math.round(top))}px`;
    this.#functionPickerEl.style.left = `${Math.max(0, Math.round(left))}px`;
  }

  #onFunctionPickerInput(): void {
    if (!this.#functionPickerOpen) return;
    this.#functionPickerSelectedIndex = 0;
    this.#renderFunctionPickerResults();
  }

  #onFunctionPickerKeyDown(e: KeyboardEvent): void {
    if (!this.#functionPickerOpen) return;

    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      this.#closeFunctionPicker({ restoreFocus: true });
      return;
    }

    if (e.key === "ArrowDown") {
      e.preventDefault();
      e.stopPropagation();
      this.#updateFunctionPickerSelection(this.#functionPickerSelectedIndex + 1);
      return;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      e.stopPropagation();
      this.#updateFunctionPickerSelection(this.#functionPickerSelectedIndex - 1);
      return;
    }

    if (e.key === "Enter") {
      e.preventDefault();
      e.stopPropagation();
      this.#selectFunctionPickerItem(this.#functionPickerSelectedIndex);
    }
  }

  #updateFunctionPickerSelection(nextIndex: number): void {
    if (this.#functionPickerItems.length === 0) {
      this.#functionPickerSelectedIndex = 0;
      return;
    }

    const clamped = Math.max(0, Math.min(nextIndex, this.#functionPickerItems.length - 1));
    const prev = this.#functionPickerSelectedIndex;
    this.#functionPickerSelectedIndex = clamped;

    const prevEl = this.#functionPickerItemEls[prev];
    if (prevEl) prevEl.setAttribute("aria-selected", "false");

    const nextEl = this.#functionPickerItemEls[clamped];
    if (nextEl) {
      nextEl.setAttribute("aria-selected", "true");
      if (typeof nextEl.scrollIntoView === "function") nextEl.scrollIntoView({ block: "nearest" });
    }
  }

  #selectFunctionPickerItem(index: number): void {
    const item = this.#functionPickerItems[index];
    if (!item) return;
    const anchor = this.#functionPickerAnchorSelection;
    if (!anchor) return;

    this.#closeFunctionPicker({ restoreFocus: false });
    this.#insertFunctionAtSelection(item.name, anchor);
  }

  #insertFunctionAtSelection(name: string, selection: { start: number; end: number }): void {
    if (!this.root.isConnected) return;
    if (!this.model.isEditing) return;

    const prevText = this.textarea.value;
    const start = Math.max(0, Math.min(selection.start, prevText.length));
    const end = Math.max(0, Math.min(selection.end, prevText.length));

    const insert = `${name}(`;
    const nextText = prevText.slice(0, start) + insert + prevText.slice(end);
    const cursor = start + insert.length;

    this.textarea.value = nextText;
    try {
      this.textarea.focus({ preventScroll: true });
    } catch {
      this.textarea.focus();
    }
    this.textarea.setSelectionRange(cursor, cursor);
    this.model.updateDraft(nextText, cursor, cursor);
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: true });
    this.#emitOverlays();
  }

  #renderFunctionPickerResults(): void {
    const query = this.#functionPickerInputEl.value;
    const limit = 50;

    const items: FunctionPickerItem[] = query.trim()
      ? searchFunctionResults(query, { limit }).map((res) => ({
          name: res.name,
          signature: res.signature,
          summary: res.summary,
        }))
      : DEFAULT_FUNCTION_NAMES.slice(0, limit).map((name) => functionPickerItemFromName(name));

    this.#functionPickerItems = items;
    this.#functionPickerItemEls = [];
    this.#functionPickerListEl.innerHTML = "";

    if (items.length === 0) {
      const empty = document.createElement("li");
      empty.className = "command-palette__empty";
      empty.textContent = query.trim() ? "No matching functions" : "Type to search functions";
      empty.setAttribute("role", "presentation");
      this.#functionPickerListEl.appendChild(empty);
      return;
    }

    for (let i = 0; i < items.length; i += 1) {
      const fn = items[i]!;
      const li = document.createElement("li");
      li.className = "command-palette__item";
      li.setAttribute("role", "option");
      li.setAttribute("aria-selected", i === this.#functionPickerSelectedIndex ? "true" : "false");

      const icon = document.createElement("div");
      icon.className = "command-palette__item-icon command-palette__item-icon--function";
      icon.textContent = "fx";

      const main = document.createElement("div");
      main.className = "command-palette__item-main";

      const label = document.createElement("div");
      label.className = "command-palette__item-label";
      label.textContent = fn.name;

      main.appendChild(label);

      const signatureOrSummary = (() => {
        const summary = fn.summary?.trim?.() ?? "";
        const signature = fn.signature?.trim?.() ?? "";
        if (signature && summary) return `${signature} — ${summary}`;
        if (signature) return signature;
        if (summary) return summary;
        return "";
      })();

      if (signatureOrSummary) {
        const desc = document.createElement("div");
        desc.className = "command-palette__item-description command-palette__item-description--mono";
        desc.textContent = signatureOrSummary;
        main.appendChild(desc);
      }

      li.appendChild(icon);
      li.appendChild(main);

      li.addEventListener("mousedown", (e) => {
        // Keep focus in the search input so we can handle selection consistently.
        e.preventDefault();
      });
      li.addEventListener("click", () => this.#selectFunctionPickerItem(i));

      this.#functionPickerListEl.appendChild(li);
      this.#functionPickerItemEls.push(li);
    }

    // Ensure selection is valid after query changes.
    this.#updateFunctionPickerSelection(this.#functionPickerSelectedIndex);
  }

  #render(opts: { preserveTextareaValue: boolean }): void {
    // If we're rendering synchronously (e.g. commit/cancel/AI suggestion), cancel any
    // pending scheduled render so we don't immediately re-render the same state.
    this.#cancelPendingRender();

    if (!this.model.isEditing) {
      this.#selectedReferenceIndex = null;
    }

    if (document.activeElement !== this.#addressEl) {
      this.#addressEl.value = this.model.activeCell.address;
    }

    if (!opts.preserveTextareaValue) {
      this.textarea.value = this.model.draft;
    }

    const showEditingActions = this.model.isEditing;
    this.#cancelButtonEl.hidden = !showEditingActions;
    this.#cancelButtonEl.disabled = !showEditingActions;
    this.#commitButtonEl.hidden = !showEditingActions;
    this.#commitButtonEl.disabled = !showEditingActions;

    const cursor = this.model.cursorStart;
    const ghost = this.model.isEditing ? this.model.aiGhostText() : "";
    const previewRaw = this.model.isEditing ? this.model.aiSuggestionPreview() : null;
    const previewText = ghost && previewRaw != null ? formatPreview(previewRaw) : "";
    let ghostInserted = false;
    let previewInserted = false;
    let highlightHtml = "";

    const isFormulaEditing = this.model.isEditing && this.model.draft.trim().startsWith("=");
    const referenceBySpanKey = new Map<string, { color: string; index: number; active: boolean }>();
    if (isFormulaEditing) {
      const activeIndex = this.model.activeReferenceIndex();
      for (const ref of this.model.coloredReferences()) {
        referenceBySpanKey.set(`${ref.start}:${ref.end}`, {
          color: ref.color,
          index: ref.index,
          active: activeIndex === ref.index
        });
      }
    }

    const renderSpan = (span: { kind: string; start: number; end: number }, text: string): string => {
      if (!isFormulaEditing) {
        return `<span data-kind="${span.kind}">${escapeHtml(text)}</span>`;
      }
      const meta = referenceBySpanKey.get(`${span.start}:${span.end}`);
      if (!meta) {
        return `<span data-kind="${span.kind}">${escapeHtml(text)}</span>`;
      }
      const activeClass = meta.active ? " formula-bar-reference--active" : "";
      return `<span data-kind="${span.kind}" data-ref-index="${meta.index}" class="formula-bar-reference${activeClass}" style="color: ${meta.color};">${escapeHtml(
        text
      )}</span>`;
    };

    for (const span of this.model.highlightedSpans()) {
      if (!ghostInserted && ghost && cursor <= span.start) {
        highlightHtml += `<span class="formula-bar-ghost">${escapeHtml(ghost)}</span>`;
        if (previewText && !previewInserted) {
          highlightHtml += `<span class="formula-bar-preview">${escapeHtml(previewText)}</span>`;
          previewInserted = true;
        }
        ghostInserted = true;
      }

      if (!ghostInserted && ghost && cursor > span.start && cursor < span.end) {
        const split = cursor - span.start;
        const before = span.text.slice(0, split);
        const after = span.text.slice(split);
        if (before) {
          highlightHtml += renderSpan(span, before);
        }
        highlightHtml += `<span class="formula-bar-ghost">${escapeHtml(ghost)}</span>`;
        if (previewText && !previewInserted) {
          highlightHtml += `<span class="formula-bar-preview">${escapeHtml(previewText)}</span>`;
          previewInserted = true;
        }
        ghostInserted = true;
        if (after) {
          highlightHtml += renderSpan(span, after);
        }
        continue;
      }

      highlightHtml += renderSpan(span, span.text);
    }

    if (!ghostInserted && ghost) {
      highlightHtml += `<span class="formula-bar-ghost">${escapeHtml(ghost)}</span>`;
      if (previewText && !previewInserted) {
        highlightHtml += `<span class="formula-bar-preview">${escapeHtml(previewText)}</span>`;
        previewInserted = true;
      }
    }
    // Avoid forcing a full DOM re-parse/layout if the highlight HTML is unchanged.
    if (highlightHtml !== this.#lastHighlightHtml) {
      this.#highlightEl.innerHTML = highlightHtml;
      this.#lastHighlightHtml = highlightHtml;
    }

    // Toggle editing UI state (textarea visibility, hover hit-testing, etc.) through CSS classes.
    this.root.classList.toggle("formula-bar--editing", this.model.isEditing);

    const hint = this.model.functionHint();
    this.#hintEl.replaceChildren();
    if (hint) {
      const panel = document.createElement("div");
      panel.className = "formula-bar-hint-panel";

      const title = document.createElement("div");
      title.className = "formula-bar-hint-title";
      title.textContent = "PARAMETERS";

      const body = document.createElement("div");
      body.className = "formula-bar-hint-body";

      const signature = document.createElement("span");
      signature.className = "formula-bar-hint-signature";

      for (const part of hint.parts) {
        const token = document.createElement("span");
        token.className = `formula-bar-hint-token formula-bar-hint-token--${part.kind}`;
        token.dataset.kind = part.kind;
        token.textContent = part.text;
        signature.appendChild(token);
      }

      body.appendChild(signature);

      const summary = hint.signature.summary?.trim?.() ?? "";
      if (summary) {
        const sep = document.createElement("span");
        sep.className = "formula-bar-hint-summary-separator";
        sep.textContent = " — ";

        const summaryEl = document.createElement("span");
        summaryEl.className = "formula-bar-hint-summary";
        summaryEl.textContent = summary;

        body.appendChild(sep);
        body.appendChild(summaryEl);
      }

      panel.appendChild(title);
      panel.appendChild(body);
      this.#hintEl.appendChild(panel);
    }

    const explanation = this.model.errorExplanation();
    if (!explanation) {
      this.root.classList.toggle("formula-bar--has-error", false);
      this.#setErrorPanelOpen(false);
      this.#errorPanel.textContent = "";
    } else {
      const address = this.model.activeCell.address;
      this.root.classList.toggle("formula-bar--has-error", true);
      this.#errorPanel.innerHTML = `
        <div class="formula-bar-error-title">${explanation.code} (${escapeHtml(address)}): ${explanation.title}</div>
        <div class="formula-bar-error-desc">${explanation.description}</div>
        <ul class="formula-bar-error-suggestions">${explanation.suggestions.map((s) => `<li>${s}</li>`).join("")}</ul>
      `;
    }

    this.#syncScroll();
    this.#adjustHeight();
  }

  #setErrorPanelOpen(open: boolean): void {
    this.#isErrorPanelOpen = open;
    this.root.classList.toggle("formula-bar--error-panel-open", open);
    this.#errorButton.setAttribute("aria-expanded", open ? "true" : "false");
  }

  #adjustHeight(): void {
    const minHeight = 24;
    const maxHeight = 140;

    const highlightEl = this.#highlightEl as HTMLElement;

    if (this.model.isEditing) {
      // Reset before measuring.
      this.textarea.style.height = `${minHeight}px`;
      const desired = Math.max(minHeight, Math.min(maxHeight, this.textarea.scrollHeight));
      this.textarea.style.height = `${desired}px`;
      highlightEl.style.height = `${desired}px`;
      return;
    }

    // In view mode, the textarea is hidden, so measure the highlighted <pre>.
    highlightEl.style.height = `${minHeight}px`;
    const desired = Math.max(minHeight, Math.min(maxHeight, highlightEl.scrollHeight));
    highlightEl.style.height = `${desired}px`;
  }

  #syncScroll(): void {
    (this.#highlightEl as HTMLElement).scrollTop = this.textarea.scrollTop;
    (this.#highlightEl as HTMLElement).scrollLeft = this.textarea.scrollLeft;
  }

  #setTextareaSelectionFromModel(): void {
    if (!this.model.isEditing) return;
    const start = this.model.cursorStart;
    const end = this.model.cursorEnd;
    this.textarea.setSelectionRange(start, end);
  }

  #emitOverlays(): void {
    const range = this.#hoverOverride ?? this.model.hoveredReference();
    this.#callbacks.onHoverRange?.(range);

    if (!this.model.isEditing || !this.model.draft.trim().startsWith("=")) {
      this.#callbacks.onReferenceHighlights?.([]);
      return;
    }
    this.#callbacks.onReferenceHighlights?.(this.model.referenceHighlights());
  }

  #onHighlightHover(e: MouseEvent): void {
    if (this.model.isEditing) return;
    const target = e.target as HTMLElement | null;
    const span = target?.closest?.("span[data-kind]") as HTMLElement | null;
    if (!span) {
      this.#clearHoverOverride();
      return;
    }
    const kind = span.dataset.kind;
    if (kind !== "reference") {
      this.#clearHoverOverride();
      return;
    }

    const text = span.textContent ?? "";
    this.#hoverOverride = text ? parseSheetQualifiedA1Range(text) : null;
    this.#callbacks.onHoverRange?.(this.#hoverOverride);
  }

  #clearHoverOverride(): void {
    if (this.#hoverOverride === null) return;
    this.#hoverOverride = null;
    this.#emitOverlays();
  }

  #inferSelectedReferenceIndex(start: number, end: number): number | null {
    if (!this.model.isEditing || !this.model.draft.trim().startsWith("=")) return null;
    if (start === end) return null;
    for (const ref of this.model.coloredReferences()) {
      if (ref.start === start && ref.end === end) return ref.index;
    }
    return null;
  }
}

function escapeHtml(text: string): string {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function formatPreview(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "string") return ` ${value}`;
  return ` ${String(value)}`;
}

function functionPickerItemFromName(name: string): FunctionPickerItem {
  const sig = getFunctionSignature(name);
  const signature = sig ? formatSignature(sig) : undefined;
  const summary = sig?.summary?.trim?.() ? sig.summary.trim() : undefined;
  return { name, signature, summary };
}

function formatSignature(sig: FunctionSignature): string {
  const params = sig.params.map((param) => (param.optional ? `[${param.name}]` : param.name)).join(", ");
  return `${sig.name}(${params})`;
}
