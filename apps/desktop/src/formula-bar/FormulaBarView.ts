import { FormulaBarModel, type FormulaBarAiSuggestion } from "./FormulaBarModel.js";
import { type ErrorExplanation } from "./errors.js";
import { type RangeAddress } from "../spreadsheet/a1.js";
import { parseSheetQualifiedA1Range } from "./parseSheetQualifiedA1Range.js";
import {
  assignFormulaReferenceColors,
  extractFormulaReferences,
  toggleA1AbsoluteAtCursor,
  type FormulaReferenceRange,
} from "@formula/spreadsheet-frontend";
import type { EngineClient, FormulaParseOptions } from "@formula/engine";
import { ContextMenu, type ContextMenuItem } from "../menus/contextMenu.js";
import { getActiveArgumentSpan } from "./highlight/activeArgument.js";
import { FormulaBarFunctionAutocompleteController } from "./completion/functionAutocomplete.js";
import { computeFormulaIndentation } from "./computeFormulaIndentation.js";
import { buildFunctionPickerItems, renderFunctionPickerList, type FunctionPickerItem } from "./functionPicker.js";

export type FixFormulaErrorWithAiInfo = {
  address: string;
  /** The committed formula text currently stored in the active cell. */
  input: string;
  /** The current formula bar draft (may differ from `input` while editing). */
  draft: string;
  value: unknown;
  explanation: ErrorExplanation;
};

type FormulaReferenceHighlight = {
  range: FormulaReferenceRange;
  color: string;
  text: string;
  index: number;
  active?: boolean;
};

export type NameBoxMenuItem = {
  /**
   * User-visible label. For named ranges/tables this is typically the workbook-defined name.
   */
  label: string;
  /**
   * Navigation reference. When provided, selecting the item behaves like typing the reference
   * into the name box + Enter.
   *
   * When omitted/null, selection falls back to populating + selecting the name box input text.
   */
  reference?: string | null;
  enabled?: boolean;
};

const FORMULA_BAR_EXPANDED_STORAGE_KEY = "formula:ui:formulaBarExpanded";
const FORMULA_BAR_MIN_HEIGHT = 24;
const FORMULA_BAR_MAX_HEIGHT_COLLAPSED = 140;
const FORMULA_BAR_MAX_HEIGHT_EXPANDED = 360;

let formulaBarExpandedFallback: boolean | null = null;

function getFormulaBarSessionStorage(): Storage | null {
  try {
    if (typeof window === "undefined") return null;
    return window.sessionStorage ?? null;
  } catch {
    return null;
  }
}

function loadFormulaBarExpandedState(): boolean {
  const storage = getFormulaBarSessionStorage();
  if (storage) {
    try {
      const raw = storage.getItem(FORMULA_BAR_EXPANDED_STORAGE_KEY);
      if (raw === "true") return true;
      if (raw === "false") return false;
    } catch {
      // ignore
    }
  }
  return formulaBarExpandedFallback ?? false;
}

function storeFormulaBarExpandedState(expanded: boolean): void {
  const storage = getFormulaBarSessionStorage();
  if (storage) {
    try {
      storage.setItem(FORMULA_BAR_EXPANDED_STORAGE_KEY, expanded ? "true" : "false");
      return;
    } catch {
      // Fall back to in-memory storage.
    }
  }

  formulaBarExpandedFallback = expanded;
}

export interface FormulaBarViewCallbacks {
  onBeginEdit?: (activeCellAddress: string) => void;
  onCommit: (text: string, commit: FormulaBarCommit) => void;
  onCancel?: () => void;
  onGoTo?: (reference: string) => void;
  onOpenNameBoxMenu?: () => void | Promise<void>;
  getNameBoxMenuItems?: () => NameBoxMenuItem[];
  onHoverRange?: (range: RangeAddress | null) => void;
  /**
   * Like `onHoverRange`, but includes the original reference text (e.g. `Sheet2!A1:B2`)
   * so consumers can display a label and/or enforce sheet-qualified preview behavior.
   */
  onHoverRangeWithText?: (range: RangeAddress | null, refText: string | null) => void;
  onReferenceHighlights?: (
    highlights: Array<{ range: FormulaReferenceRange; color: string; text: string; index: number; active?: boolean }>
  ) => void;
  onFixFormulaErrorWithAi?: (info: FixFormulaErrorWithAiInfo) => void;
}

let errorPanelIdCounter = 0;
function nextErrorPanelId(): string {
  errorPanelIdCounter += 1;
  return `formula-bar-error-panel-${errorPanelIdCounter}`;
}

export type FormulaBarCommitReason = "enter" | "tab" | "command";

export interface FormulaBarCommit {
  reason: FormulaBarCommitReason;
  /**
   * Shift modifier for enter/tab (Shift+Enter moves up, Shift+Tab moves left).
   */
  shift: boolean;
}

export type FormulaBarViewToolingOptions = {
  /**
   * Returns the current WASM engine instance (may be null while the worker/WASM is still loading).
   *
   * When present, the formula bar will use `lexFormulaPartial` / `parseFormulaPartial` for syntax
   * highlighting, function parameter hints, and syntax error spans.
   */
  getWasmEngine?: () => EngineClient | null;
  /**
   * Formula locale id (e.g. "en-US", "de-DE"). Defaults to `document.documentElement.lang` and then "en-US".
   */
  getLocaleId?: () => string;
  referenceStyle?: NonNullable<FormulaParseOptions["referenceStyle"]>;
};

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
  #lastHighlightDraft: string | null = null;
  #lastHighlightIsFormulaEditing = false;
  #lastHighlightHadGhost = false;
  #lastActiveReferenceIndex: number | null = null;
  #lastHighlightSpans: ReturnType<FormulaBarModel["highlightedSpans"]> | null = null;

  #argumentPreviewProvider: ((expr: string) => unknown | Promise<unknown>) | null = null;
  #argumentPreviewKey: string | null = null;
  #argumentPreviewValue: unknown | null = null;
  #argumentPreviewPending = false;
  #argumentPreviewTimer: ReturnType<typeof setTimeout> | null = null;
  #argumentPreviewRequestId = 0;

  #functionAutocomplete: FormulaBarFunctionAutocompleteController;
  #nameBoxDropdownEl: HTMLButtonElement;
  #cancelButtonEl: HTMLButtonElement;
  #commitButtonEl: HTMLButtonElement;
  #fxButtonEl: HTMLButtonElement;
  #expandButtonEl: HTMLButtonElement;
  #addressEl: HTMLInputElement;
  #highlightEl: HTMLElement;
  #hintEl: HTMLElement;
  #errorButton: HTMLButtonElement;
  #errorPanel: HTMLElement;
  #errorTitleEl: HTMLElement;
  #errorDescEl: HTMLElement;
  #errorSuggestionsEl: HTMLUListElement;
  #errorFixAiButton: HTMLButtonElement;
  #errorShowRangesButton: HTMLButtonElement;
  #errorCloseButton: HTMLButtonElement;
  #isErrorPanelOpen = false;
  #errorPanelReferenceHighlights: FormulaReferenceHighlight[] | null = null;
  #hoverOverride: RangeAddress | null = null;
  #hoverOverrideText: string | null = null;
  #selectedReferenceIndex: number | null = null;
  #mouseDownSelectedReferenceIndex: number | null = null;
  #nameBoxValue = "A1";
  #isExpanded = false;
  #callbacks: FormulaBarViewCallbacks;
  #tooling: FormulaBarViewToolingOptions | null = null;
  #toolingRequestId = 0;
  #toolingScheduled:
    | { id: number; kind: "raf" }
    | { id: ReturnType<typeof setTimeout>; kind: "timeout" }
    | null = null;
  #toolingPending: {
    requestId: number;
    draft: string;
    cursor: number;
    localeId: string;
    referenceStyle: NonNullable<FormulaParseOptions["referenceStyle"]>;
  } | null = null;
  #nameBoxMenu: ContextMenu | null = null;
  #restoreNameBoxFocusOnMenuClose = false;
  #nameBoxMenuEscapeListener: ((e: KeyboardEvent) => void) | null = null;

  #functionPickerEl: HTMLDivElement;
  #functionPickerInputEl: HTMLInputElement;
  #functionPickerListEl: HTMLUListElement;
  #functionPickerOpen = false;
  #functionPickerItems: FunctionPickerItem[] = [];
  #functionPickerItemEls: HTMLLIElement[] = [];
  #functionPickerSelectedIndex = 0;
  #functionPickerAnchorSelection: { start: number; end: number } | null = null;
  #functionPickerDocMouseDown = (e: MouseEvent) => this.#onFunctionPickerDocMouseDown(e);

  constructor(root: HTMLElement, callbacks: FormulaBarViewCallbacks, tooling?: FormulaBarViewToolingOptions) {
    this.root = root;
    this.#callbacks = callbacks;
    this.#tooling = tooling ?? null;

    root.classList.add("formula-bar");
    this.#isExpanded = loadFormulaBarExpandedState();
    root.classList.toggle("formula-bar--expanded", this.#isExpanded);

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
    nameBoxDropdown.setAttribute("aria-haspopup", "menu");
    nameBoxDropdown.setAttribute("aria-expanded", "false");

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

    const expandButton = document.createElement("button");
    expandButton.className = "formula-bar-expand-button";
    expandButton.type = "button";
    expandButton.dataset.testid = "formula-expand-button";
    // Label/text synced after elements are assigned.
    expandButton.textContent = "▾";
    expandButton.title = "Expand formula bar";
    expandButton.setAttribute("aria-label", "Expand formula bar");
    expandButton.setAttribute("aria-pressed", "false");

    const errorButton = document.createElement("button");
    errorButton.className = "formula-bar-error-button";
    errorButton.type = "button";
    errorButton.textContent = "!";
    errorButton.title = "Show error details";
    errorButton.setAttribute("aria-label", "Show formula error");
    errorButton.setAttribute("aria-expanded", "false");
    errorButton.setAttribute("aria-haspopup", "dialog");
    errorButton.dataset.testid = "formula-error-button";
    errorButton.hidden = true;
    errorButton.disabled = true;

    const errorPanel = document.createElement("div");
    errorPanel.className = "formula-bar-error-panel";
    errorPanel.dataset.testid = "formula-error-panel";
    errorPanel.hidden = true;

    const errorPanelId = nextErrorPanelId();
    errorPanel.id = errorPanelId;
    errorButton.setAttribute("aria-controls", errorPanelId);
    errorPanel.setAttribute("role", "dialog");
    errorPanel.setAttribute("aria-modal", "false");

    const errorHeader = document.createElement("div");
    errorHeader.className = "formula-bar-error-header";

    const errorTitle = document.createElement("div");
    errorTitle.className = "formula-bar-error-title";
    errorTitle.id = `${errorPanelId}-title`;

    const errorCloseButton = document.createElement("button");
    errorCloseButton.className = "formula-bar-error-close";
    errorCloseButton.type = "button";
    errorCloseButton.textContent = "✕";
    errorCloseButton.title = "Dismiss (Esc)";
    errorCloseButton.setAttribute("aria-label", "Dismiss formula error details");
    errorCloseButton.dataset.testid = "formula-error-close";

    errorHeader.appendChild(errorTitle);
    errorHeader.appendChild(errorCloseButton);

    const errorDesc = document.createElement("div");
    errorDesc.className = "formula-bar-error-desc";
    errorDesc.id = `${errorPanelId}-desc`;

    const errorSuggestions = document.createElement("ul");
    errorSuggestions.className = "formula-bar-error-suggestions";

    const errorActions = document.createElement("div");
    errorActions.className = "formula-bar-error-actions";

    const errorFixAiButton = document.createElement("button");
    errorFixAiButton.className = "formula-bar-error-action formula-bar-error-action--primary";
    errorFixAiButton.type = "button";
    errorFixAiButton.textContent = "Fix with AI";
    errorFixAiButton.dataset.testid = "formula-error-fix-ai";

    const errorShowRangesButton = document.createElement("button");
    errorShowRangesButton.className = "formula-bar-error-action formula-bar-error-action--secondary";
    errorShowRangesButton.type = "button";
    errorShowRangesButton.textContent = "Show referenced ranges";
    errorShowRangesButton.setAttribute("aria-pressed", "false");
    errorShowRangesButton.dataset.testid = "formula-error-show-ranges";

    errorActions.appendChild(errorFixAiButton);
    errorActions.appendChild(errorShowRangesButton);

    errorPanel.appendChild(errorHeader);
    errorPanel.appendChild(errorDesc);
    errorPanel.appendChild(errorSuggestions);
    errorPanel.appendChild(errorActions);
    errorPanel.setAttribute("aria-labelledby", errorTitle.id);
    errorPanel.setAttribute("aria-describedby", errorDesc.id);

    row.appendChild(nameBox);
    row.appendChild(actions);
    row.appendChild(editor);
    row.appendChild(expandButton);
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
    this.#expandButtonEl = expandButton;
    this.#addressEl = address;
    this.#highlightEl = highlight;
    this.#hintEl = hint;
    this.#errorButton = errorButton;
    this.#errorPanel = errorPanel;
    this.#errorTitleEl = errorTitle;
    this.#errorDescEl = errorDesc;
    this.#errorSuggestionsEl = errorSuggestions;
    this.#errorFixAiButton = errorFixAiButton;
    this.#errorShowRangesButton = errorShowRangesButton;
    this.#errorCloseButton = errorCloseButton;
    this.#functionPickerEl = functionPicker;
    this.#functionPickerInputEl = functionPickerInput;
    this.#functionPickerListEl = functionPickerList;

    address.addEventListener("focus", () => {
      address.select();
    });

    nameBoxDropdown.addEventListener("click", () => {
      if (this.#callbacks.getNameBoxMenuItems) {
        this.#toggleNameBoxMenu();
        return;
      }
      if (this.#callbacks.onOpenNameBoxMenu) {
        Promise.resolve(this.#callbacks.onOpenNameBoxMenu())
          .catch((err) => {
            console.error("Failed to open name box menu:", err);
          });
        return;
      }

      // Fallback affordance: focus the address input so keyboard "Go To" still feels natural.
      address.focus();
    });

    address.addEventListener("keydown", (e) => {
      // Excel-style name box dropdown affordance.
      if (
        (e.key === "ArrowDown" && e.altKey && !e.ctrlKey && !e.metaKey) ||
        (e.key === "F4" && !e.altKey && !e.ctrlKey && !e.metaKey)
      ) {
        e.preventDefault();
        if (this.#callbacks.getNameBoxMenuItems) {
          this.#toggleNameBoxMenu();
        } else if (this.#callbacks.onOpenNameBoxMenu) {
          void this.#callbacks.onOpenNameBoxMenu();
        } else {
          this.#toggleNameBoxMenu();
        }
        return;
      }

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
        address.value = this.#nameBoxValue;
        address.blur();
      }
    });

    textarea.addEventListener("focus", () => this.#beginEditFromFocus());
    textarea.addEventListener("input", () => this.#onInputOrSelection());
    textarea.addEventListener("mousedown", () => this.#onTextareaMouseDown());
    textarea.addEventListener("click", () => this.#onTextareaClick());
    textarea.addEventListener("keyup", () => this.#onInputOrSelection());
    textarea.addEventListener("select", () => this.#onInputOrSelection());
    textarea.addEventListener("scroll", () => this.#syncScroll());
    textarea.addEventListener("keydown", (e) => this.#onKeyDown(e));

    // Non-AI function autocomplete dropdown (Excel-like).
    // Mount after registering FormulaBarView's own listeners so focus/input updates keep the model in sync first.
    this.#functionAutocomplete = new FormulaBarFunctionAutocompleteController({ formulaBar: this, anchor: editor });

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

    errorCloseButton.addEventListener("click", () => this.#setErrorPanelOpen(false, { restoreFocus: true }));
    errorPanel.addEventListener("keydown", (e) => this.#onErrorPanelKeyDown(e));
    errorFixAiButton.addEventListener("click", () => this.#fixFormulaErrorWithAi());
    errorShowRangesButton.addEventListener("click", () => this.#toggleErrorReferenceHighlights());

    cancelButton.addEventListener("click", () => this.#cancel());
    commitButton.addEventListener("click", () => this.#commit({ reason: "command", shift: false }));
    fxButton.addEventListener("click", () => this.#focusFx());
    fxButton.addEventListener("mousedown", (e) => {
      // Preserve the caret/selection in the textarea when clicking the fx button.
      e.preventDefault();
    });

    expandButton.addEventListener("click", () => this.#toggleExpanded());
    expandButton.addEventListener("mousedown", (e) => {
      // Preserve the caret/selection in the textarea when clicking the toggle button.
      e.preventDefault();
    });

    functionPickerInput.addEventListener("input", () => this.#onFunctionPickerInput());
    const pickerKeyDown = (e: KeyboardEvent) => this.#onFunctionPickerKeyDown(e);
    functionPickerInput.addEventListener("keydown", pickerKeyDown);
    functionPickerList.addEventListener("keydown", pickerKeyDown);

    this.#syncExpandedUi();

    // Initial render.
    this.model.setActiveCell({ address: "A1", input: "", value: "" });
    this.#render({ preserveTextareaValue: false });
  }

  #toggleNameBoxMenu(): void {
    const menu = (this.#nameBoxMenu ??= new ContextMenu({
      testId: "name-box-menu",
      onClose: () => {
        this.#nameBoxDropdownEl.setAttribute("aria-expanded", "false");
        if (this.#nameBoxMenuEscapeListener) {
          window.removeEventListener("keydown", this.#nameBoxMenuEscapeListener, true);
          this.#nameBoxMenuEscapeListener = null;
        }

        if (this.#restoreNameBoxFocusOnMenuClose) {
          this.#restoreNameBoxFocusOnMenuClose = false;
          try {
            this.#addressEl.focus({ preventScroll: true });
          } catch {
            this.#addressEl.focus();
          }
          this.#addressEl.select();
        }
      },
    }));

    if (menu.isOpen()) {
      this.#restoreNameBoxFocusOnMenuClose = true;
      menu.close();
      return;
    }

    // Track Esc-driven closes so we can restore focus without stealing it on outside clicks.
    this.#restoreNameBoxFocusOnMenuClose = false;
    if (this.#nameBoxMenuEscapeListener) {
      window.removeEventListener("keydown", this.#nameBoxMenuEscapeListener, true);
      this.#nameBoxMenuEscapeListener = null;
    }
    this.#nameBoxMenuEscapeListener = (e: KeyboardEvent) => {
      const isEscape =
        e.key === "Escape" || e.key === "Esc" || e.code === "Escape" || (e as unknown as { keyCode?: number }).keyCode === 27;
      if (!isEscape) return;
      if (!menu.isOpen()) return;
      this.#restoreNameBoxFocusOnMenuClose = true;
    };
    window.addEventListener("keydown", this.#nameBoxMenuEscapeListener, true);

    const rawItems = this.#callbacks.getNameBoxMenuItems?.() ?? [];
    const items: ContextMenuItem[] = [];

    for (const item of rawItems) {
      const label = String(item?.label ?? "").trim();
      if (!label) continue;
      const enabled = item.enabled ?? true;
      const reference = typeof item.reference === "string" ? item.reference.trim() : item.reference ?? null;

      items.push({
        type: "item",
        label,
        enabled,
        onSelect: () => {
          if (reference) {
            this.#callbacks.onGoTo?.(reference);
            return;
          }

          // Fallback: populate + select the text so the user can edit/confirm.
          this.#addressEl.value = label;
          try {
            this.#addressEl.focus({ preventScroll: true });
          } catch {
            this.#addressEl.focus();
          }
          this.#addressEl.select();
        },
      });
    }

    if (items.length === 0) {
      items.push({
        type: "item",
        label: "No named ranges",
        enabled: false,
        onSelect: () => {},
      });
    }

    const rect = this.#nameBoxDropdownEl.getBoundingClientRect();
    this.#nameBoxDropdownEl.setAttribute("aria-expanded", "true");
    menu.open({ x: rect.left, y: rect.bottom, items });
  }

  setArgumentPreviewProvider(provider: ((expr: string) => unknown | Promise<unknown>) | null): void {
    this.#argumentPreviewProvider = provider;
    this.#clearArgumentPreviewState();
    this.#render({ preserveTextareaValue: true });
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

  setActiveCell(info: { address: string; input: string; value: unknown; nameBox?: string }): void {
    const { nameBox, ...activeCell } = info;
    this.#nameBoxValue = nameBox ?? activeCell.address;

    // Keep the Name Box display in sync with selection changes even while editing
    // (but never clobber the user's in-progress typing in the Name Box itself).
    if (document.activeElement !== this.#addressEl) {
      this.#addressEl.value = this.#nameBoxValue;
    }

    if (this.model.isEditing) return;
    this.model.setActiveCell(activeCell);
    this.#hoverOverride = null;
    this.#hoverOverrideText = null;
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
    this.#scheduleEngineTooling();
  }

  updateRangeSelection(range: RangeAddress, sheetId?: string): void {
    this.model.updateRangeSelection(range, sheetId);
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    this.#setTextareaSelectionFromModel();
    this.#emitOverlays();
    this.#scheduleEngineTooling();
  }

  endRangeSelection(): void {
    this.model.endRangeSelection();
  }

  #beginEditFromFocus(): void {
    if (this.model.isEditing) return;
    this.#errorPanelReferenceHighlights = null;
    this.model.beginEdit();
    this.#callbacks.onBeginEdit?.(this.model.activeCell.address);
    // Hover overrides are a view-mode affordance and should not leak into editing behavior.
    this.#hoverOverride = null;
    this.#hoverOverrideText = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: true });
    this.#emitOverlays();
    this.#scheduleEngineTooling();
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
    this.#scheduleEngineTooling();
  }

  #onTextareaMouseDown(): void {
    if (!this.model.isEditing) return;

    // Clicking a textarea can collapse an existing selection *before* the `click` event
    // fires (often emitting a `select` event in between). Capture whether a full
    // reference token was selected at pointer-down time so `#onTextareaClick()` can
    // reliably implement Excel-style click-to-select / click-again-to-edit toggling.
    const start = this.textarea.selectionStart ?? this.textarea.value.length;
    const end = this.textarea.selectionEnd ?? this.textarea.value.length;
    this.model.updateDraft(this.textarea.value, start, end);
    this.#mouseDownSelectedReferenceIndex = this.#inferSelectedReferenceIndex(start, end);
  }

  #onTextareaClick(): void {
    if (!this.model.isEditing) return;

    const prevSelectedReferenceIndex = this.#mouseDownSelectedReferenceIndex ?? this.#selectedReferenceIndex;
    this.#mouseDownSelectedReferenceIndex = null;
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
    this.#scheduleEngineTooling();
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

  #scheduleEngineTooling(): void {
    // Only run editor-tooling calls while editing; the formula bar view mode already
    // uses stable highlights and we want to avoid late async updates after commit/cancel.
    if (!this.model.isEditing) return;

    // Only ask the engine to lex/parse when the draft is actually a formula.
    // This avoids surfacing parse errors while editing plain text values.
    const draft = this.model.draft;
    if (!draft.trim().startsWith("=")) return;

    const engine = this.#tooling?.getWasmEngine?.() ?? null;
    if (!engine) return;

    const localeId =
      this.#tooling?.getLocaleId?.() ??
      (typeof document !== "undefined" ? document.documentElement?.lang : "") ??
      "en-US";
    const referenceStyle = this.#tooling?.referenceStyle ?? "A1";

    const cursor = this.model.cursorStart;
    const requestId = ++this.#toolingRequestId;

    this.#toolingPending = { requestId, draft, cursor, localeId: localeId || "en-US", referenceStyle };

    // Coalesce multiple rapid edits/selection changes into one tooling request per frame.
    if (this.#toolingScheduled) return;

    const flush = (): void => {
      this.#toolingScheduled = null;
      const pending = this.#toolingPending;
      this.#toolingPending = null;
      if (!pending) return;
      void this.#runEngineTooling(pending);
    };

    if (typeof requestAnimationFrame === "function") {
      const id = requestAnimationFrame(() => flush());
      this.#toolingScheduled = { id, kind: "raf" };
    } else {
      const id = setTimeout(() => flush(), 0);
      this.#toolingScheduled = { id, kind: "timeout" };
    }
  }

  async #runEngineTooling(pending: {
    requestId: number;
    draft: string;
    cursor: number;
    localeId: string;
    referenceStyle: NonNullable<FormulaParseOptions["referenceStyle"]>;
  }): Promise<void> {
    const engine = this.#tooling?.getWasmEngine?.() ?? null;
    if (!engine) return;
    if (!pending.draft.trim().startsWith("=")) return;

    try {
      const options: FormulaParseOptions = { localeId: pending.localeId, referenceStyle: pending.referenceStyle };
      const [lexResult, parseResult] = await Promise.all([
        engine.lexFormulaPartial(pending.draft, options),
        engine.parseFormulaPartial(pending.draft, pending.cursor, options),
      ]);

      // Ignore stale/out-of-order results.
      if (pending.requestId !== this.#toolingRequestId) return;
      if (!this.model.isEditing) return;
      if (this.model.draft !== pending.draft) return;
      if (!this.model.draft.trim().startsWith("=")) return;

      this.model.applyEngineToolingResult({ formula: pending.draft, localeId: pending.localeId, lexResult, parseResult });
      this.#requestRender({ preserveTextareaValue: true });
    } catch {
      // Best-effort: if the engine worker is unavailable/uninitialized, keep the local
      // tokenizer/highlighter without surfacing errors to the user.
    }
  }

  #onKeyDown(e: KeyboardEvent): void {
    if (!this.model.isEditing) return;

    if (this.#functionAutocomplete.handleKeyDown(e)) return;

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
      this.#scheduleEngineTooling();
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
          this.#scheduleEngineTooling();
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
    if (e.key === "Enter" && e.altKey) {
      e.preventDefault();

      const prevText = this.textarea.value;
      const cursorStart = this.textarea.selectionStart ?? prevText.length;
      const cursorEnd = this.textarea.selectionEnd ?? prevText.length;

      const indentation = computeFormulaIndentation(prevText, cursorStart);
      const insertion = `\n${indentation}`;

      this.textarea.value = prevText.slice(0, cursorStart) + insertion + prevText.slice(cursorEnd);

      const nextCursor = cursorStart + insertion.length;
      this.textarea.setSelectionRange(nextCursor, nextCursor);

      // Reuse the standard input/selection path to keep the model + highlight in sync.
      this.#onInputOrSelection();
      return;
    }

    if (e.key === "Enter" && !e.altKey) {
      e.preventDefault();
      this.#commit({ reason: "enter", shift: e.shiftKey });
      return;
    }
  }

  #cancel(): void {
    if (!this.model.isEditing) return;
    this.#functionAutocomplete.close();
    this.#closeFunctionPicker({ restoreFocus: false });
    this.textarea.blur();
    this.model.cancel();
    this.#hoverOverride = null;
    this.#hoverOverrideText = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    this.#callbacks.onCancel?.();
    this.#emitOverlays();
  }

  #commit(commit: FormulaBarCommit): void {
    if (!this.model.isEditing) return;
    this.#functionAutocomplete.close();
    this.#closeFunctionPicker({ restoreFocus: false });
    this.textarea.blur();
    const committed = this.model.commit();
    this.#hoverOverride = null;
    this.#hoverOverrideText = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    this.#callbacks.onCommit(committed, commit);
    this.#emitOverlays();
  }

  #focusFx(): void {
    // If the formula bar isn't mounted, avoid stealing focus (and avoid creating global pickers).
    if (!this.root.isConnected) return;

    // Avoid overlapping UI affordances (typing autocomplete vs. explicit fx picker).
    this.#functionAutocomplete.close();

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
      this.#scheduleEngineTooling();
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
      this.#scheduleEngineTooling();
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

    const insert = `${name}()`;
    const nextText = prevText.slice(0, start) + insert + prevText.slice(end);
    // Place the caret inside the parentheses so users can immediately type arguments.
    const cursor = start + Math.max(0, insert.length - 1);

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
    this.#scheduleEngineTooling();
  }

  #renderFunctionPickerResults(): void {
    const limit = 50;
    const query = this.#functionPickerInputEl.value;
    const items: FunctionPickerItem[] = buildFunctionPickerItems(query, limit);

    this.#functionPickerItems = items;
    this.#functionPickerItemEls = renderFunctionPickerList({
      listEl: this.#functionPickerListEl,
      query,
      items,
      selectedIndex: this.#functionPickerSelectedIndex,
      onSelect: (index) => this.#selectFunctionPickerItem(index),
    });

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
      this.#addressEl.value = this.#nameBoxValue;
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
    const draft = this.model.draft;

    const isFormulaEditing = this.model.isEditing && draft.trim().startsWith("=");
    const coloredReferences = isFormulaEditing ? this.model.coloredReferences() : [];
    const activeReferenceIndex = isFormulaEditing ? this.model.activeReferenceIndex() : null;
    const highlightedSpans = this.model.highlightedSpans();

    const canFastUpdateActiveReference =
      isFormulaEditing &&
      !ghost &&
      this.#lastHighlightDraft === draft &&
      this.#lastHighlightIsFormulaEditing &&
      !this.#lastHighlightHadGhost &&
      this.#lastHighlightSpans === highlightedSpans;

    if (canFastUpdateActiveReference) {
      if (this.#lastActiveReferenceIndex !== activeReferenceIndex) {
        const prev = this.#lastActiveReferenceIndex;
        const next = activeReferenceIndex;
        if (prev != null) {
          this.#highlightEl
            .querySelectorAll(`[data-ref-index="${prev}"]`)
            .forEach((el) => el.classList.remove("formula-bar-reference--active"));
        }
        if (next != null) {
          this.#highlightEl
            .querySelectorAll(`[data-ref-index="${next}"]`)
            .forEach((el) => el.classList.add("formula-bar-reference--active"));
        }
        this.#lastActiveReferenceIndex = next;
        // We updated class attributes without rebuilding the HTML string; invalidate the
        // string cache so future full renders don't compare against a stale snapshot.
        this.#lastHighlightHtml = null;
      }
      this.#lastHighlightDraft = draft;
      this.#lastHighlightIsFormulaEditing = true;
      this.#lastHighlightHadGhost = false;
      this.#lastHighlightSpans = highlightedSpans;
    } else {
      let ghostInserted = false;
      let previewInserted = false;
      let highlightHtml = "";

      const referenceBySpanKey = new Map<string, { color: string; index: number; active: boolean }>();
      if (isFormulaEditing) {
        for (const ref of coloredReferences) {
          referenceBySpanKey.set(`${ref.start}:${ref.end}`, {
            color: ref.color,
            index: ref.index,
            active: activeReferenceIndex === ref.index,
          });
        }
      }

      const renderSpan = (
        span: { kind: string; start: number; end: number; className?: string },
        text: string
      ): string => {
        const extraClass = span.className?.trim?.() || "";
        const classAttr = (base: string | null): string => {
          const classes = [base, extraClass].filter(Boolean).join(" ").trim();
          return classes ? ` class="${classes}"` : "";
        };

        if (!isFormulaEditing) {
          return `<span data-kind="${span.kind}"${classAttr(null)}>${escapeHtml(text)}</span>`;
        }

        let meta = referenceBySpanKey.get(`${span.start}:${span.end}`) ?? null;
        if (!meta && span.kind === "reference") {
          // Engine-backed syntax error highlighting can split reference spans; preserve
          // reference colors by falling back to a containment lookup.
          const containing = coloredReferences.find((ref) => ref.start <= span.start && span.end <= ref.end) ?? null;
          if (containing) {
            meta = {
              color: containing.color,
              index: containing.index,
              active: activeReferenceIndex === containing.index,
            };
          }
        }

        if (!meta) {
          return `<span data-kind="${span.kind}"${classAttr(null)}>${escapeHtml(text)}</span>`;
        }

        const activeClass = meta.active ? " formula-bar-reference--active" : "";
        const baseClass = `formula-bar-reference${activeClass}`;
        return `<span data-kind="${span.kind}" data-ref-index="${meta.index}"${classAttr(baseClass)} style="color: ${meta.color};">${escapeHtml(
          text
        )}</span>`;
      };

      for (const span of highlightedSpans) {
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
      }

      this.#lastHighlightHtml = highlightHtml;
      this.#lastHighlightDraft = draft;
      this.#lastHighlightIsFormulaEditing = isFormulaEditing;
      this.#lastHighlightHadGhost = Boolean(ghost);
      this.#lastActiveReferenceIndex = activeReferenceIndex;
      this.#lastHighlightSpans = highlightedSpans;
    }

    // Toggle editing UI state (textarea visibility, hover hit-testing, etc.) through CSS classes.
    this.root.classList.toggle("formula-bar--editing", this.model.isEditing);

    const syntaxError = this.model.syntaxError();
    this.#hintEl.classList.toggle("formula-bar-hint--syntax-error", Boolean(syntaxError));
    const hint = this.model.functionHint();
    this.#hintEl.replaceChildren();
    if (syntaxError) {
      this.#clearArgumentPreviewState();
      const message = document.createElement("div");
      message.className = "formula-bar-hint-error";
      message.textContent = syntaxError.message;
      this.#hintEl.appendChild(message);
    }

    if (!hint) {
      this.#clearArgumentPreviewState();
    } else {
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

      if (syntaxError) {
        this.#clearArgumentPreviewState();
      } else {
        const provider = this.#argumentPreviewProvider;
        const activeArg = getActiveArgumentSpan(this.model.draft, this.model.cursorStart);
        const wantsArgPreview = Boolean(
          activeArg &&
            typeof provider === "function" &&
            typeof activeArg.argText === "string" &&
            activeArg.argText.trim() !== ""
        );

        if (wantsArgPreview && activeArg) {
          const key = `${activeArg.fnName}|${activeArg.argIndex}|${activeArg.span.start}:${activeArg.span.end}|${activeArg.argText}`;
          if (this.#argumentPreviewKey !== key) {
            this.#argumentPreviewKey = key;
            this.#argumentPreviewValue = null;
            this.#argumentPreviewPending = true;
            this.#scheduleArgumentPreviewEvaluation(activeArg, key);
          }

          const previewEl = document.createElement("div");
          previewEl.className = "formula-bar-hint-arg-preview";
          previewEl.dataset.testid = "formula-hint-arg-preview";
          previewEl.dataset.argStart = String(activeArg.span.start);
          previewEl.dataset.argEnd = String(activeArg.span.end);

          const rhs = this.#argumentPreviewPending ? "…" : formatArgumentPreviewValue(this.#argumentPreviewValue);
          previewEl.textContent = `↳ ${activeArg.argText}  →  ${rhs}`;
          body.appendChild(previewEl);
        } else {
          this.#clearArgumentPreviewState();
        }
      }

      panel.appendChild(title);
      panel.appendChild(body);
      this.#hintEl.appendChild(panel);
    }

    const explanation = this.model.errorExplanation();
    if (!explanation) {
      this.root.classList.toggle("formula-bar--has-error", false);
      this.#errorButton.hidden = true;
      this.#errorButton.disabled = true;
      this.#errorTitleEl.textContent = "";
      this.#errorDescEl.textContent = "";
      this.#errorSuggestionsEl.replaceChildren();
      this.#setErrorPanelOpen(false, { restoreFocus: false });
    } else {
      const address = this.model.activeCell.address;
      this.root.classList.toggle("formula-bar--has-error", true);
      this.#errorButton.hidden = false;
      this.#errorButton.disabled = false;
      this.#errorTitleEl.textContent = `${explanation.code} (${address}): ${explanation.title}`;
      this.#errorDescEl.textContent = explanation.description;
      this.#errorSuggestionsEl.replaceChildren(
        ...explanation.suggestions.map((s) => {
          const li = document.createElement("li");
          li.textContent = s;
          return li;
        })
      );
    }

    this.#syncErrorPanelActions();

    this.#syncScroll();
    this.#adjustHeight();
  }

  #clearArgumentPreviewState(): void {
    this.#argumentPreviewKey = null;
    this.#argumentPreviewValue = null;
    this.#argumentPreviewPending = false;
    this.#argumentPreviewRequestId += 1;
    if (this.#argumentPreviewTimer != null) {
      clearTimeout(this.#argumentPreviewTimer);
      this.#argumentPreviewTimer = null;
    }
  }

  #scheduleArgumentPreviewEvaluation(activeArg: ReturnType<typeof getActiveArgumentSpan>, key: string): void {
    if (!activeArg) return;
    const provider = this.#argumentPreviewProvider;
    if (typeof provider !== "function") return;

    // Cancel any pending timer before scheduling a new evaluation. This keeps typing responsive
    // and avoids doing work for stale cursor positions.
    if (this.#argumentPreviewTimer != null) {
      clearTimeout(this.#argumentPreviewTimer);
      this.#argumentPreviewTimer = null;
    }

    const requestId = ++this.#argumentPreviewRequestId;
    const expr = activeArg.argText;

    this.#argumentPreviewTimer = setTimeout(() => {
      this.#argumentPreviewTimer = null;

      // Allow the preview provider to be async, but bound the time we wait for it.
      const timeoutMs = 100;
      let timeoutId: ReturnType<typeof setTimeout> | null = null;
      const timeoutPromise = new Promise<unknown>((resolve) => {
        timeoutId = setTimeout(() => resolve("(preview unavailable)"), timeoutMs);
      });

      Promise.race([Promise.resolve().then(() => provider(expr)), timeoutPromise])
        .then((value) => {
          if (timeoutId != null) clearTimeout(timeoutId);
          if (requestId !== this.#argumentPreviewRequestId) return;
          if (this.#argumentPreviewKey !== key) return;
          this.#argumentPreviewPending = false;
          this.#argumentPreviewValue = value === undefined ? "(preview unavailable)" : value;
          this.#requestRender({ preserveTextareaValue: true });
        })
        .catch(() => {
          if (timeoutId != null) clearTimeout(timeoutId);
          if (requestId !== this.#argumentPreviewRequestId) return;
          if (this.#argumentPreviewKey !== key) return;
          this.#argumentPreviewPending = false;
          this.#argumentPreviewValue = "(preview unavailable)";
          this.#requestRender({ preserveTextareaValue: true });
        });
    }, 0);
  }

  #setErrorPanelOpen(open: boolean, opts: { restoreFocus: boolean } = { restoreFocus: true }): void {
    const wasOpen = this.#isErrorPanelOpen;
    this.#isErrorPanelOpen = open;
    this.root.classList.toggle("formula-bar--error-panel-open", open);
    this.#errorButton.setAttribute("aria-expanded", open ? "true" : "false");
    this.#errorPanel.hidden = !open;

    if (!open) {
      const hadReferenceHighlights = this.#errorPanelReferenceHighlights != null;
      this.#errorPanelReferenceHighlights = null;
      this.#syncErrorPanelActions();
      // Clear view-mode highlights; preserve formula-editing highlights.
      if (hadReferenceHighlights) {
        this.#callbacks.onReferenceHighlights?.(this.#currentReferenceHighlights());
      }
      if (opts.restoreFocus) {
        try {
          this.#errorButton.focus({ preventScroll: true });
        } catch {
          this.#errorButton.focus();
        }
      }
      return;
    }

    this.#syncErrorPanelActions();

    if (!wasOpen) {
      this.#focusFirstErrorPanelControl();
    }
  }

  #adjustHeight(): void {
    const minHeight = FORMULA_BAR_MIN_HEIGHT;
    const maxHeight = this.#isExpanded ? FORMULA_BAR_MAX_HEIGHT_EXPANDED : FORMULA_BAR_MAX_HEIGHT_COLLAPSED;

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

  #toggleExpanded(): void {
    this.#isExpanded = !this.#isExpanded;
    storeFormulaBarExpandedState(this.#isExpanded);
    this.#syncExpandedUi();
    this.#adjustHeight();
  }

  #syncExpandedUi(): void {
    this.root.classList.toggle("formula-bar--expanded", this.#isExpanded);
    this.#expandButtonEl.textContent = this.#isExpanded ? "▴" : "▾";
    const label = this.#isExpanded ? "Collapse formula bar" : "Expand formula bar";
    this.#expandButtonEl.title = label;
    this.#expandButtonEl.setAttribute("aria-label", label);
    this.#expandButtonEl.setAttribute("aria-pressed", this.#isExpanded ? "true" : "false");
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
    const refText = this.#hoverOverrideText ?? this.model.hoveredReferenceText();
    this.#callbacks.onHoverRangeWithText?.(range, refText ?? null);

    this.#callbacks.onReferenceHighlights?.(this.#currentReferenceHighlights());
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
    const text = span.textContent ?? "";
    if (!text) {
      this.#clearHoverOverride();
      return;
    }

    if (kind === "reference") {
      this.#hoverOverrideText = text;
      this.#hoverOverride = parseSheetQualifiedA1Range(text);
      this.#callbacks.onHoverRange?.(this.#hoverOverride);
      this.#callbacks.onHoverRangeWithText?.(this.#hoverOverride, this.#hoverOverrideText);
      return;
    }

    if (kind === "identifier") {
      const resolved = this.model.resolveNameRange(text);
      this.#hoverOverrideText = resolved ? text : null;
      this.#hoverOverride = resolved
        ? {
            start: { row: resolved.startRow, col: resolved.startCol },
            end: { row: resolved.endRow, col: resolved.endCol }
          }
        : null;
      this.#callbacks.onHoverRange?.(this.#hoverOverride);
      this.#callbacks.onHoverRangeWithText?.(this.#hoverOverride, this.#hoverOverrideText);
      return;
    }

    this.#clearHoverOverride();
  }

  #clearHoverOverride(): void {
    if (this.#hoverOverride === null && this.#hoverOverrideText === null) return;
    this.#hoverOverride = null;
    this.#hoverOverrideText = null;
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

  #fixFormulaErrorWithAi(): void {
    const explanation = this.model.errorExplanation();
    if (!explanation) return;
    this.#callbacks.onFixFormulaErrorWithAi?.({
      address: this.model.activeCell.address,
      input: this.model.activeCell.input,
      draft: this.model.draft,
      value: this.model.activeCell.value,
      explanation
    });
  }

  #toggleErrorReferenceHighlights(): void {
    if (this.#errorPanelReferenceHighlights) {
      this.#errorPanelReferenceHighlights = null;
    } else {
      this.#errorPanelReferenceHighlights = computeReferenceHighlights(this.model.draft);
      if (this.#errorPanelReferenceHighlights.length === 0) {
        this.#errorPanelReferenceHighlights = null;
      }
    }

    this.#syncErrorPanelActions();
    this.#callbacks.onReferenceHighlights?.(this.#currentReferenceHighlights());
  }

  #syncErrorPanelActions(): void {
    const explanation = this.model.errorExplanation();
    const canFix = Boolean(explanation) && typeof this.#callbacks.onFixFormulaErrorWithAi === "function";
    this.#errorFixAiButton.disabled = !canFix;

    const isFormula = this.model.draft.trim().startsWith("=");
    const isShowingRanges = this.#errorPanelReferenceHighlights != null;
    this.#errorShowRangesButton.disabled = !isFormula;
    this.#errorShowRangesButton.setAttribute("aria-pressed", isShowingRanges ? "true" : "false");
    this.#errorShowRangesButton.textContent = isShowingRanges ? "Hide referenced ranges" : "Show referenced ranges";
  }

  #currentReferenceHighlights(): FormulaReferenceHighlight[] {
    const isFormula = this.model.draft.trim().startsWith("=");
    if (this.model.isEditing && isFormula) {
      return this.model.referenceHighlights();
    }
    if (this.#errorPanelReferenceHighlights) {
      return this.#errorPanelReferenceHighlights;
    }
    return [];
  }

  #focusFirstErrorPanelControl(): void {
    const candidates: HTMLElement[] = [];
    if (!this.#errorFixAiButton.disabled) candidates.push(this.#errorFixAiButton);
    if (!this.#errorShowRangesButton.disabled) candidates.push(this.#errorShowRangesButton);
    candidates.push(this.#errorCloseButton);

    const target = candidates.find((el) => !el.hidden && !(el as HTMLButtonElement).disabled) ?? null;
    if (!target) return;
    try {
      target.focus({ preventScroll: true });
    } catch {
      target.focus();
    }
  }

  #onErrorPanelKeyDown(e: KeyboardEvent): void {
    if (!this.#isErrorPanelOpen) return;

    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      this.#setErrorPanelOpen(false, { restoreFocus: true });
      return;
    }

    if (e.key !== "Tab") return;

    const focusable = [this.#errorFixAiButton, this.#errorShowRangesButton, this.#errorCloseButton].filter(
      (el) => !el.hidden && !el.disabled
    );
    if (focusable.length === 0) return;

    const active = document.activeElement as HTMLElement | null;
    const currentIdx = active ? focusable.indexOf(active as HTMLButtonElement) : -1;
    const nextIdx = (() => {
      if (e.shiftKey) {
        if (currentIdx <= 0) return focusable.length - 1;
        return currentIdx - 1;
      }
      if (currentIdx < 0 || currentIdx === focusable.length - 1) return 0;
      return currentIdx + 1;
    })();
    const next = focusable[nextIdx]!;
    e.preventDefault();
    try {
      next.focus({ preventScroll: true });
    } catch {
      next.focus();
    }
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

function formatArgumentPreviewValue(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  return String(value);
}

function computeReferenceHighlights(text: string): FormulaReferenceHighlight[] {
  if (!text.trim().startsWith("=")) return [];
  const { references } = extractFormulaReferences(text);
  if (references.length === 0) return [];
  const { colored } = assignFormulaReferenceColors(references, null);
  return colored.map((ref) => ({
    range: ref.range,
    color: ref.color,
    text: ref.text,
    index: ref.index,
    active: false
  }));
}
