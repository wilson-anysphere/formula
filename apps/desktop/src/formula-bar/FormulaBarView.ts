import { FormulaBarModel, type FormulaBarAiSuggestion } from "./FormulaBarModel.js";
import { type RangeAddress } from "../spreadsheet/a1.js";
import { parseSheetQualifiedA1Range } from "./parseSheetQualifiedA1Range.js";
import { toggleA1AbsoluteAtCursor, type FormulaReferenceRange } from "@formula/spreadsheet-frontend";

export interface FormulaBarViewCallbacks {
  onBeginEdit?: (activeCellAddress: string) => void;
  onCommit: (text: string) => void;
  onCancel?: () => void;
  onGoTo?: (reference: string) => void;
  onHoverRange?: (range: RangeAddress | null) => void;
  onReferenceHighlights?: (
    highlights: Array<{ range: FormulaReferenceRange; color: string; text: string; index: number; active?: boolean }>
  ) => void;
}

export class FormulaBarView {
  readonly model = new FormulaBarModel();

  readonly root: HTMLElement;
  readonly textarea: HTMLTextAreaElement;

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

    root.appendChild(row);
    root.appendChild(hint);
    root.appendChild(errorPanel);

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

    address.addEventListener("focus", () => {
      address.select();
    });

    nameBoxDropdown.addEventListener("click", () => {
      // Placeholder affordance only (Excel-style name box dropdown).
      // Focus the address input so keyboard "Go To" still feels natural.
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
    commitButton.addEventListener("click", () => this.#commit());
    fxButton.addEventListener("click", () => this.#focusFx());

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

  commitEdit(): void {
    this.#commit();
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

    const start = this.textarea.selectionStart ?? this.textarea.value.length;
    const end = this.textarea.selectionEnd ?? this.textarea.value.length;
    this.model.updateDraft(this.textarea.value, start, end);
    this.#selectedReferenceIndex = this.#inferSelectedReferenceIndex(start, end);
    this.#render({ preserveTextareaValue: true });
    this.#emitOverlays();
  }

  #onTextareaClick(): void {
    if (!this.model.isEditing) return;

    const prevSelectedReferenceIndex = this.#selectedReferenceIndex;
    const start = this.textarea.selectionStart ?? this.textarea.value.length;
    const end = this.textarea.selectionEnd ?? this.textarea.value.length;
    this.model.updateDraft(this.textarea.value, start, end);
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

    this.#render({ preserveTextareaValue: true });
    this.#emitOverlays();
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

    if (e.key === "Escape") {
      e.preventDefault();
      this.#cancel();
      return;
    }

    // Excel behavior: Enter commits, Alt+Enter inserts newline.
    if (e.key === "Enter" && !e.altKey) {
      e.preventDefault();
      this.#commit();
      return;
    }
  }

  #cancel(): void {
    if (!this.model.isEditing) return;
    this.textarea.blur();
    this.model.cancel();
    this.#hoverOverride = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    this.#callbacks.onCancel?.();
    this.#emitOverlays();
  }

  #commit(): void {
    if (!this.model.isEditing) return;
    this.textarea.blur();
    const committed = this.model.commit();
    this.#hoverOverride = null;
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: false });
    this.#callbacks.onCommit(committed);
    this.#emitOverlays();
  }

  #focusFx(): void {
    // Excel-style: clicking fx focuses the formula input and commonly starts a formula.
    this.focus({ cursor: "end" });

    if (!this.model.isEditing) return;
    if (this.textarea.value.trim() !== "") return;

    this.textarea.value = "=";
    this.textarea.setSelectionRange(1, 1);
    this.model.updateDraft(this.textarea.value, 1, 1);
    this.#selectedReferenceIndex = null;
    this.#render({ preserveTextareaValue: true });
    this.#emitOverlays();
  }

  #render(opts: { preserveTextareaValue: boolean }): void {
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
      if (span.kind !== "reference" || !isFormulaEditing) {
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
    this.#highlightEl.innerHTML = highlightHtml;

    // Toggle editing UI state (textarea visibility, hover hit-testing, etc.) through CSS classes.
    this.root.classList.toggle("formula-bar--editing", this.model.isEditing);

    const hint = this.model.functionHint();
    if (!hint) {
      this.#hintEl.textContent = "";
    } else {
      const sig = hint.parts
        .map((p) => {
          if (p.kind !== "paramActive") return p.text;
          // Signature parts use brackets for optional params; avoid double-bracketing
          // when the active param is already optional.
          if (p.text.startsWith("[") && p.text.endsWith("]")) return p.text;
          return `[${p.text}]`;
        })
        .join("");
      const summary = hint.signature.summary?.trim?.() ?? "";
      this.#hintEl.textContent = summary ? `${sig} — ${summary}` : sig;
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
