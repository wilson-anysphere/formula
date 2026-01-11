import { FormulaBarModel } from "./FormulaBarModel.js";
import { parseA1Range, type RangeAddress } from "../spreadsheet/a1.js";

export interface FormulaBarViewCallbacks {
  onBeginEdit?: (activeCellAddress: string) => void;
  onCommit: (text: string) => void;
  onCancel?: () => void;
  onGoTo?: (reference: string) => void;
  onHoverRange?: (range: RangeAddress | null) => void;
}

export class FormulaBarView {
  readonly model = new FormulaBarModel();

  readonly root: HTMLElement;
  readonly textarea: HTMLTextAreaElement;

  #addressEl: HTMLInputElement;
  #highlightEl: HTMLElement;
  #hintEl: HTMLElement;
  #errorButton: HTMLButtonElement;
  #errorPanel: HTMLElement;
  #hoverOverride: RangeAddress | null = null;
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

    const fx = document.createElement("div");
    fx.className = "formula-bar-fx";
    fx.textContent = "fx";

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
    errorButton.dataset.testid = "formula-error-button";
    errorButton.style.display = "none";

    const errorPanel = document.createElement("div");
    errorPanel.className = "formula-bar-error-panel";
    errorPanel.dataset.testid = "formula-error-panel";
    errorPanel.style.display = "none";

    row.appendChild(address);
    row.appendChild(fx);
    row.appendChild(editor);
    row.appendChild(errorButton);

    const hint = document.createElement("div");
    hint.className = "formula-bar-hint";
    hint.dataset.testid = "formula-hint";

    root.appendChild(row);
    root.appendChild(hint);
    root.appendChild(errorPanel);

    this.textarea = textarea;
    this.#addressEl = address;
    this.#highlightEl = highlight;
    this.#hintEl = hint;
    this.#errorButton = errorButton;
    this.#errorPanel = errorPanel;

    address.addEventListener("focus", () => {
      address.select();
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
    textarea.addEventListener("click", () => this.#onInputOrSelection());
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
      const isOpen = this.#errorPanel.style.display !== "none";
      this.#errorPanel.style.display = isOpen ? "none" : "block";
    });

    // Initial render.
    this.model.setActiveCell({ address: "A1", input: "", value: "" });
    this.#render({ preserveTextareaValue: false });
  }

  setAiSuggestion(suggestion: string | null): void {
    this.model.setAiSuggestion(suggestion);
    this.#render({ preserveTextareaValue: true });
  }

  focus(opts: { cursor?: "end" | "all" } = {}): void {
    this.textarea.style.display = "block";
    this.textarea.focus();
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
    this.#render({ preserveTextareaValue: false });
  }

  isEditing(): boolean {
    return this.model.isEditing;
  }

  isFormulaEditing(): boolean {
    return this.model.isEditing && this.model.draft.trim().startsWith("=");
  }

  beginRangeSelection(range: RangeAddress): void {
    this.model.beginEdit();
    this.model.beginRangeSelection(range);
    this.#render({ preserveTextareaValue: false });
    this.#setTextareaSelectionFromModel();
    this.#emitHover();
  }

  updateRangeSelection(range: RangeAddress): void {
    this.model.updateRangeSelection(range);
    this.#render({ preserveTextareaValue: false });
    this.#setTextareaSelectionFromModel();
    this.#emitHover();
  }

  endRangeSelection(): void {
    this.model.endRangeSelection();
  }

  #beginEditFromFocus(): void {
    if (this.model.isEditing) return;
    this.model.beginEdit();
    this.#callbacks.onBeginEdit?.(this.model.activeCell.address);
    this.#render({ preserveTextareaValue: true });
    this.#emitHover();
  }

  #onInputOrSelection(): void {
    if (!this.model.isEditing) return;

    const start = this.textarea.selectionStart ?? this.textarea.value.length;
    const end = this.textarea.selectionEnd ?? this.textarea.value.length;
    this.model.updateDraft(this.textarea.value, start, end);
    this.#render({ preserveTextareaValue: true });
    this.#emitHover();
  }

  #onKeyDown(e: KeyboardEvent): void {
    if (!this.model.isEditing) return;

    if (e.key === "Tab") {
      const accepted = this.model.acceptAiSuggestion();
      if (accepted) {
        e.preventDefault();
        this.#render({ preserveTextareaValue: false });
        this.#setTextareaSelectionFromModel();
        this.#emitHover();
        return;
      }
    }

    if (e.key === "Escape") {
      e.preventDefault();
      this.textarea.blur();
      this.model.cancel();
      this.#hoverOverride = null;
      this.#render({ preserveTextareaValue: false });
      this.#callbacks.onCancel?.();
      this.#emitHover();
      return;
    }

    // Excel behavior: Enter commits, Alt+Enter inserts newline.
    if (e.key === "Enter" && !e.altKey) {
      e.preventDefault();
      this.textarea.blur();
      const committed = this.model.commit();
      this.#hoverOverride = null;
      this.#render({ preserveTextareaValue: false });
      this.#callbacks.onCommit(committed);
      this.#emitHover();
      return;
    }
  }

  #render(opts: { preserveTextareaValue: boolean }): void {
    if (document.activeElement !== this.#addressEl) {
      this.#addressEl.value = this.model.activeCell.address;
    }

    if (!opts.preserveTextareaValue) {
      this.textarea.value = this.model.draft;
    }

    const cursor = this.model.cursorStart;
    const ghost = this.model.isEditing ? this.model.aiGhostText() : "";
    let ghostInserted = false;
    let highlightHtml = "";

    for (const span of this.model.highlightedSpans()) {
      if (!ghostInserted && ghost && cursor <= span.start) {
        highlightHtml += `<span class="formula-bar-ghost">${escapeHtml(ghost)}</span>`;
        ghostInserted = true;
      }

      if (!ghostInserted && ghost && cursor > span.start && cursor < span.end) {
        const split = cursor - span.start;
        const before = span.text.slice(0, split);
        const after = span.text.slice(split);
        if (before) {
          highlightHtml += `<span data-kind="${span.kind}">${escapeHtml(before)}</span>`;
        }
        highlightHtml += `<span class="formula-bar-ghost">${escapeHtml(ghost)}</span>`;
        ghostInserted = true;
        if (after) {
          highlightHtml += `<span data-kind="${span.kind}">${escapeHtml(after)}</span>`;
        }
        continue;
      }

      highlightHtml += `<span data-kind="${span.kind}">${escapeHtml(span.text)}</span>`;
    }

    if (!ghostInserted && ghost) {
      highlightHtml += `<span class="formula-bar-ghost">${escapeHtml(ghost)}</span>`;
    }
    this.#highlightEl.innerHTML = highlightHtml;

    // When not editing, hide the textarea and allow hover interactions directly on the highlighted text.
    this.textarea.style.display = this.model.isEditing ? "block" : "none";

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
      this.#errorButton.style.display = "none";
      this.#errorPanel.style.display = "none";
      this.#errorPanel.textContent = "";
    } else {
      this.#errorButton.style.display = "inline-flex";
      this.#errorPanel.innerHTML = `
        <div class="formula-bar-error-title">${explanation.code}: ${explanation.title}</div>
        <div class="formula-bar-error-desc">${explanation.description}</div>
        <ul class="formula-bar-error-suggestions">${explanation.suggestions.map((s) => `<li>${s}</li>`).join("")}</ul>
      `;
    }

    this.#syncScroll();
    this.#adjustHeight();
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

  #emitHover(): void {
    const range = this.#hoverOverride ?? this.model.hoveredReference();
    this.#callbacks.onHoverRange?.(range);
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
    this.#hoverOverride = text ? parseA1Range(text) : null;
    this.#callbacks.onHoverRange?.(this.#hoverOverride);
  }

  #clearHoverOverride(): void {
    if (this.#hoverOverride === null) return;
    this.#hoverOverride = null;
    this.#emitHover();
  }
}

function escapeHtml(text: string): string {
  return text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}
