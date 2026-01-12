import type { CellCoord } from "../selection/types";
import type { Rect } from "../selection/renderer";
import { toggleA1AbsoluteAtCursor } from "@formula/spreadsheet-frontend";

export type EditorCommitReason = "enter" | "tab" | "command";

export interface EditorCommit {
  cell: CellCoord;
  value: string;
  reason: EditorCommitReason;
  /**
   * Shift modifier for enter/tab (Shift+Enter moves up, Shift+Tab moves left).
   */
  shift: boolean;
}

export interface CellEditorOverlayCallbacks {
  onCommit: (commit: EditorCommit) => void;
  onCancel: (cell: CellCoord) => void;
}

export class CellEditorOverlay {
  readonly element: HTMLTextAreaElement;
  private editingCell: CellCoord | null = null;
  private minWidth = 0;
  private minHeight = 0;

  constructor(
    private container: HTMLElement,
    private callbacks: CellEditorOverlayCallbacks
  ) {
    this.element = document.createElement("textarea");
    this.element.className = "cell-editor";
    this.element.spellcheck = false;
    this.element.autocapitalize = "off";
    this.element.autocomplete = "off";
    this.element.wrap = "off";
    this.element.style.display = "none";
    // Ensure the editor sits above canvas overlay layers (selection, charts, etc),
    // especially in shared-grid mode where canvases are explicitly z-indexed.
    this.element.style.zIndex = "10";

    this.element.addEventListener("input", () => this.adjustSize());
    this.element.addEventListener("keydown", (e) => this.onKeyDown(e));

    this.container.appendChild(this.element);
  }

  isOpen(): boolean {
    return this.editingCell !== null;
  }

  commit(reason: EditorCommitReason = "enter", shift = false): void {
    if (!this.editingCell) return;
    const cell = this.editingCell;
    const value = this.element.value;
    this.close();
    this.callbacks.onCommit({ cell, value, reason, shift });
  }

  cancel(): void {
    if (!this.editingCell) return;
    const cell = this.editingCell;
    this.close();
    this.callbacks.onCancel(cell);
  }

  open(cell: CellCoord, bounds: Rect, initialValue: string, opts: { cursor?: "end" | "all" } = {}): void {
    this.editingCell = cell;
    this.minWidth = bounds.width;
    this.minHeight = bounds.height;

    this.element.value = initialValue;
    this.element.style.display = "block";
    this.element.style.left = `${bounds.x}px`;
    this.element.style.top = `${bounds.y}px`;
    this.element.style.width = `${bounds.width}px`;
    this.element.style.height = `${bounds.height}px`;

    // Focus before setting selection for consistent behavior.
    this.element.focus();

    if (opts.cursor === "all") {
      this.element.setSelectionRange(0, this.element.value.length);
    } else {
      // Excel's F2 semantics: put caret at end of cell contents.
      const end = this.element.value.length;
      this.element.setSelectionRange(end, end);
    }

    this.adjustSize();
  }

  close(): void {
    this.editingCell = null;
    this.element.style.display = "none";
    this.element.value = "";
  }

  reposition(bounds: Rect): void {
    if (!this.isOpen()) return;
    this.minWidth = bounds.width;
    this.minHeight = bounds.height;
    this.element.style.left = `${bounds.x}px`;
    this.element.style.top = `${bounds.y}px`;
    this.element.style.width = `${bounds.width}px`;
    this.element.style.height = `${bounds.height}px`;
    this.adjustSize();
  }

  private adjustSize(): void {
    if (!this.isOpen()) return;

    // Keep the editor at least as large as the cell, but allow it to expand
    // horizontally/vertically as the user types.
    //
    // Note: scrollWidth/scrollHeight are measured using the current size, so
    // start from the minimum size first.
    this.element.style.width = `${this.minWidth}px`;
    this.element.style.height = `${this.minHeight}px`;

    const width = Math.max(this.minWidth, this.element.scrollWidth + 2);
    const height = Math.max(this.minHeight, this.element.scrollHeight + 2);

    this.element.style.width = `${width}px`;
    this.element.style.height = `${height}px`;
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (!this.editingCell) return;

    if (e.key === "F4" && !e.altKey && !e.ctrlKey && !e.metaKey && this.element.value.trim().startsWith("=")) {
      e.preventDefault();
      e.stopPropagation();

      const prevText = this.element.value;
      const cursorStart = this.element.selectionStart ?? prevText.length;
      const cursorEnd = this.element.selectionEnd ?? prevText.length;

      const toggled = toggleA1AbsoluteAtCursor(prevText, cursorStart, cursorEnd);
      if (!toggled) return;

      this.element.value = toggled.text;
      this.element.setSelectionRange(toggled.cursorStart, toggled.cursorEnd);
      this.adjustSize();
      return;
    }

    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      this.cancel();
      return;
    }

    if (e.key === "Enter" && !e.altKey) {
      e.preventDefault();
      e.stopPropagation();
      this.commit("enter", e.shiftKey);
      return;
    }

    if (e.key === "Tab") {
      e.preventDefault();
      e.stopPropagation();
      this.commit("tab", e.shiftKey);
    }
  }
}
