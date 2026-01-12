export interface SecondaryGridShortcutApp {
  copy: () => void;
  cut: () => void;
  paste: () => void;
  clearSelection: () => void;
  openCommentsPanel: () => void;
  fillDown: () => void;
  fillRight: () => void;
  insertDate: () => void;
  insertTime: () => void;
  autoSum: () => void;
}

export interface SecondaryGridShortcutOptions {
  app: SecondaryGridShortcutApp;
  /**
   * Returns true when *any* spreadsheet editing surface is active (primary cell editor,
   * formula bar, inline edit, or the split-view secondary editor).
   */
  isSpreadsheetEditing: () => boolean;
  /**
   * Returns true when the event target is a text-input surface (input/textarea/contenteditable).
   */
  isTextInputTarget: (target: EventTarget | null) => boolean;
  /**
   * Optional hook to cancel an in-progress secondary-pane fill-handle drag.
   */
  cancelFillHandleDrag?: () => boolean;
}

/**
 * Keydown handler for the split-view secondary grid.
 *
 * SpreadsheetApp owns most Excel-style shortcuts, but its grid-level keydown handler only runs
 * when the primary grid element (#grid) has focus. In split-view mode, the secondary pane needs
 * to map the same keystrokes back into SpreadsheetApp APIs so Excel shortcuts behave consistently.
 */
export function handleSecondaryGridKeyDown(event: KeyboardEvent, opts: SecondaryGridShortcutOptions): boolean {
  if (event.defaultPrevented) return false;
  if (opts.isTextInputTarget(event.target)) return false;
  if (opts.isSpreadsheetEditing()) return false;

  // Excel semantics: Shift+F2 opens the comments panel.
  if (event.key === "F2" && event.shiftKey) {
    event.preventDefault();
    opts.app.openCommentsPanel();
    return true;
  }

  // Cancel an in-progress fill-handle drag (matches primary-grid Escape semantics).
  if (event.key === "Escape" && opts.cancelFillHandleDrag?.()) {
    event.preventDefault();
    return true;
  }

  const primary = event.ctrlKey || event.metaKey;
  const keyLower = (event.key ?? "").toLowerCase();

  // Clipboard shortcuts (Excel-like).
  if (primary && !event.altKey && !event.shiftKey) {
    if (keyLower === "c") {
      event.preventDefault();
      opts.app.copy();
      return true;
    }
    if (keyLower === "x") {
      event.preventDefault();
      opts.app.cut();
      return true;
    }
    if (keyLower === "v") {
      event.preventDefault();
      opts.app.paste();
      return true;
    }

    // Excel fill shortcuts:
    // - Ctrl/Cmd+D: Fill Down
    // - Ctrl/Cmd+R: Fill Right
    if (keyLower === "d") {
      event.preventDefault();
      opts.app.fillDown();
      return true;
    }
    if (keyLower === "r") {
      event.preventDefault();
      opts.app.fillRight();
      return true;
    }
  }

  // Excel-style date/time insertion:
  // - Ctrl/Cmd+; inserts the current date
  // - Ctrl/Cmd+Shift+; inserts the current time
  if (primary && !event.altKey && event.code === "Semicolon") {
    event.preventDefault();
    if (event.shiftKey) opts.app.insertTime();
    else opts.app.insertDate();
    return true;
  }

  // Excel AutoSum: Alt+=
  if (event.altKey && event.code === "Equal" && !event.ctrlKey && !event.metaKey) {
    event.preventDefault();
    opts.app.autoSum();
    return true;
  }

  // Delete/backspace clears the selection (unless editing).
  if (event.key === "Delete" || event.key === "Backspace") {
    event.preventDefault();
    opts.app.clearSelection();
    return true;
  }

  return false;
}

