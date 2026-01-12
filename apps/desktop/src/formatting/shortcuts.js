/**
 * @typedef {{
 *   key?: string,
 *   code?: string,
 *   ctrlKey?: boolean,
 *   metaKey?: boolean,
 *   shiftKey?: boolean,
 *   preventDefault?: () => void
 * }} KeyboardEventLike
 */

/**
 * Ctrl+1 / Cmd+1 (Excel's Format Cells).
 *
 * @param {KeyboardEventLike} event
 */
export function isFormatCellsKeyboardEvent(event) {
  const key = (event.key ?? "").toLowerCase();
  const code = event.code ?? "";
  const mod = Boolean(event.metaKey || event.ctrlKey);
  if (!mod || event.shiftKey) return false;

  // Prefer `event.key` (layout-aware), but fall back to `event.code` so this shortcut
  // remains usable on layouts where the digit requires Shift (e.g. AZERTY `Digit1` -> "&").
  return key === "1" || code === "Digit1";
}

/**
 * @param {{ addEventListener: (type: string, listener: (event: any) => void) => void, removeEventListener: (type: string, listener: (event: any) => void) => void }} target
 * @param {{ openFormatCells: () => void }} handlers
 */
export function installFormattingShortcuts(target, handlers) {
  /** @param {any} event */
  function onKeyDown(event) {
    if (isFormatCellsKeyboardEvent(event)) {
      event.preventDefault?.();
      handlers.openFormatCells();
    }
  }

  target.addEventListener("keydown", onKeyDown);
  return () => target.removeEventListener("keydown", onKeyDown);
}
