/**
 * @typedef {{ key?: string, ctrlKey?: boolean, metaKey?: boolean, shiftKey?: boolean, preventDefault?: () => void }} UndoRedoKeyboardEvent
 */

/**
 * @param {UndoRedoKeyboardEvent} event
 */
export function isUndoKeyboardEvent(event) {
  const key = (event.key ?? "").toLowerCase();
  const mod = Boolean(event.metaKey || event.ctrlKey);
  return mod && !event.shiftKey && key === "z";
}

/**
 * @param {UndoRedoKeyboardEvent} event
 */
export function isRedoKeyboardEvent(event) {
  const key = (event.key ?? "").toLowerCase();
  const mod = Boolean(event.metaKey || event.ctrlKey);
  return mod && ((event.shiftKey && key === "z") || (!event.shiftKey && key === "y"));
}

/**
 * Install standard undo/redo shortcuts onto a DOM-like target.
 *
 * @param {{ addEventListener: (type: string, listener: (event: any) => void) => void, removeEventListener: (type: string, listener: (event: any) => void) => void }} target
 * @param {{ undo: () => boolean, redo: () => boolean }} controller
 * @returns {() => void} cleanup
 */
export function installUndoRedoShortcuts(target, controller) {
  /** @param {any} event */
  function onKeyDown(event) {
    if (isUndoKeyboardEvent(event)) {
      event.preventDefault?.();
      controller.undo();
      return;
    }
    if (isRedoKeyboardEvent(event)) {
      event.preventDefault?.();
      controller.redo();
    }
  }

  target.addEventListener("keydown", onKeyDown);
  return () => target.removeEventListener("keydown", onKeyDown);
}

