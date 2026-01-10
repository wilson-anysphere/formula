/**
 * A small batching undo stack intended for single-user mode.
 *
 * In collaborative mode we defer to Yjs' UndoManager so undo/redo only affects
 * local-origin transactions.
 */
export class LocalUndoStack {
  /**
   * @param {object} [opts]
   * @param {number} [opts.captureTimeoutMs]
   */
  constructor(opts = {}) {
    const { captureTimeoutMs = 750 } = opts;
    this._captureTimeoutMs = captureTimeoutMs;

    /** @type {Array<{ actions: Array<UndoableAction> }>} */
    this._undoStack = [];
    /** @type {Array<{ actions: Array<UndoableAction> }>} */
    this._redoStack = [];

    this._lastPushAt = 0;
    this._capturing = false;
  }

  /** @returns {boolean} */
  canUndo() {
    return this._undoStack.length > 0;
  }

  /** @returns {boolean} */
  canRedo() {
    return this._redoStack.length > 0;
  }

  stopCapturing() {
    this._capturing = false;
    this._lastPushAt = 0;
  }

  /**
   * @param {UndoableAction} action
   */
  push(action) {
    const now = Date.now();
    const shouldBatch = this._capturing && now - this._lastPushAt < this._captureTimeoutMs;

    if (shouldBatch && this._undoStack.length > 0) {
      this._undoStack[this._undoStack.length - 1].actions.push(action);
    } else {
      this._undoStack.push({ actions: [action] });
    }

    this._redoStack.length = 0;
    this._capturing = true;
    this._lastPushAt = now;
  }

  /**
   * Convenience API used by UndoService - executes the action and records it.
   * @param {UndoableAction} action
   */
  perform(action) {
    action.redo();
    this.push(action);
  }

  undo() {
    const batch = this._undoStack.pop();
    if (!batch) return;

    // Undo in reverse order.
    for (let i = batch.actions.length - 1; i >= 0; i -= 1) {
      batch.actions[i].undo();
    }
    this._redoStack.push(batch);
    this.stopCapturing();
  }

  redo() {
    const batch = this._redoStack.pop();
    if (!batch) return;

    for (const action of batch.actions) {
      action.redo();
    }
    this._undoStack.push(batch);
    this.stopCapturing();
  }
}

/**
 * @typedef {object} UndoableAction
 * @property {() => void} undo
 * @property {() => void} redo
 */

