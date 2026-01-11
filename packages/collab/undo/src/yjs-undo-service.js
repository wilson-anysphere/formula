import * as Y from "yjs";

/**
 * Origin token used when applying remote updates in tests / providers.
 *
 * In Yjs, UndoManager can track changes by origin. By ensuring remote updates
 * use a non-local origin, we guarantee undo/redo only affects local changes.
 */
export const REMOTE_ORIGIN = { type: "remote" };

/**
 * @typedef {object} CollabUndoService
 * @property {"collab"} mode
 * @property {Y.UndoManager} undoManager
 * @property {object} origin
 * @property {Set<any>} localOrigins
 * @property {() => boolean} canUndo
 * @property {() => boolean} canRedo
 * @property {() => void} undo
 * @property {() => void} redo
 * @property {() => void} stopCapturing
 * @property {(fn: () => void) => void} transact
 */

/**
 * @param {object} opts
 * @param {Y.Doc} opts.doc
 * @param {Y.AbstractType<any>|Array<Y.AbstractType<any>>} opts.scope
 * @param {number} [opts.captureTimeoutMs]
 * @param {object} [opts.origin] Optional stable origin token for local changes.
 * @returns {CollabUndoService}
 */
export function createCollabUndoService(opts) {
  const { doc, scope, captureTimeoutMs = 750 } = opts;
  const origin = opts.origin ?? { type: "local" };

  const undoManager = new Y.UndoManager(scope, {
    captureTimeout: captureTimeoutMs,
    trackedOrigins: new Set([origin])
  });

  const localOrigins = new Set([origin, undoManager]);

  const undoOnce = () => {
    if (!undoManager.canUndo()) return;
    const saved = undoManager.undoStack;
    const item = saved.pop();
    if (!item) return;
    undoManager.undoStack = [item];
    try {
      undoManager.undo();
    } finally {
      undoManager.undoStack = saved;
    }
  };

  const redoOnce = () => {
    if (!undoManager.canRedo()) return;
    const saved = undoManager.redoStack;
    const item = saved.pop();
    if (!item) return;
    undoManager.redoStack = [item];
    try {
      undoManager.redo();
    } finally {
      undoManager.redoStack = saved;
    }
  };

  return {
    mode: "collab",
    undoManager,
    origin,
    localOrigins,
    canUndo: () => undoManager.canUndo(),
    canRedo: () => undoManager.canRedo(),
    undo: () => undoOnce(),
    redo: () => redoOnce(),
    stopCapturing: () => undoManager.stopCapturing(),
    transact: (fn) => {
      doc.transact(fn, origin);
    }
  };
}
