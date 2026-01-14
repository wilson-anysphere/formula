import * as Y from "yjs";
import { isYAbstractType, patchForeignAbstractTypeConstructor, patchForeignItemConstructor } from "@formula/collab-yjs-utils";

function patchForeignItemsInType(type) {
  if (!type || typeof type !== "object") return;
  patchForeignAbstractTypeConstructor(type);

  const map = type._map;
  if (map instanceof Map) {
    for (const value of map.values()) {
      patchForeignItemConstructor(value);
      // Also patch left chain because map entries are linked lists.
      let cur = value;
      while (cur) {
        patchForeignItemConstructor(cur);
        cur = cur.left;
      }
    }
  }

  let cur = type._start;
  while (cur) {
    patchForeignItemConstructor(cur);
    cur = cur.right;
  }
}

function patchForeignItemsInScope(scope) {
  const types = Array.isArray(scope) ? scope : [scope];
  for (const type of types) {
    patchForeignItemsInType(type);
  }
}

function patchForeignItemsInTransaction(transaction) {
  if (!transaction || typeof transaction !== "object") return;

  // Patch foreign item constructors referenced by deleted structs. This is the
  // critical path for correctness: Yjs UndoManager uses `instanceof Item` checks
  // when tracking deletions, and will otherwise fail to undo overwrites of
  // foreign items (e.g. when remote updates were applied by a different Yjs
  // module instance).
  try {
    Y.iterateDeletedStructs(transaction, transaction.deleteSet, patchForeignItemConstructor);
  } catch {
    // ignore
  }

  // Also patch constructors for newly-inserted structs so undo can later delete
  // them even if they were created by foreign type methods.
  //
  // (We intentionally only do this for transactions that we expect to be tracked
  // by UndoManager; see the caller in `createCollabUndoService`.)
  try {
    const insertions = Y.createDeleteSet();
    const beforeState = transaction.beforeState;
    transaction.afterState?.forEach?.((endClock, client) => {
      const startClock = beforeState?.get?.(client) ?? 0;
      const len = endClock - startClock;
      if (len <= 0) return;
      const arr = insertions.clients.get(client);
      const entry = { clock: startClock, len };
      if (arr) arr.push(entry);
      else insertions.clients.set(client, [entry]);
    });
    Y.iterateDeletedStructs(transaction, insertions, patchForeignItemConstructor);
  } catch {
    // ignore
  }
}

/**
 * Origin token used when applying remote updates in tests / providers.
 *
 * In Yjs, UndoManager can track changes by origin. By ensuring remote updates
 * use a non-local origin, we guarantee undo/redo only affects local changes.
 */
export const REMOTE_ORIGIN = { type: "remote" };

function undoOnce(undoManager) {
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
}

function redoOnce(undoManager) {
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
}

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

  patchForeignItemsInScope(scope);

  // Patch foreign `Item` constructors opportunistically for transactions tracked
  // by this UndoManager.
  //
  // Important: register this handler *before* constructing the UndoManager so it
  // runs before Yjs' own `afterTransaction` handler. This ensures that when a
  // local transaction overwrites a foreign item (e.g. remote updates were applied
  // via a different Yjs module instance), the UndoManager still recognizes the
  // deleted foreign Item structs and records an undoable change.
  //
  // See regression: `yjs-undo-service.foreign-items-late.test.js`.
  /** @type {any} */
  let undoManager = null;
  const patchAfterTransaction = (transaction) => {
    const txnOrigin = transaction?.origin;
    if (txnOrigin !== origin && txnOrigin !== undoManager) return;
    patchForeignItemsInTransaction(transaction);
  };
  doc.on("afterTransaction", patchAfterTransaction);

  undoManager = new Y.UndoManager(scope, {
    captureTimeout: captureTimeoutMs,
    trackedOrigins: new Set([origin])
  });

  // Also patch any types added later via `undoManager.addToScope(...)` so we
  // don't regress in desktop-style flows that lazily extend the undo scope.
  const addToScope = undoManager.addToScope.bind(undoManager);
  undoManager.addToScope = (ytypes) => {
    patchForeignItemsInScope(ytypes);
    addToScope(ytypes);
  };

  // Ensure our patching handler is removed if callers explicitly destroy the
  // UndoManager instance.
  const destroyUndoManager = undoManager.destroy.bind(undoManager);
  undoManager.destroy = () => {
    doc.off("afterTransaction", patchAfterTransaction);
    destroyUndoManager();
  };

  const localOrigins = new Set([origin, undoManager]);

  // Yjs UndoManager will skip stack items that it can't apply (e.g. a Y.Map key
  // was overwritten by a remote collaborator) and continue undoing older items.
  //
  // Formula's collaborative undo semantics are stricter: a single undo action
  // should never skip past an un-undoable local change and revert earlier edits
  // (which can indirectly undo other users' work, e.g. undoing a sheet insert
  // after a remote rename overwrote our local rename).
  //
  // To match that behavior, we constrain each `undo()` / `redo()` call to at most
  // one stack item by temporarily isolating the top entry.
  return {
    mode: "collab",
    undoManager,
    origin,
    localOrigins,
    canUndo: () => undoManager.canUndo(),
    canRedo: () => undoManager.canRedo(),
    undo: () => undoOnce(undoManager),
    redo: () => redoOnce(undoManager),
    stopCapturing: () => undoManager.stopCapturing(),
    transact: (fn) => {
      doc.transact(fn, origin);
    }
  };
}
