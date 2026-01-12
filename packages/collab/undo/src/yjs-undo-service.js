import * as Y from "yjs";

const patchedItemConstructors = new WeakSet();

function isYjsItemStruct(value) {
  if (!value || typeof value !== "object") return false;
  const maybe = value;
  // Yjs internal `Item` structs have these fields (see yjs/src/structs/Item).
  if (!("id" in maybe)) return false;
  if (typeof maybe.length !== "number") return false;
  if (!("content" in maybe)) return false;
  if (!("parent" in maybe)) return false;
  if (!("parentSub" in maybe)) return false;
  if (typeof maybe.content?.getContent !== "function") return false;
  return true;
}

function patchForeignItemConstructor(item) {
  if (!item || typeof item !== "object") return;
  if (!isYjsItemStruct(item)) return;
  if (item instanceof Y.Item) return;
  const ctor = item.constructor;
  if (!ctor || ctor === Y.Item) return;
  if (patchedItemConstructors.has(ctor)) return;
  patchedItemConstructors.add(ctor);

  // When Yjs is loaded more than once (e.g. ESM + CJS in Node), documents can
  // contain Item instances created by a different module instance. Yjs'
  // UndoManager uses `instanceof Item` checks, so it will refuse to undo
  // transactions that touch those foreign items.
  //
  // Patch the foreign constructor prototype chain so foreign Item instances pass
  // `instanceof Y.Item` checks in this module.
  try {
    Object.setPrototypeOf(ctor.prototype, Y.Item.prototype);
    ctor.prototype.constructor = Y.Item;
  } catch {
    // Best-effort: if we can't patch (frozen prototypes, etc), undo will behave
    // like upstream Yjs in mixed-module environments.
  }
}

function patchForeignItemsInType(type) {
  if (!type || typeof type !== "object") return;

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

  const undoManager = new Y.UndoManager(scope, {
    captureTimeout: captureTimeoutMs,
    trackedOrigins: new Set([origin])
  });

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
