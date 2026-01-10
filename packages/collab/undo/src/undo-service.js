import { LocalUndoStack } from "./local-undo-stack.js";
import { createCollabUndoService } from "./yjs-undo-service.js";

/**
 * @typedef {object} UndoService
 * @property {"single"|"collab"} mode
 * @property {() => boolean} canUndo
 * @property {() => boolean} canRedo
 * @property {() => void} undo
 * @property {() => void} redo
 * @property {() => void} stopCapturing
 * @property {(change: { redo: () => void, undo?: () => void }) => void} perform
 * @property {(fn: () => void) => void} [transact] In collab mode, runs fn in a local-origin Yjs transaction.
 * @property {Set<any>} [localOrigins] In collab mode, origins considered "local" (useful for conflict detection).
 */

/**
 * Creates an undo service that matches the application's undo semantics:
 * - In single-user mode we maintain our own undo stack.
 * - In collaborative mode we use Yjs' UndoManager so undo/redo only affects
 *   local-origin changes (remote users' edits are never reverted).
 *
 * @param {object} opts
 * @param {"single"|"collab"} opts.mode
 * @param {import("yjs").Doc} [opts.doc] Required for collab mode.
 * @param {import("yjs").AbstractType<any>|Array<import("yjs").AbstractType<any>>} [opts.scope] Required for collab mode.
 * @param {number} [opts.captureTimeoutMs]
 * @param {object} [opts.origin]
 * @returns {UndoService}
 */
export function createUndoService(opts) {
  const { mode, captureTimeoutMs = 750 } = opts;

  if (mode === "single") {
    const stack = new LocalUndoStack({ captureTimeoutMs });

    return {
      mode,
      canUndo: () => stack.canUndo(),
      canRedo: () => stack.canRedo(),
      undo: () => stack.undo(),
      redo: () => stack.redo(),
      stopCapturing: () => stack.stopCapturing(),
      perform: (change) => {
        if (!change.undo) {
          throw new Error("Single-user undo requires an undo() function for each change.");
        }

        stack.perform({
          redo: change.redo,
          undo: change.undo
        });
      }
    };
  }

  if (mode === "collab") {
    if (!opts.doc || !opts.scope) {
      throw new Error("Collaborative undo requires { doc, scope }.");
    }

    const collab = createCollabUndoService({
      doc: opts.doc,
      scope: opts.scope,
      captureTimeoutMs,
      origin: opts.origin
    });

    return {
      mode,
      localOrigins: collab.localOrigins,
      transact: collab.transact,
      canUndo: collab.canUndo,
      canRedo: collab.canRedo,
      undo: collab.undo,
      redo: collab.redo,
      stopCapturing: collab.stopCapturing,
      perform: (change) => {
        collab.transact(change.redo);
      }
    };
  }

  throw new Error(`Unknown undo mode: ${mode}`);
}

