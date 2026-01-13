import { CellStructuralConflictMonitor, FormulaConflictMonitor } from "../../../../packages/collab/conflicts/index.js";

/**
 * Yjs transaction origin used by `CollabVersioning.restoreVersion`.
 *
 * This origin should be *ignored* by formula/value conflict monitoring (to avoid
 * spurious conflict UI *and* avoid polluting local edit tracking with bulk
 * snapshot rewrites).
 *
 * It must also not be treated as local by the structural conflict monitor
 * (otherwise it would log thousands of structural ops into `cellStructuralOps`).
 */
export const VERSIONING_RESTORE_ORIGIN = "versioning-restore";

/**
 * Yjs transaction origin used when applying branch checkout/merge state back into
 * the live workbook.
 */
export const BRANCHING_APPLY_ORIGIN = "branching-apply";

/**
 * Create a FormulaConflictMonitor configured for desktop collaboration wiring.
 *
 * Key behaviors:
 * - Treat `binderOrigin` (DocumentController-driven edits) as local.
 * - Treat `sessionOrigin` (conflict resolution writes, etc) as local.
 * - Ignore bulk "time travel" operations (`VERSIONING_RESTORE_ORIGIN`, `BRANCHING_APPLY_ORIGIN`)
 *   so they do not surface conflict UI or pollute local-edit tracking.
 *
 * @param {object} opts
 * @param {import("yjs").Doc} opts.doc
 * @param {import("yjs").Map<any>} [opts.cells]
 * @param {string} opts.localUserId
 * @param {any} opts.sessionOrigin
 * @param {any} opts.binderOrigin
 * @param {Set<any>} [opts.undoLocalOrigins]
 * @param {(conflict: import("../../../../packages/collab/conflicts/index.js").FormulaConflict) => void} opts.onConflict
 * @param {(ref: { sheetId: string, row: number, col: number }) => any} [opts.getCellValue]
 * @param {"formula" | "formula+value"} [opts.mode]
 */
export function createDesktopFormulaConflictMonitor(opts) {
  const localOrigins = new Set();
  localOrigins.add(opts.sessionOrigin);
  localOrigins.add(opts.binderOrigin);

  if (opts.undoLocalOrigins) {
    for (const origin of opts.undoLocalOrigins) {
      localOrigins.add(origin);
    }
  }

  const ignoredOrigins = new Set([VERSIONING_RESTORE_ORIGIN, BRANCHING_APPLY_ORIGIN]);

  return new FormulaConflictMonitor({
    doc: opts.doc,
    cells: opts.cells,
    localUserId: opts.localUserId,
    // Use the session origin for monitor writes so conflict resolutions propagate
    // through the Yjs->DocumentController binder.
    origin: opts.sessionOrigin,
    localOrigins,
    ignoredOrigins,
    onConflict: opts.onConflict,
    getCellValue: opts.getCellValue,
    mode: opts.mode,
  });
}

/**
 * Create a CellStructuralConflictMonitor configured for desktop collaboration wiring.
 *
 * Key behavior:
 * - Only treat DocumentController-driven edits (binderOrigin, plus optional undo origins)
 *   as local so they are logged into `cellStructuralOps`.
 * - Do NOT treat `sessionOrigin` as local (it is used for programmatic/bulk writes like
 *   conflict resolutions and should not be logged).
 * - Do NOT treat `VERSIONING_RESTORE_ORIGIN` as local (version restore is a bulk write).
 * - Do NOT treat `BRANCHING_APPLY_ORIGIN` as local (branch checkout/merge apply is a bulk write).
 *
 * @param {object} opts
 * @param {import("yjs").Doc} opts.doc
 * @param {import("yjs").Map<any>} [opts.cells]
 * @param {string} opts.localUserId
 * @param {any} opts.sessionOrigin
 * @param {any} opts.binderOrigin
 * @param {Set<any>} [opts.undoLocalOrigins]
 * @param {(conflict: import("../../../../packages/collab/conflicts/index.js").CellStructuralConflict) => void} opts.onConflict
 * @param {number} [opts.maxOpRecordsPerUser]
 * @param {number | null} [opts.maxOpRecordAgeMs]
 */
export function createDesktopCellStructuralConflictMonitor(opts) {
  const localOrigins = new Set();
  localOrigins.add(opts.binderOrigin);

  if (opts.undoLocalOrigins) {
    for (const origin of opts.undoLocalOrigins) {
      localOrigins.add(origin);
    }
  }

  // Critical exclusions: bulk operations should never be treated as local for op logging.
  localOrigins.delete(opts.sessionOrigin);
  localOrigins.delete(VERSIONING_RESTORE_ORIGIN);
  localOrigins.delete(BRANCHING_APPLY_ORIGIN);

  return new CellStructuralConflictMonitor({
    doc: opts.doc,
    cells: opts.cells,
    localUserId: opts.localUserId,
    // Use session origin for conflict resolution writes so they propagate through the binder.
    origin: opts.sessionOrigin,
    localOrigins,
    ignoredOrigins: new Set([VERSIONING_RESTORE_ORIGIN, BRANCHING_APPLY_ORIGIN]),
    onConflict: opts.onConflict,
    maxOpRecordsPerUser: opts.maxOpRecordsPerUser,
    maxOpRecordAgeMs: opts.maxOpRecordAgeMs,
  });
}
