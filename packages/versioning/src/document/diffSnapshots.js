import { semanticDiff } from "../diff/semanticDiff.js";
import { sheetStateFromDocumentSnapshot } from "./sheetState.js";

/**
 * Compute a semantic diff between two DocumentController snapshots for a single sheet.
 *
 * @param {{
 *   beforeSnapshot: Uint8Array;
 *   afterSnapshot: Uint8Array;
 *   sheetId: string;
 * }} opts
 */
export function diffDocumentSnapshots(opts) {
  const before = sheetStateFromDocumentSnapshot(opts.beforeSnapshot, { sheetId: opts.sheetId });
  const after = sheetStateFromDocumentSnapshot(opts.afterSnapshot, { sheetId: opts.sheetId });
  return semanticDiff(before, after);
}

