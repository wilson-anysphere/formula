import { semanticDiff } from "../diff/semanticDiff.js";
import { sheetStateFromYjsSnapshot } from "./sheetState.js";

/**
 * Compute a semantic diff between two Yjs snapshots for a single sheet.
 *
 * @param {{
 *   beforeSnapshot: Uint8Array;
 *   afterSnapshot: Uint8Array;
 *   sheetId?: string | null;
 * }} opts
 */
export function diffYjsSnapshots(opts) {
  const before = sheetStateFromYjsSnapshot(opts.beforeSnapshot, { sheetId: opts.sheetId ?? null });
  const after = sheetStateFromYjsSnapshot(opts.afterSnapshot, { sheetId: opts.sheetId ?? null });
  return semanticDiff(before, after);
}

