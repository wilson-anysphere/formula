import { semanticDiff } from "../diff/semanticDiff.js";
import { sheetStateFromDocumentSnapshot } from "./sheetState.js";
import { diffDocumentWorkbookSnapshots } from "./diffWorkbookSnapshots.js";

/**
 * Compute a semantic diff between a stored version snapshot and the current
 * in-memory DocumentController state.
 *
 * This is the primary helper for "compare mode" in the desktop UI when the
 * document model is the local DocumentController.
 *
 * @param {{
 *   versionManager: { doc: { encodeState(): Uint8Array }, getVersion(versionId: string): Promise<{ snapshot: Uint8Array } | null> };
 *   versionId: string;
 *   sheetId: string;
 * }} opts
 */
export async function diffDocumentVersionAgainstCurrent(opts) {
  const version = await opts.versionManager.getVersion(opts.versionId);
  if (!version) throw new Error(`Version not found: ${opts.versionId}`);

  const before = sheetStateFromDocumentSnapshot(version.snapshot, { sheetId: opts.sheetId });

  // Prefer the direct sheet export when available to avoid serializing the entire workbook.
  const doc = opts.versionManager.doc;
  const after =
    typeof doc.exportSheetForSemanticDiff === "function"
      ? doc.exportSheetForSemanticDiff(opts.sheetId)
      : sheetStateFromDocumentSnapshot(doc.encodeState(), { sheetId: opts.sheetId });

  return semanticDiff(before, after);
}

/**
 * Compute a workbook-level diff between a stored version snapshot and the
 * current in-memory DocumentController state.
 *
 * @param {{
 *   versionManager: { doc: { encodeState(): Uint8Array }, getVersion(versionId: string): Promise<{ snapshot: Uint8Array } | null> };
 *   versionId: string;
 * }} opts
 */
export async function diffDocumentWorkbookVersionAgainstCurrent(opts) {
  const version = await opts.versionManager.getVersion(opts.versionId);
  if (!version) throw new Error(`Version not found: ${opts.versionId}`);
  return diffDocumentWorkbookSnapshots({
    beforeSnapshot: version.snapshot,
    afterSnapshot: opts.versionManager.doc.encodeState(),
  });
}

/**
 * Compute a semantic diff between two stored versions.
 *
 * @param {{
 *   getVersion: (versionId: string) => Promise<{ snapshot: Uint8Array } | null>;
 *   beforeVersionId: string;
 *   afterVersionId: string;
 *   sheetId: string;
 * }} opts
 */
export async function diffDocumentVersions(opts) {
  const before = await opts.getVersion(opts.beforeVersionId);
  if (!before) throw new Error(`Version not found: ${opts.beforeVersionId}`);
  const after = await opts.getVersion(opts.afterVersionId);
  if (!after) throw new Error(`Version not found: ${opts.afterVersionId}`);

  const beforeState = sheetStateFromDocumentSnapshot(before.snapshot, { sheetId: opts.sheetId });
  const afterState = sheetStateFromDocumentSnapshot(after.snapshot, { sheetId: opts.sheetId });
  return semanticDiff(beforeState, afterState);
}

/**
 * Compute a workbook-level diff between two stored versions.
 *
 * @param {{
 *   getVersion: (versionId: string) => Promise<{ snapshot: Uint8Array } | null>;
 *   beforeVersionId: string;
 *   afterVersionId: string;
 * }} opts
 */
export async function diffDocumentWorkbookVersions(opts) {
  const before = await opts.getVersion(opts.beforeVersionId);
  if (!before) throw new Error(`Version not found: ${opts.beforeVersionId}`);
  const after = await opts.getVersion(opts.afterVersionId);
  if (!after) throw new Error(`Version not found: ${opts.afterVersionId}`);
  return diffDocumentWorkbookSnapshots({ beforeSnapshot: before.snapshot, afterSnapshot: after.snapshot });
}
