import { diffDocumentSnapshots } from "./diffSnapshots.js";

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
  const currentSnapshot = opts.versionManager.doc.encodeState();
  return diffDocumentSnapshots({
    beforeSnapshot: version.snapshot,
    afterSnapshot: currentSnapshot,
    sheetId: opts.sheetId,
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
  return diffDocumentSnapshots({
    beforeSnapshot: before.snapshot,
    afterSnapshot: after.snapshot,
    sheetId: opts.sheetId,
  });
}

