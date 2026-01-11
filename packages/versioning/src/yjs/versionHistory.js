import { diffYjsSnapshots } from "./diffSnapshots.js";
import { diffYjsWorkbookSnapshots } from "./diffWorkbookSnapshots.js";

/**
 * Compute a semantic diff between a stored version snapshot and the current
 * in-memory Yjs document state.
 *
 * This is the primary helper for "compare mode" in the UI: select a checkpoint
 * from history and see what has changed in the current grid.
 *
 * @param {{
 *   versionManager: { doc: { encodeState(): Uint8Array }, getVersion(versionId: string): Promise<{ snapshot: Uint8Array } | null> };
 *   versionId: string;
 *   sheetId?: string | null;
 * }} opts
 */
export async function diffYjsVersionAgainstCurrent(opts) {
  const version = await opts.versionManager.getVersion(opts.versionId);
  if (!version) throw new Error(`Version not found: ${opts.versionId}`);
  const currentSnapshot = opts.versionManager.doc.encodeState();
  return diffYjsSnapshots({
    beforeSnapshot: version.snapshot,
    afterSnapshot: currentSnapshot,
    sheetId: opts.sheetId ?? null,
  });
}

/**
 * Compute a workbook-level diff between a stored version snapshot and the
 * current in-memory Yjs document state.
 *
 * @param {{
 *   versionManager: { doc: { encodeState(): Uint8Array }, getVersion(versionId: string): Promise<{ snapshot: Uint8Array } | null> };
 *   versionId: string;
 * }} opts
 */
export async function diffYjsWorkbookVersionAgainstCurrent(opts) {
  const version = await opts.versionManager.getVersion(opts.versionId);
  if (!version) throw new Error(`Version not found: ${opts.versionId}`);
  const currentSnapshot = opts.versionManager.doc.encodeState();
  return diffYjsWorkbookSnapshots({ beforeSnapshot: version.snapshot, afterSnapshot: currentSnapshot });
}

/**
 * Compute a semantic diff between two stored versions.
 *
 * @param {{
 *   getVersion: (versionId: string) => Promise<{ snapshot: Uint8Array } | null>;
 *   beforeVersionId: string;
 *   afterVersionId: string;
 *   sheetId?: string | null;
 * }} opts
 */
export async function diffYjsVersions(opts) {
  const before = await opts.getVersion(opts.beforeVersionId);
  if (!before) throw new Error(`Version not found: ${opts.beforeVersionId}`);
  const after = await opts.getVersion(opts.afterVersionId);
  if (!after) throw new Error(`Version not found: ${opts.afterVersionId}`);
  return diffYjsSnapshots({
    beforeSnapshot: before.snapshot,
    afterSnapshot: after.snapshot,
    sheetId: opts.sheetId ?? null,
  });
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
export async function diffYjsWorkbookVersions(opts) {
  const before = await opts.getVersion(opts.beforeVersionId);
  if (!before) throw new Error(`Version not found: ${opts.beforeVersionId}`);
  const after = await opts.getVersion(opts.afterVersionId);
  if (!after) throw new Error(`Version not found: ${opts.afterVersionId}`);
  return diffYjsWorkbookSnapshots({ beforeSnapshot: before.snapshot, afterSnapshot: after.snapshot });
}
