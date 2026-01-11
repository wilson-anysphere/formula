import { VersionManager } from "../../versioning/src/versioning/versionManager.js";
import { createYjsSpreadsheetDocAdapter } from "../../versioning/src/yjs/yjsSpreadsheetDocAdapter.js";
import { YjsVersionStore } from "../../versioning/src/store/yjsVersionStore.js";

const DEFAULT_EXCLUDED_ROOTS = ["versions", "versionsMeta"];

/**
 * CollabVersioning is glue between:
 * - a collaborative Yjs document (CollabSession / y-websocket / sync-server)
 * - VersionManager (snapshots/checkpoints/restores)
 * - a VersionStore implementation
 *
 * When using `YjsVersionStore`, version history is itself a collaborative Yjs
 * artifact stored inside the same document. In that configuration we must
 * exclude the version-history roots from snapshots/restores to avoid
 * self-referential snapshots and accidental history rollback during restores.
 *
 * @typedef {import("../../versioning/src/versioning/versionManager.js").VersionRetention} VersionRetention
 *
 * @typedef {{
 *   ydoc: import("yjs").Doc;
 *   store?: import("../../versioning/src/versioning/versionManager.js").VersionStore;
 *   user?: { userId?: string, userName?: string };
 *   autoSnapshotIntervalMs?: number;
 *   nowMs?: () => number;
 *   autoStart?: boolean;
 *   retention?: VersionRetention;
 * }} CollabVersioningOptions
 */

/**
 * @param {CollabVersioningOptions} opts
 * @returns {VersionManager}
 */
export function createCollabVersioning(opts) {
  if (!opts?.ydoc) throw new Error("createCollabVersioning: ydoc is required");

  const store = opts.store ?? new YjsVersionStore({ doc: opts.ydoc });

  const doc = createYjsSpreadsheetDocAdapter(opts.ydoc, {
    excludeRoots: DEFAULT_EXCLUDED_ROOTS,
  });

  return new VersionManager({
    doc,
    store,
    user: opts.user,
    autoSnapshotIntervalMs: opts.autoSnapshotIntervalMs,
    nowMs: opts.nowMs,
    autoStart: opts.autoStart,
    retention: opts.retention,
  });
}

// Backwards-compatible alias.
export const createCollabVersionManager = createCollabVersioning;

