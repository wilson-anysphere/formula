import type { CollabSession } from "@formula/collab-session";

import { VersionManager } from "../../../versioning/src/versioning/versionManager.js";
import { createYjsSpreadsheetDocAdapter } from "../../../versioning/src/yjs/yjsSpreadsheetDocAdapter.js";
import { YjsVersionStore } from "../../../versioning/src/store/yjsVersionStore.js";

export type CollabVersioningUser = { userId: string; userName: string };

export type VersionKind = "snapshot" | "checkpoint" | "restore";

export type VersionRecord = {
  id: string;
  kind: VersionKind;
  timestampMs: number;
  userId: string | null;
  userName: string | null;
  description: string | null;
  checkpointName: string | null;
  checkpointLocked: boolean | null;
  checkpointAnnotations: string | null;
  snapshot: Uint8Array;
};

export type VersionRetention = {
  maxSnapshots?: number;
  maxAgeMs?: number;
  keepRestores?: boolean;
  keepCheckpoints?: boolean;
};

export type VersionStore = {
  saveVersion(version: VersionRecord): Promise<void>;
  getVersion(versionId: string): Promise<VersionRecord | null>;
  listVersions(): Promise<VersionRecord[]>;
  updateVersion(versionId: string, patch: { checkpointLocked?: boolean }): Promise<void>;
  deleteVersion(versionId: string): Promise<void>;
};

export type DefaultYjsVersionStoreOptions = {
  /**
   * Controls how snapshot blobs are written into the collaborative Y.Doc.
   *
   * - `"single"` writes the full snapshot in a single Yjs transaction/update
   *   (legacy behavior; can exceed websocket message size limits for large docs).
   * - `"stream"` appends snapshot chunks across multiple transactions/updates so
   *   each update stays small.
   */
  writeMode?: "single" | "stream";
  /**
   * Snapshot chunk size in bytes when storing as `snapshotChunks`.
   */
  chunkSize?: number;
  /**
   * Maximum number of chunks appended per Yjs transaction when `writeMode:
   * "stream"`.
   */
  maxChunksPerTransaction?: number | null;
  /**
   * Optional snapshot compression.
   *
   * Note: `"gzip"` is only supported in environments with either Node's `zlib`
   * or the web `CompressionStream` API. When unsupported, `YjsVersionStore` will
   * throw when saving/loading versions. Defaults to `"none"`.
   */
  compression?: "none" | "gzip";
  /**
   * How the snapshot bytes are encoded inside the Y.Doc.
   *
   * Note: when `writeMode: "stream"`, snapshots are always stored as `"chunks"`
   * regardless of this option (so streaming mode can append bytes across many
   * transactions).
   */
  snapshotEncoding?: "chunks" | "base64";
};

export type CollabVersioningOptions = {
  session: CollabSession;
  /**
   * VersionStore implementation.
   *
   * If omitted, defaults to {@link YjsVersionStore} so version history is stored
   * inside the collaborative Yjs document and syncs via sync-server/y-websocket.
   */
  store?: VersionStore;
  /**
   * Options used when constructing the default {@link YjsVersionStore}.
   *
   * This is ignored when `store` is provided.
   *
   * Defaults are tuned to be robust under sync-server websocket message limits:
   * `writeMode: "stream"`, `chunkSize: 64KiB`, and a conservative
   * `maxChunksPerTransaction` (8 at the default chunk size).
   */
  yjsStoreOptions?: DefaultYjsVersionStoreOptions;
  user?: Partial<CollabVersioningUser>;
  retention?: VersionRetention;
  autoSnapshotIntervalMs?: number;
  /**
   * Whether to start the VersionManager auto-snapshot interval immediately.
   *
   * Defaults to `true` (matches VersionManager).
   */
  autoStart?: boolean;
  /**
   * Additional Yjs root names to exclude from version snapshots/restores.
   *
   * CollabVersioning always excludes built-in internal collaboration roots
   * (e.g. `cellStructuralOps`, `branching:*`, and the reserved versioning roots
   * `versions`/`versionsMeta`). This option lets callers extend that list (for
   * example when {@link CollabBranchingWorkflow} is configured with a non-default
   * branch root name).
   */
  excludeRoots?: string[];
};

/**
 * Glue layer that connects {@link VersionManager} to a {@link CollabSession}'s Y.Doc.
 *
 * Key property: restoring a version mutates the existing Y.Doc instance, so the
 * collab provider + awareness remain intact and restore propagates to other
 * collaborators via normal Yjs updates.
 */
export class CollabVersioning {
  readonly session: CollabSession;
  readonly manager: VersionManager;

  constructor(opts: CollabVersioningOptions) {
    this.session = opts.session;

    const store =
      opts.store ??
      (() => {
        // Default to streaming snapshot writes so large version snapshots are
        // split across many small Yjs updates. This creates more Yjs transactions
        // (and therefore more websocket messages), but avoids catastrophic
        // message-size failures (WS close code 1009) when snapshots exceed the
        // sync-server's `SYNC_SERVER_MAX_MESSAGE_BYTES`.
        const DEFAULT_CHUNK_SIZE_BYTES = 64 * 1024;
        const DEFAULT_TARGET_UPDATE_BYTES = 512 * 1024;
        const DEFAULT_MAX_CHUNKS = 16;

        const userOpts = opts.yjsStoreOptions ?? {};
        const chunkSize = userOpts.chunkSize ?? DEFAULT_CHUNK_SIZE_BYTES;

        // Keep each streamed update comfortably below typical sync-server message limits
        // while avoiding excessively chatty streams for very large docs.
        const defaultMaxChunksPerTransaction = Math.max(
          1,
          Math.min(DEFAULT_MAX_CHUNKS, Math.floor(DEFAULT_TARGET_UPDATE_BYTES / Math.max(1, chunkSize))),
        );

        const maxChunksPerTransaction =
          userOpts.maxChunksPerTransaction === undefined
            ? defaultMaxChunksPerTransaction
            : userOpts.maxChunksPerTransaction;

        return new YjsVersionStore({
          ...userOpts,
          // Always bind the store to the session's doc (ignore any accidental `doc`
          // field passed via `yjsStoreOptions`).
          doc: opts.session.doc,
          writeMode: userOpts.writeMode ?? "stream",
          chunkSize,
          maxChunksPerTransaction,
        });
      })();

    // CollabVersioning snapshots/restores should only affect user workbook state,
    // not internal collaboration metadata stored inside the same Y.Doc.
    //
    // Always exclude:
    // - Structural conflict op log (`cellStructuralOps`)
    // - Default YjsBranchStore graph roots (`branching:*`)
    //
    // Note: CollabBranchingWorkflow allows configuring the branch root name
    // (default "branching"). We only exclude the default roots here; callers
    // can extend the list via `CollabVersioningOptions.excludeRoots`.
    //
    // Additionally, we always exclude version history roots (`versions`,
    // `versionsMeta`). When history is stored in-doc (YjsVersionStore) this
    // avoids recursive snapshots and prevents restores from rolling back
    // internal history.
    //
    // Even when history is stored out-of-doc (Api/SQLite/etc), a document may
    // still contain these reserved roots (e.g. from earlier in-doc usage or dev
    // sessions). Restoring must never attempt to rewrite them, both to avoid
    // rewinding internal state and to prevent server-side "reserved root
    // mutation" disconnects.
    const builtInExcludeRoots = [
      // Always excluded internal collaboration roots.
      "cellStructuralOps",
      "branching:branches",
      "branching:commits",
      "branching:meta",
      // Always excluded internal versioning roots.
      "versions",
      "versionsMeta",
    ];

    const excludeRoots = Array.from(
      new Set([...builtInExcludeRoots, ...(Array.isArray(opts.excludeRoots) ? opts.excludeRoots : [])]),
    );

    const doc = createYjsSpreadsheetDocAdapter(opts.session.doc, { excludeRoots });
    this.manager = new VersionManager({
      doc,
      store,
      user: opts.user,
      retention: opts.retention,
      autoSnapshotIntervalMs: opts.autoSnapshotIntervalMs,
      autoStart: opts.autoStart,
    });
  }

  /**
   * Stops auto snapshot timers created by {@link VersionManager} and removes
   * its document update listeners.
   *
   * This does not destroy the underlying CollabSession or VersionStore.
   */
  destroy(): void {
    this.manager.destroy();
  }

  listVersions(): Promise<VersionRecord[]> {
    return this.manager.listVersions() as Promise<VersionRecord[]>;
  }

  getVersion(versionId: string): Promise<VersionRecord | null> {
    return this.manager.getVersion(versionId) as Promise<VersionRecord | null>;
  }

  createSnapshot(opts: { description?: string } = {}): Promise<VersionRecord> {
    return this.manager.createSnapshot(opts) as Promise<VersionRecord>;
  }

  createCheckpoint(opts: {
    name: string;
    annotations?: string;
    locked?: boolean;
  }): Promise<VersionRecord> {
    return this.manager.createCheckpoint(opts) as Promise<VersionRecord>;
  }

  restoreVersion(versionId: string): Promise<VersionRecord> {
    return this.manager.restoreVersion(versionId) as Promise<VersionRecord>;
  }

  setCheckpointLocked(checkpointId: string, locked: boolean): Promise<void> {
    return this.manager.setCheckpointLocked(checkpointId, locked);
  }

  deleteVersion(versionId: string): Promise<void> {
    return this.manager.deleteVersion(versionId);
  }
}

export function createCollabVersioning(opts: CollabVersioningOptions): CollabVersioning {
  return new CollabVersioning(opts);
}
