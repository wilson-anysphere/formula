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

function isYMapLike(value: unknown): value is {
  get: (...args: any[]) => any;
  set: (...args: any[]) => any;
  delete: (...args: any[]) => any;
  observeDeep: (...args: any[]) => any;
  unobserveDeep: (...args: any[]) => any;
} {
  if (!value || typeof value !== "object") return false;
  const maybe = value as any;
  // Plain JS Maps also have get/set/delete, so additionally require Yjs'
  // deep observer APIs.
  return (
    typeof maybe.get === "function" &&
    typeof maybe.set === "function" &&
    typeof maybe.delete === "function" &&
    typeof maybe.observeDeep === "function" &&
    typeof maybe.unobserveDeep === "function"
  );
}

function isYjsVersionStore(store: VersionStore): store is YjsVersionStore {
  // Prefer `instanceof`, but tolerate multiple copies of the class and/or mixed
  // module instances (ESM + CJS) via structural checks.
  if (store instanceof (YjsVersionStore as any)) return true;
  const maybe = store as any;
  if (!maybe || typeof maybe !== "object") return false;
  const doc = maybe.doc;
  if (!doc || typeof doc !== "object") return false;
  if (typeof doc.getMap !== "function") return false;
  if (!isYMapLike(maybe.versions)) return false;
  if (!isYMapLike(maybe.meta)) return false;
  return true;
}

export type CollabVersioningOptions = {
  session: CollabSession;
  /**
   * VersionStore implementation.
   *
   * If omitted, defaults to {@link YjsVersionStore} so version history is stored
   * inside the collaborative Yjs document and syncs via sync-server/y-websocket.
   */
  store?: VersionStore;
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
   * (e.g. `cellStructuralOps`, `branching:*`, and `versions*` when using
   * {@link YjsVersionStore}). This option lets callers extend that list (for
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

    const store = opts.store ?? new YjsVersionStore({ doc: opts.session.doc });

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
    // Additionally, when version history itself is stored in the Yjs doc
    // (YjsVersionStore), we must exclude those roots from snapshots/restores to
    // avoid recursive snapshots and to prevent restores from rolling back
    // history.
    const storeInDoc = isYjsVersionStore(store);
    const builtInExcludeRoots = [
      // Always excluded internal collaboration roots.
      "cellStructuralOps",
      "branching:branches",
      "branching:commits",
      "branching:meta",
      // Exclude versioning history roots only when history is stored in-doc.
      ...(storeInDoc ? ["versions", "versionsMeta"] : []),
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
   * Stops auto snapshot timers created by {@link VersionManager}.
   *
   * This does not destroy the underlying CollabSession or VersionStore.
   */
  destroy(): void {
    this.manager.stopAutoSnapshot();
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
