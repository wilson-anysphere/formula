import type { CollabSession } from "@formula/collab-session";

import { VersionManager } from "../../../versioning/src/versioning/versionManager.js";
import { createYjsSpreadsheetDocAdapter } from "../../../versioning/src/yjs/yjsSpreadsheetDocAdapter.js";

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

export type CollabVersioningOptions = {
  session: CollabSession;
  store: VersionStore;
  user?: Partial<CollabVersioningUser>;
  retention?: VersionRetention;
  autoSnapshotIntervalMs?: number;
  /**
   * Whether to start the VersionManager auto-snapshot interval immediately.
   *
   * Defaults to `true` (matches VersionManager).
   */
  autoStart?: boolean;
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
    const doc = createYjsSpreadsheetDocAdapter(opts.session.doc);
    this.manager = new VersionManager({
      doc,
      store: opts.store,
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
