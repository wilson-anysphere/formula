import { EventEmitter } from "node:events";

/**
 * @typedef {{ userId: string, userName: string }} UserInfo
 *
 * @typedef {"snapshot" | "checkpoint" | "restore"} VersionKind
 *
 * @typedef {{
 *   id: string;
 *   kind: VersionKind;
 *   timestampMs: number;
 *   userId: string | null;
 *   userName: string | null;
 *   description: string | null;
 *   checkpointName: string | null;
 *   checkpointLocked: boolean | null;
 *   checkpointAnnotations: string | null;
 *   snapshot: Uint8Array;
 * }} VersionRecord
 *
 * @typedef {{
 *   saveVersion(version: VersionRecord): Promise<void>;
 *   getVersion(versionId: string): Promise<VersionRecord | null>;
 *   listVersions(): Promise<VersionRecord[]>;
 *   updateVersion(versionId: string, patch: { checkpointLocked?: boolean }): Promise<void>;
 * }} VersionStore
 *
 * @typedef {{
 *   encodeState(): Uint8Array;
 *   applyState(snapshot: Uint8Array): void;
 *   on(event: "update", listener: () => void): any;
 * }} VersionedDoc
 */

/**
 * VersionManager creates immutable snapshots of a document over time.
 *
 * The implementation is intentionally storage-agnostic (VersionStore) and
 * document-agnostic (VersionedDoc adapter) so it can be used with a Yjs doc in
 * production and a lightweight fake doc in tests.
 */
export class VersionManager extends EventEmitter {
  /**
   * @param {{
   *   doc: VersionedDoc;
   *   store: VersionStore;
   *   user?: Partial<UserInfo>;
   *   autoSnapshotIntervalMs?: number;
   *   nowMs?: () => number;
   *   autoStart?: boolean;
   * }} opts
   */
  constructor(opts) {
    super();
    this.doc = opts.doc;
    this.store = opts.store;
    this.userId = opts.user?.userId ?? null;
    this.userName = opts.user?.userName ?? null;
    this.autoSnapshotIntervalMs = opts.autoSnapshotIntervalMs ?? 5 * 60 * 1000;
    this.nowMs = opts.nowMs ?? (() => Date.now());
    this.dirty = false;
    this._timer = null;

    // Mark dirty on any document update.
    if (this.doc?.on) {
      this.doc.on("update", () => {
        this.markDirty();
      });
    }

    if (opts.autoStart ?? true) {
      this.startAutoSnapshot();
    }
  }

  markDirty() {
    this.dirty = true;
  }

  /**
   * @returns {Promise<VersionRecord[]>}
   */
  async listVersions() {
    return this.store.listVersions();
  }

  /**
   * Create a periodic snapshot (auto-save) iff the document is dirty.
   * @returns {Promise<VersionRecord | null>}
   */
  async maybeSnapshot() {
    if (!this.dirty) return null;
    const version = await this._createVersion({
      kind: "snapshot",
      description: "Auto-save",
    });
    this.dirty = false;
    return version;
  }

  /**
   * @param {{ description?: string }} opts
   */
  async createSnapshot(opts = {}) {
    const version = await this._createVersion({
      kind: "snapshot",
      description: opts.description ?? null,
    });
    this.dirty = false;
    return version;
  }

  /**
   * @param {{ name: string, annotations?: string, locked?: boolean }} opts
   */
  async createCheckpoint(opts) {
    const version = await this._createVersion({
      kind: "checkpoint",
      description: opts.name,
      checkpointName: opts.name,
      checkpointAnnotations: opts.annotations ?? null,
      checkpointLocked: opts.locked ?? false,
    });
    this.dirty = false;
    return version;
  }

  /**
   * @param {string} checkpointId
   * @param {boolean} locked
   */
  async setCheckpointLocked(checkpointId, locked) {
    const v = await this.store.getVersion(checkpointId);
    if (!v) throw new Error(`Checkpoint not found: ${checkpointId}`);
    if (v.kind !== "checkpoint") throw new Error(`Not a checkpoint: ${checkpointId}`);
    await this.store.updateVersion(checkpointId, { checkpointLocked: locked });
  }

  /**
   * Restores the document to the given version, but records the action by
   * creating a new "restore" head version (history remains intact).
   *
   * @param {string} versionId
   */
  async restoreVersion(versionId) {
    const target = await this.store.getVersion(versionId);
    if (!target) throw new Error(`Version not found: ${versionId}`);

    // Apply snapshot to the live document.
    this.doc.applyState(target.snapshot);

    // Record new head version with the restored snapshot bytes.
    const restoredSnapshot = this.doc.encodeState();
    const head = /** @type {VersionRecord} */ ({
      id: crypto.randomUUID(),
      kind: "restore",
      timestampMs: this.nowMs(),
      userId: this.userId,
      userName: this.userName,
      description: `Restored ${versionId}`,
      checkpointName: null,
      checkpointLocked: null,
      checkpointAnnotations: null,
      snapshot: restoredSnapshot,
    });
    await this.store.saveVersion(head);

    this.emit("restored", { from: versionId, to: head.id });
    this.dirty = false;
    return head;
  }

  startAutoSnapshot() {
    if (this._timer) return;
    this._timer = setInterval(() => {
      void this.maybeSnapshot();
    }, this.autoSnapshotIntervalMs);
  }

  stopAutoSnapshot() {
    if (!this._timer) return;
    clearInterval(this._timer);
    this._timer = null;
  }

  /**
   * @param {{
   *   kind: VersionKind,
   *   description?: string | null,
   *   checkpointName?: string | null,
   *   checkpointLocked?: boolean | null,
   *   checkpointAnnotations?: string | null,
   * }} opts
   */
  async _createVersion(opts) {
    const snapshot = this.doc.encodeState();
    const version = /** @type {VersionRecord} */ ({
      id: crypto.randomUUID(),
      kind: opts.kind,
      timestampMs: this.nowMs(),
      userId: this.userId,
      userName: this.userName,
      description: opts.description ?? null,
      checkpointName: opts.checkpointName ?? null,
      checkpointLocked: opts.checkpointLocked ?? null,
      checkpointAnnotations: opts.checkpointAnnotations ?? null,
      snapshot,
    });
    await this.store.saveVersion(version);
    this.emit("versionCreated", version);
    return version;
  }
}

