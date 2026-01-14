/**
 * Minimal EventEmitter implementation that works in both browser and Node runtimes.
 *
 * The versioning package is consumed by the desktop/web UI bundles (Vite), so we
 * cannot depend on Node built-ins like `node:events` here.
 *
 * Only the small subset of the Node EventEmitter API that we rely on is
 * implemented (`on`, `off`/`removeListener`, `once`, `emit`).
 */
class SimpleEventEmitter {
  constructor() {
    /** @type {Map<any, Set<(...args: any[]) => void>>} */
    this._listeners = new Map();
  }

  /**
   * @param {string | symbol} event
   * @param {(...args: any[]) => void} listener
   */
  on(event, listener) {
    let set = this._listeners.get(event);
    if (!set) {
      set = new Set();
      this._listeners.set(event, set);
    }
    set.add(listener);
    return this;
  }

  /**
   * Alias for Node compatibility.
   */
  addListener(event, listener) {
    return this.on(event, listener);
  }

  /**
   * @param {string | symbol} event
   * @param {(...args: any[]) => void} listener
   */
  off(event, listener) {
    const set = this._listeners.get(event);
    if (!set) return this;
    set.delete(listener);
    if (set.size === 0) this._listeners.delete(event);
    return this;
  }

  /**
   * Alias for Node compatibility.
   */
  removeListener(event, listener) {
    return this.off(event, listener);
  }

  /**
   * @param {string | symbol} event
   * @param {(...args: any[]) => void} listener
   */
  once(event, listener) {
    const wrapped = (...args) => {
      this.off(event, wrapped);
      listener(...args);
    };
    return this.on(event, wrapped);
  }

  /**
   * @param {string | symbol} event
   * @param  {...any} args
   */
  emit(event, ...args) {
    const set = this._listeners.get(event);
    if (!set || set.size === 0) return false;
    // Snapshot so listeners can mutate subscriptions safely.
    for (const listener of Array.from(set)) {
      try {
        listener(...args);
      } catch (err) {
        // Avoid crashing the app due to a listener; surface the failure asynchronously
        // like many event emitter implementations do.
        const enqueue =
          typeof queueMicrotask === "function"
            ? queueMicrotask
            : (cb) => {
                void Promise.resolve()
                  .then(cb)
                  .catch((thrown) => {
                    // Promise microtasks turn thrown errors into rejections; rethrow on a timer so
                    // failures surface as an error instead of an unhandled promise rejection.
                    if (typeof setTimeout === "function") {
                      setTimeout(() => {
                        throw thrown;
                      }, 0);
                      return;
                    }
                    // eslint-disable-next-line no-console
                    console.error(thrown);
                  });
              };
        enqueue(() => {
          throw err;
        });
      }
    }
    return true;
  }

  removeAllListeners(event) {
    if (event == null) {
      this._listeners.clear();
      return this;
    }
    this._listeners.delete(event);
    return this;
  }
}

function randomUUID() {
  const cryptoObj = typeof globalThis !== "undefined" ? globalThis.crypto : undefined;
  const randomUuid = cryptoObj && typeof cryptoObj.randomUUID === "function" ? cryptoObj.randomUUID : null;
  if (randomUuid) {
    try {
      return randomUuid.call(cryptoObj);
    } catch {
      // fall through
    }
  }

  const getRandomValues = cryptoObj && typeof cryptoObj.getRandomValues === "function" ? cryptoObj.getRandomValues : null;
  if (getRandomValues) {
    const bytes = new Uint8Array(16);
    getRandomValues.call(cryptoObj, bytes);
    // RFC 4122 v4 UUID.
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    const hex = Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
    return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
  }

  return `${Date.now().toString(16)}-${Math.random().toString(16).slice(2)}`;
}

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
 *   maxSnapshots?: number;
 *   maxAgeMs?: number;
 *   keepRestores?: boolean;
 *   keepCheckpoints?: boolean;
 * }} VersionRetention
 *
 * @typedef {{
 *   saveVersion(version: VersionRecord): Promise<void>;
 *   getVersion(versionId: string): Promise<VersionRecord | null>;
 *   listVersions(): Promise<VersionRecord[]>;
 *   updateVersion(versionId: string, patch: { checkpointLocked?: boolean }): Promise<void>;
 *   deleteVersion(versionId: string): Promise<void>;
 * }} VersionStore
 *
 * @typedef {{
 *   encodeState(): Uint8Array;
 *   applyState(snapshot: Uint8Array): void;
 *   on(event: "update", listener: () => void): void | (() => void);
 * }} VersionedDoc
 */

/**
 * VersionManager creates immutable snapshots of a document over time.
 *
 * The implementation is intentionally storage-agnostic (VersionStore) and
 * document-agnostic (VersionedDoc adapter) so it can be used with a Yjs doc in
 * production and a lightweight fake doc in tests.
 */
export class VersionManager extends SimpleEventEmitter {
  /**
   * @param {{
   *   doc: VersionedDoc;
   *   store: VersionStore;
   *   user?: Partial<UserInfo>;
   *   autoSnapshotIntervalMs?: number;
   *   nowMs?: () => number;
   *   autoStart?: boolean;
   *   retention?: VersionRetention;
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
    this._autoSnapshotInFlight = false;
    this._destroyed = false;
    /** @type {Promise<void>} */
    this._pruneChain = Promise.resolve();
    /** @type {null | (() => void)} */
    this._unsubscribeDocUpdates = null;
    /** @type {null | (() => void)} */
    this._docUpdateListener = null;
    this.retention = opts.retention
      ? {
          maxSnapshots: opts.retention.maxSnapshots,
          maxAgeMs: opts.retention.maxAgeMs,
          keepRestores: opts.retention.keepRestores ?? true,
          keepCheckpoints: opts.retention.keepCheckpoints ?? true,
        }
      : null;

    // Mark dirty on any document update.
    if (this.doc?.on) {
      const onUpdate = () => {
        this.markDirty();
      };
      this._docUpdateListener = onUpdate;
      const maybeUnsubscribe = this.doc.on("update", onUpdate);
      if (typeof maybeUnsubscribe === "function") {
        this._unsubscribeDocUpdates = maybeUnsubscribe;
      }
    }

    if (opts.autoStart ?? true) {
      this.startAutoSnapshot();
    }
  }

  markDirty() {
    if (this._destroyed) return;
    this.dirty = true;
  }

  destroy() {
    if (this._destroyed) return;
    this._destroyed = true;
    // Treat destroyed managers as permanently clean so they cannot create
    // snapshots (even if they happened to be dirty at the time of teardown).
    this.dirty = false;
    this.stopAutoSnapshot();
    const unsubscribe = this._unsubscribeDocUpdates;
    this._unsubscribeDocUpdates = null;
    if (typeof unsubscribe === "function") {
      try {
        unsubscribe();
      } catch {
        // ignore
      }
    } else if (this.doc && this._docUpdateListener) {
      const off =
        typeof this.doc.off === "function"
          ? this.doc.off
          : typeof this.doc.removeListener === "function"
            ? this.doc.removeListener
            : null;
      if (typeof off === "function") {
        try {
          off.call(this.doc, "update", this._docUpdateListener);
        } catch {
          // ignore
        }
      }
    }
    this._docUpdateListener = null;
    this.removeAllListeners();
  }

  /**
   * @returns {Promise<VersionRecord[]>}
   */
  async listVersions() {
    return this.store.listVersions();
  }

  /**
   * @param {string} versionId
   * @returns {Promise<VersionRecord | null>}
   */
  async getVersion(versionId) {
    return this.store.getVersion(versionId);
  }

  /**
   * @param {string} versionId
   */
  async deleteVersion(versionId) {
    await this.store.deleteVersion(versionId);
    this.emit("versionDeleted", { id: versionId });
  }

  /**
   * Create a periodic snapshot (auto-save) iff the document is dirty.
   * @returns {Promise<VersionRecord | null>}
   */
  async maybeSnapshot() {
    if (this._destroyed) return null;
    if (!this.dirty) return null;
    const version = await this._createVersion({
      kind: "snapshot",
    });
    // If we're destroyed while an autosnapshot is in flight (e.g. interval tick
    // awaited store writes), ensure we don't leave behind a stray snapshot.
    if (this._destroyed) {
      try {
        await this.store.deleteVersion(version.id);
      } catch {
        // ignore (best-effort cleanup)
      }
      return null;
    }
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
      id: randomUUID(),
      kind: "restore",
      timestampMs: this.nowMs(),
      userId: this.userId,
      userName: this.userName,
      description: null,
      checkpointName: null,
      checkpointLocked: null,
      checkpointAnnotations: null,
      snapshot: restoredSnapshot,
    });
    await this.store.saveVersion(head);

    this.emit("restored", { from: versionId, to: head.id });
    await this._queueRetention();
    this.dirty = false;
    return head;
  }

  startAutoSnapshot() {
    if (this._destroyed) return;
    if (this._timer) return;
    this._timer = setInterval(() => {
      void this._autoSnapshotTick();
    }, this.autoSnapshotIntervalMs);
  }

  stopAutoSnapshot() {
    if (!this._timer) return;
    clearInterval(this._timer);
    this._timer = null;
  }

  async _autoSnapshotTick() {
    if (this._destroyed) return;
    if (this._autoSnapshotInFlight) return;
    this._autoSnapshotInFlight = true;
    try {
      const created = await this.maybeSnapshot();
      if (this._destroyed) return;
      // Prune on the autosnapshot cadence even when the document is idle.
      if (!created) {
        await this._queueRetention();
      }
    } finally {
      this._autoSnapshotInFlight = false;
    }
  }

  async _queueRetention() {
    if (this._destroyed) return;
    if (!this.retention) return;
    const next = this._pruneChain.catch(() => {}).then(() => this._applyRetention());
    this._pruneChain = next;
    return next;
  }

  async _applyRetention() {
    if (this._destroyed) return;
    const retention = this.retention;
    if (!retention) return;
    if (retention.maxSnapshots == null && retention.maxAgeMs == null) return;

    const versions = await this.store.listVersions();
    if (this._destroyed) return;
    // Preserve the store's ordering as a stable tie-breaker when timestamps are equal.
    // (Some stores maintain an insertion index; re-sorting purely by timestamp can
    // lead to unstable pruning when snapshots are created within the same ms.)
    const originalIndex = new Map(versions.map((v, idx) => [v.id, idx]));
    const ordered = [...versions].sort((a, b) => {
      const dt = b.timestampMs - a.timestampMs;
      if (dt !== 0) return dt;
      return (originalIndex.get(a.id) ?? 0) - (originalIndex.get(b.id) ?? 0);
    });

    /** @type {Set<string>} */
    const deleteIds = new Set();

    const keepRestores = retention.keepRestores ?? true;
    const keepCheckpoints = retention.keepCheckpoints ?? true;

    /**
     * @param {VersionRecord} v
     */
    const protectedFromDeletion = (v) => {
      if (v.kind === "checkpoint") {
        if (v.checkpointLocked === true) return true;
        if (keepCheckpoints) return true;
        return false;
      }
      if (v.kind === "restore") {
        return keepRestores;
      }
      return false;
    };

    if (typeof retention.maxAgeMs === "number") {
      const now = this.nowMs();
      const cutoff = now - retention.maxAgeMs;
      for (const v of ordered) {
        if (protectedFromDeletion(v)) continue;
        if (v.timestampMs < cutoff) {
          deleteIds.add(v.id);
        }
      }
    }

    if (typeof retention.maxSnapshots === "number") {
      const snapshots = ordered.filter((v) => v.kind === "snapshot");
      for (let i = retention.maxSnapshots; i < snapshots.length; i += 1) {
        deleteIds.add(snapshots[i].id);
      }
    }

    if (deleteIds.size === 0) return;

    const deletedIds = ordered.filter((v) => deleteIds.has(v.id)).map((v) => v.id);
    for (const id of deletedIds) {
      if (this._destroyed) return;
      await this.store.deleteVersion(id);
    }
    this.emit("versionsPruned", { deletedIds });
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
      id: randomUUID(),
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
    await this._queueRetention();
    return version;
  }
}
