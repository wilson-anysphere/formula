import * as Y from "yjs";
import { IndexeddbPersistence } from "y-indexeddb";

import type { CollabPersistence, CollabPersistenceBinding } from "./index.js";

type Entry = {
  doc: Y.Doc;
  persistence: IndexeddbPersistence;
  destroyed: Promise<void>;
  resolveDestroyed: () => void;
  onDocDestroy: () => void;
  onDocUpdate: (update: Uint8Array, origin: unknown) => void;
  synced: boolean;
};

const UPDATES_STORE_NAME = "updates";
const DEFAULT_MAX_UPDATES = 500;
const DEFAULT_COMPACT_DEBOUNCE_MS = 250;

export type IndexedDbCollabPersistenceOptions = {
  /**
   * Maximum number of incremental Yjs updates to keep before compacting the IndexedDB
   * log into a snapshot update.
   *
   * - Set to `0` to disable automatic compaction.
   * - Defaults to `500` (chosen to keep load/replay bounded without rewriting snapshots
   *   on every small edit).
   */
  maxUpdates?: number;
  /**
   * Debounce delay (ms) before running a scheduled compaction once `maxUpdates` is exceeded.
   *
   * This avoids rewriting large snapshots repeatedly during bursts of edits.
   */
  compactDebounceMs?: number;
};

function normalizeNonNegativeInt(value: unknown, fallback: number): number {
  const num = Number(value);
  if (!Number.isFinite(num)) return fallback;
  return Math.max(0, Math.trunc(num));
}

/**
 * Browser (IndexedDB) persistence using `y-indexeddb`.
 *
 * In tests, this works with `fake-indexeddb` by installing the IndexedDB globals
 * (`globalThis.indexedDB`, `globalThis.IDBKeyRange`, ...).
 */
export class IndexedDbCollabPersistence implements CollabPersistence {
  private readonly entries = new Map<string, Entry>();
  private readonly queues = new Map<string, Promise<void>>();
  private readonly updateCounts = new Map<string, number>();
  private readonly compactTimers = new Map<string, ReturnType<typeof setTimeout>>();
  private readonly maxUpdates: number;
  private readonly compactDebounceMs: number;

  constructor(opts: IndexedDbCollabPersistenceOptions = {}) {
    this.maxUpdates =
      opts.maxUpdates === undefined ? DEFAULT_MAX_UPDATES : normalizeNonNegativeInt(opts.maxUpdates, DEFAULT_MAX_UPDATES);
    this.compactDebounceMs =
      opts.compactDebounceMs === undefined
        ? DEFAULT_COMPACT_DEBOUNCE_MS
        : normalizeNonNegativeInt(opts.compactDebounceMs, DEFAULT_COMPACT_DEBOUNCE_MS);
  }

  private enqueue(docId: string, task: () => Promise<void>): Promise<void> {
    const prev = this.queues.get(docId) ?? Promise.resolve();
    const next = prev
      .catch(() => {
        // Keep the queue alive even if a previous task failed.
      })
      .then(task);
    this.queues.set(docId, next);

    const cleanup = () => {
      if (this.queues.get(docId) === next) {
        this.queues.delete(docId);
      }
    };
    void next.then(cleanup, cleanup);
    return next;
  }

  private scheduleCompaction(docId: string): void {
    if (this.maxUpdates <= 0) return;
    if (this.compactTimers.has(docId)) return;

    const timer = setTimeout(() => {
      this.compactTimers.delete(docId);
      void this.compact(docId).catch(() => {
        // Best-effort; compaction should never crash the app.
      });
    }, this.compactDebounceMs);
    this.compactTimers.set(docId, timer);
  }

  private async countUpdateRecords(db: any): Promise<number> {
    return await new Promise<number>((resolve) => {
      try {
        const tx = (db as any).transaction([UPDATES_STORE_NAME], "readonly");
        const store = (tx as any).objectStore(UPDATES_STORE_NAME);
        const req = store.count();
        req.onsuccess = () => resolve(Number(req.result) || 0);
        req.onerror = () => resolve(0);
      } catch {
        resolve(0);
      }
    });
  }

  private destroyEntry(docId: string, entry: Entry): void {
    entry.resolveDestroyed();
    entry.doc.off("destroy", entry.onDocDestroy);
    entry.doc.off("update", entry.onDocUpdate);
    entry.persistence.destroy();
    this.entries.delete(docId);
    this.updateCounts.delete(docId);
    const timer = this.compactTimers.get(docId);
    if (timer) clearTimeout(timer);
    this.compactTimers.delete(docId);
  }

  private getOrCreateEntry(docId: string, doc: Y.Doc): Entry {
    const existing = this.entries.get(docId);
    if (existing) {
      if (existing.doc === doc) return existing;
      this.destroyEntry(docId, existing);
    }

    const persistence = new IndexeddbPersistence(docId, doc);
    let resolveDestroyed: () => void = () => {};
    const destroyed = new Promise<void>((resolve) => {
      resolveDestroyed = resolve;
    });

    const entry: Entry = {
      doc,
      persistence,
      destroyed,
      resolveDestroyed,
      onDocDestroy: () => {},
      onDocUpdate: () => {},
      synced: false,
    };

    entry.onDocUpdate = () => {
      if (!entry.synced) return;
      if (this.maxUpdates <= 0) return;

      const count = (this.updateCounts.get(docId) ?? 0) + 1;
      this.updateCounts.set(docId, count);
      if (count >= this.maxUpdates) {
        this.scheduleCompaction(docId);
      }
    };
    doc.on("update", entry.onDocUpdate);

    entry.onDocDestroy = () => {
      // y-indexeddb's `whenSynced` promise can hang forever if the persistence
      // instance is destroyed before the initial sync completes. Ensure any
      // pending `load()` calls are unblocked when the Y.Doc lifecycle ends.
      this.destroyEntry(docId, entry);
    };
    doc.on("destroy", entry.onDocDestroy);

    // Mark as "synced" once the underlying y-indexeddb initial load completes so we:
    // - don't treat replayed hydration updates as "new" updates for compaction counting, and
    // - can optionally compact huge historical logs in the background.
    Promise.resolve(persistence.whenSynced)
      .then(() => {
        if (this.entries.get(docId) !== entry) return;
        entry.synced = true;

        if (this.maxUpdates <= 0) return;
        const db = entry.persistence.db;
        if (!db) return;

        // If the log is already large (e.g. long-lived doc across restarts),
        // schedule an async compaction so future loads stay bounded.
        void this.enqueue(docId, async () => {
          const current = await this.countUpdateRecords(db);
          if (current >= this.maxUpdates) {
            this.updateCounts.set(docId, current);
            this.scheduleCompaction(docId);
          }
        });
      })
      .catch(() => {
        // Best-effort.
      });

    this.entries.set(docId, entry);
    return entry;
  }

  async load(docId: string, doc: Y.Doc): Promise<void> {
    const entry = this.getOrCreateEntry(docId, doc);
    await Promise.race([Promise.resolve(entry.persistence.whenSynced), entry.destroyed]);
  }

  bind(docId: string, doc: Y.Doc): CollabPersistenceBinding {
    const entry = this.getOrCreateEntry(docId, doc);
    return {
      destroy: async () => {
        this.destroyEntry(docId, entry);
      },
    };
  }

  /**
   * Best-effort durability flush for IndexedDB persistence.
   *
   * `y-indexeddb` writes updates to IndexedDB asynchronously and does not expose a
   * built-in API to await the completion of those writes. For Formula's collab
   * persistence contract (`CollabPersistence.flush`), we implement flush by
   * storing a full document snapshot update into the same `updates` object store
   * used by `y-indexeddb`.
   *
   * This guarantees that all in-memory state at the time of the call is durable
   * (even if some incremental updates are still in flight).
   */
  async flush(docId: string): Promise<void> {
    const entry = this.entries.get(docId);
    if (!entry) return;

    await this.enqueue(docId, async () => {
      const ready = await Promise.race([
        Promise.resolve(entry.persistence.whenSynced).then(() => "synced" as const),
        entry.destroyed.then(() => "destroyed" as const),
      ]);
      if (ready === "destroyed") return;

      const db = entry.persistence.db;
      if (!db) return;

      const snapshot = Y.encodeStateAsUpdate(entry.doc);

      await new Promise<void>((resolve, reject) => {
        try {
          const tx = (db as any).transaction([UPDATES_STORE_NAME], "readwrite");
          const finishError = () => reject((tx as any).error ?? new Error("IndexedDB flush transaction failed"));
          const finishOk = () => resolve();

          // Prefer EventTarget listeners when available.
          if (typeof tx?.addEventListener === "function") {
            tx.addEventListener("complete", finishOk, { once: true });
            tx.addEventListener("error", finishError, { once: true });
            tx.addEventListener("abort", finishError, { once: true });
          } else {
            (tx as any).oncomplete = finishOk;
            (tx as any).onerror = finishError;
            (tx as any).onabort = finishError;
          }

          const store = (tx as any).objectStore(UPDATES_STORE_NAME);
          store.add(snapshot);
        } catch (err) {
          reject(err);
        }
      });
    });
  }

  /**
   * Compact the IndexedDB update log for `docId` by rewriting it to a single snapshot update.
   *
   * This is best-effort and should be treated as a performance optimization (it reduces load/replay cost).
   */
  async compact(docId: string): Promise<void> {
    const entry = this.entries.get(docId);
    if (!entry) return;

    // If a compaction was scheduled, cancel the timer: this call supersedes it.
    const timer = this.compactTimers.get(docId);
    if (timer) clearTimeout(timer);
    this.compactTimers.delete(docId);

    await this.enqueue(docId, async () => {
      const ready = await Promise.race([
        Promise.resolve(entry.persistence.whenSynced).then(() => "synced" as const),
        entry.destroyed.then(() => "destroyed" as const),
      ]);
      if (ready === "destroyed") return;

      const db = entry.persistence.db;
      if (!db) return;

      const snapshot = Y.encodeStateAsUpdate(entry.doc);

      await new Promise<void>((resolve, reject) => {
        try {
          const tx = (db as any).transaction([UPDATES_STORE_NAME], "readwrite");
          const finishError = () => reject((tx as any).error ?? new Error("IndexedDB compaction transaction failed"));
          const finishOk = () => resolve();

          if (typeof tx?.addEventListener === "function") {
            tx.addEventListener("complete", finishOk, { once: true });
            tx.addEventListener("error", finishError, { once: true });
            tx.addEventListener("abort", finishError, { once: true });
          } else {
            (tx as any).oncomplete = finishOk;
            (tx as any).onerror = finishError;
            (tx as any).onabort = finishError;
          }

          const store = (tx as any).objectStore(UPDATES_STORE_NAME);
          store.clear();
          store.add(snapshot);
        } catch (err) {
          reject(err);
        }
      });

      this.updateCounts.set(docId, 0);
    });
  }

  async clear(docId: string): Promise<void> {
    const entry = this.entries.get(docId);
    if (entry) {
      entry.resolveDestroyed();
      await entry.persistence.clearData();
      this.destroyEntry(docId, entry);
      return;
    }

    const tmpDoc = new Y.Doc();
    const tmp = new IndexeddbPersistence(docId, tmpDoc);
    await tmp.clearData();
    tmp.destroy();
    tmpDoc.destroy();
  }
}
