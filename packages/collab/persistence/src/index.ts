import type * as Y from "yjs";

export interface CollabPersistenceBinding {
  destroy(): Promise<void>;
}

export type CollabPersistenceFlushOptions = {
  /**
   * When true (default for IndexedDB persistence), `flush()` may also perform compaction
   * to keep persisted logs bounded.
   *
   * Implementations that do not support compaction may ignore this flag.
   */
  compact?: boolean;
};

/**
 * Local persistence for a Yjs document.
 *
 * Implementations should:
 * - Apply any stored CRDT updates into the provided `Y.Doc` during `load()`.
 * - Persist subsequent updates during `bind()`.
 *
 * `docId` must be stable across app restarts.
 */
export interface CollabPersistence {
  /**
   * Apply any persisted state into `doc`.
   *
   * Implementations should not clear existing document state; applying updates
   * must merge with any in-memory edits.
   */
  load(docId: string, doc: Y.Doc): Promise<void>;
  /**
   * Begin persisting subsequent updates for `doc`.
   */
  bind(docId: string, doc: Y.Doc): CollabPersistenceBinding;
  /**
   * Wait until any pending persistence work for `docId` is durably written.
   */
  flush?(docId: string, opts?: CollabPersistenceFlushOptions): Promise<void>;
  /**
   * Best-effort compaction for `docId`.
   *
   * Some persistence backends store an append-only sequence of incremental CRDT updates.
   * Over time, replaying that log can become slow and can consume significant disk space.
   *
   * Compaction rewrites persisted state into a smaller representation (typically a single
   * full-document snapshot update), reducing startup/replay cost.
   */
  compact?(docId: string): Promise<void>;
  /**
   * Remove any persisted state for `docId`.
   */
  clear?(docId: string): Promise<void>;
}
