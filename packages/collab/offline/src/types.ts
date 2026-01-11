import type * as Y from "yjs";

export type OfflinePersistenceMode = "indexeddb" | "file";

export type OfflinePersistenceHandle = {
  /**
   * Resolves once persisted state has been loaded into the provided Y.Doc.
   *
   * This is safe to await multiple times.
   */
  whenLoaded: () => Promise<void>;
  /**
   * Stop persisting updates and release any resources (e.g. close IndexedDB connection,
   * close file handles, remove event listeners).
   */
  destroy: () => void;
  /**
   * Clear persisted state for this document key/path.
   *
   * This does not modify the in-memory Y.Doc; it only wipes the offline storage.
   *
   * Note: clearing typically detaches persistence. Re-attach a new persistence
   * instance if you want to resume persisting after clearing.
   */
  clear: () => Promise<void>;
};

export type OfflinePersistenceOptions = {
  mode: OfflinePersistenceMode;
  /**
   * Persistence key used by IndexedDB backends (and as a general identifier).
   * Defaults to `doc.guid`.
   */
  key?: string;
  /**
   * Absolute path for file-based persistence (Node only).
   */
  filePath?: string;
  /**
   * When false, persistence is only started when `whenLoaded()` is first called.
   * Defaults to true.
   */
  autoLoad?: boolean;
};

export type OfflinePersistenceAttach = (doc: Y.Doc, opts: OfflinePersistenceOptions) => OfflinePersistenceHandle;
