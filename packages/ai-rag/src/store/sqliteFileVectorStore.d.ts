import type { SqliteVectorStore } from "./sqliteVectorStore.js";

/**
 * Node-only convenience wrapper for file-backed persistence.
 *
 * This is intentionally kept out of the browser-safe entrypoint; Tauri webviews
 * and browsers should pass a non-filesystem BinaryStorage implementation to
 * {@link SqliteVectorStore.create}.
 */
export function createSqliteFileVectorStore(opts: {
  filePath: string;
  dimension: number;
  autoSave?: boolean;
  /**
   * When true (default), failures to load/initialize an existing persisted DB
   * (e.g. corrupted bytes) will cause the store to clear the persisted payload
   * and create a fresh empty DB so callers can re-index.
   *
   * Set to false to preserve the historical behaviour (throw on corruption).
   */
  resetOnCorrupt?: boolean;
  /**
   * When true (default), if a persisted DB exists with a different embedding
   * dimension than requested, the store will wipe the persisted bytes and create
   * a fresh empty DB so callers can re-index.
   *
   * Set to false to preserve the historical behaviour (throw on mismatch).
   */
  resetOnDimensionMismatch?: boolean;
  locateFile?: (file: string, prefix?: string) => string;
}): Promise<SqliteVectorStore>;
