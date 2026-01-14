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
  locateFile?: (file: string, prefix?: string) => string;
}): Promise<SqliteVectorStore>;

