import { NodeFileBinaryStorage } from "./nodeFileBinaryStorage.js";
import { SqliteVectorStore } from "./sqliteVectorStore.js";

/**
 * Node-only convenience wrapper for file-backed persistence.
 *
 * This is intentionally kept out of the browser-safe entrypoint; Tauri webviews
 * and browsers should pass a non-filesystem BinaryStorage implementation (e.g.
 * IndexedDBBinaryStorage / ChunkedLocalStorageBinaryStorage / LocalStorageBinaryStorage)
 * to `SqliteVectorStore.create`.
 */
export async function createSqliteFileVectorStore(opts) {
  if (!opts?.filePath) throw new Error("createSqliteFileVectorStore requires filePath");
  const storage = new NodeFileBinaryStorage(opts.filePath);
  return SqliteVectorStore.create({
    storage,
    dimension: opts.dimension,
    autoSave: opts.autoSave,
    resetOnDimensionMismatch: opts.resetOnDimensionMismatch,
    resetOnCorrupt: opts.resetOnCorrupt,
    locateFile: opts.locateFile,
  });
}
