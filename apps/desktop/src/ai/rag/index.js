import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import {
  HashEmbedder,
  IndexedDBBinaryStorage,
  ChunkedLocalStorageBinaryStorage,
  SqliteVectorStore,
  indexWorkbook,
} from "../../../../../packages/ai-rag/src/index.js";

/**
 * Desktop-oriented wiring for workbook RAG.
 *
 * Tauri webviews do not expose Node filesystem APIs, so persistence defaults to
 * browser storage (stable per-workbook key).
 */
function defaultSqliteStorage(workbookId) {
  const namespace = "formula.desktop.rag.sqlite";
  const idb = globalThis?.indexedDB;
  const hasIndexedDB = idb && typeof idb.open === "function";

  const bytesEqual = (a, b) => {
    if (a.byteLength !== b.byteLength) return false;
    for (let i = 0; i < a.byteLength; i += 1) {
      if (a[i] !== b[i]) return false;
    }
    return true;
  };

  const isStorage = (value) =>
    value && typeof value.getItem === "function" && typeof value.setItem === "function" && typeof value.removeItem === "function";

  const localStorageHasKey = (key) => {
    try {
      const storage = globalThis?.localStorage;
      if (!isStorage(storage)) return false;
      return storage.getItem(key) != null;
    } catch {
      return false;
    }
  };

  // Prefer IndexedDB for large SQLite exports (binary storage + higher quotas).
  if (hasIndexedDB) {
    const primary = new IndexedDBBinaryStorage({ namespace, workbookId, dbName: "formula.desktop.rag.sqlite" });
    // Backwards compatibility:
    // - Older desktop builds used LocalStorageBinaryStorage (single-key base64).
    // - Newer builds can use ChunkedLocalStorageBinaryStorage (multi-key base64).
    // ChunkedLocalStorageBinaryStorage also knows how to load legacy single-key
    // values and migrate them into chunks.
    const fallback = new ChunkedLocalStorageBinaryStorage({ namespace, workbookId });

    let fallbackCleared = false;
    let primaryBroken = false;

    const clearFallbackOnce = async () => {
      if (fallbackCleared) return;
      fallbackCleared = true;
      try {
        await fallback.remove?.();
      } catch {
        // ignore
      }
    };

    const saveToPrimaryAndVerify = async (data) => {
      try {
        await primary.save(data);
        const check = await primary.load();
        return check != null && bytesEqual(check, data);
      } catch {
        return false;
      }
    };

    return {
      async load() {
        // If localStorage contains any bytes, prefer them. This avoids a scenario
        // where IndexedDB writes previously failed (so localStorage has the latest
        // DB), but IndexedDB still contains an older copy.
        const key = `${namespace}:${workbookId}`;
        const hasLocal = localStorageHasKey(`${key}:meta`) || localStorageHasKey(key);

        const idbBytes = await primary.load();
        if (!hasLocal) return idbBytes;

        const localBytes = await fallback.load();
        if (!localBytes) return idbBytes;

        // If both storages exist and differ, treat localStorage as authoritative
        // (it's only written when IndexedDB persistence is not working).
        if (idbBytes && bytesEqual(idbBytes, localBytes)) {
          // Local bytes are redundant; clean them up.
          await clearFallbackOnce();
          return idbBytes;
        }

        // Best-effort migration into IndexedDB for future sessions.
        if (!primaryBroken) {
          const ok = await saveToPrimaryAndVerify(localBytes);
          if (ok) {
            await clearFallbackOnce();
          } else {
            primaryBroken = true;
          }
        }

        return localBytes;
      },
      async save(data) {
        if (!primaryBroken) {
          const ok = await saveToPrimaryAndVerify(data);
          if (ok) {
            await clearFallbackOnce();
            return;
          }
          primaryBroken = true;
        }

        // Fallback persistence when IndexedDB is unavailable/blocked.
        await fallback.save(data);
      },
      async remove() {
        await primary.remove?.();
        await fallback.remove?.();
      },
    };
  }

  // Fallback for restricted environments that disable IndexedDB.
  return new ChunkedLocalStorageBinaryStorage({ namespace, workbookId });
}

export async function createDesktopRagSqlite(opts) {
  const workbookId = opts.workbookId;
  const dimension = opts.dimension ?? 384;
  const storage = opts.storage ?? defaultSqliteStorage(workbookId);

  const vectorStore = await SqliteVectorStore.create({
    storage,
    dimension,
    autoSave: true,
    locateFile: opts.locateFile,
  });
  // Desktop workbook RAG uses deterministic, offline hash embeddings by default.
  // This avoids user API keys, local model setup, and third-party embedding providers.
  const embedder = opts.embedder ?? new HashEmbedder({ dimension });

  const contextManager = new ContextManager({
    tokenBudgetTokens: opts.tokenBudgetTokens ?? 16_000,
    tokenEstimator: opts.tokenEstimator,
    workbookRag: {
      vectorStore,
      embedder,
      topK: opts.topK ?? 8,
      sampleRows: opts.sampleRows ?? 5,
    },
  });

  return {
    vectorStore,
    embedder,
    contextManager,
    indexWorkbook: (workbook, params) => indexWorkbook({ workbook, vectorStore, embedder, ...params }),
  };
}

export async function createDesktopRag(opts) {
  return createDesktopRagSqlite(opts);
}
