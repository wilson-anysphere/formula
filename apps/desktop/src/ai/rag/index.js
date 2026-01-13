import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import {
  HashEmbedder,
  IndexedDBBinaryStorage,
  ChunkedLocalStorageBinaryStorage,
  LocalStorageBinaryStorage,
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

  // Prefer IndexedDB for large SQLite exports (binary storage + higher quotas).
  if (hasIndexedDB) {
    const primary = new IndexedDBBinaryStorage({ namespace, workbookId, dbName: "formula.desktop.rag.sqlite" });
    // Backwards compatibility: older desktop builds used single-key base64 localStorage persistence.
    // Also support the newer chunked localStorage format in case hosts opt into it.
    const legacy = new LocalStorageBinaryStorage({ namespace, workbookId });
    const chunked = new ChunkedLocalStorageBinaryStorage({ namespace, workbookId });

    const bytesEqual = (a, b) => {
      if (a.byteLength !== b.byteLength) return false;
      for (let i = 0; i < a.byteLength; i += 1) {
        if (a[i] !== b[i]) return false;
      }
      return true;
    };

    const migrate = async (data, source) => {
      try {
        await primary.save(data);
        const check = await primary.load();
        if (check && bytesEqual(check, data)) {
          await source.remove?.();
        }
      } catch {
        // ignore migration failures; callers will still get the legacy data for this session
      }
    };

    return {
      async load() {
        const existing = await primary.load();
        if (existing) return existing;

        // Try legacy single-key storage first so we don't trigger chunked migration writes
        // if we immediately plan to move the bytes into IndexedDB.
        const legacyBytes = await legacy.load();
        if (legacyBytes) {
          await migrate(legacyBytes, legacy);
          return legacyBytes;
        }

        const chunkedBytes = await chunked.load();
        if (chunkedBytes) {
          await migrate(chunkedBytes, chunked);
          return chunkedBytes;
        }

        return null;
      },
      async save(data) {
        await primary.save(data);
      },
      async remove() {
        await primary.remove?.();
        await legacy.remove?.();
        await chunked.remove?.();
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
