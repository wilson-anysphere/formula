import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import {
  HashEmbedder,
  IndexedDBBinaryStorage,
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
  const hasIndexedDB =
    // eslint-disable-next-line no-undef
    typeof indexedDB !== "undefined" || (globalThis && "indexedDB" in globalThis && globalThis.indexedDB);

  // Prefer IndexedDB for large SQLite exports (binary storage + higher quotas).
  if (hasIndexedDB) {
    return new IndexedDBBinaryStorage({ namespace, workbookId, dbName: "formula.desktop.rag.sqlite" });
  }

  // Fallback for restricted environments that disable IndexedDB.
  return new LocalStorageBinaryStorage({ namespace, workbookId });
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
