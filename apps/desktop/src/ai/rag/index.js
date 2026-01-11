import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import {
  HashEmbedder,
  LocalStorageBinaryStorage,
  SqliteVectorStore,
  indexWorkbook,
} from "../../../../../packages/ai-rag/src/index.js";

/**
 * Desktop-oriented wiring for workbook RAG.
 *
 * Tauri webviews do not expose Node filesystem APIs, so persistence defaults to
 * LocalStorage (stable per-workbook key).
 */
function defaultSqliteStorage(workbookId) {
  return new LocalStorageBinaryStorage({ namespace: "formula.desktop.rag.sqlite", workbookId });
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
  const embedder = opts.embedder ?? new HashEmbedder({ dimension });

  const contextManager = new ContextManager({
    tokenBudgetTokens: opts.tokenBudgetTokens ?? 16_000,
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
