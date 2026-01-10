import os from "node:os";
import path from "node:path";

import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import {
  HashEmbedder,
  JsonFileVectorStore,
  SqliteVectorStore,
  indexWorkbook,
} from "../../../../../packages/ai-rag/src/index.js";

/**
 * Desktop-oriented wiring for workbook RAG.
 *
 * The real desktop app can pass a per-workbook storage directory; by default we
 * keep a stable on-disk index so chat can retrieve context without re-sending
 * entire sheets.
 */
function defaultJsonStorePath(workbookId) {
  // In a real desktop application we'd use the platform-specific app data dir.
  // For this baseline, keep it deterministic and outside the repo.
  return path.join(os.homedir(), ".formula", "rag", `${workbookId}.vectors.json`);
}

export function createDesktopRag(opts) {
  const workbookId = opts.workbookId;
  const dimension = opts.dimension ?? 384;
  const storePath = opts.storePath ?? defaultJsonStorePath(workbookId);

  const vectorStore = new JsonFileVectorStore({ filePath: storePath, dimension });
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

function defaultSqliteStorePath(workbookId) {
  return path.join(os.homedir(), ".formula", "rag", `${workbookId}.vectors.sqlite`);
}

export async function createDesktopRagSqlite(opts) {
  const workbookId = opts.workbookId;
  const dimension = opts.dimension ?? 384;
  const storePath = opts.storePath ?? defaultSqliteStorePath(workbookId);

  const vectorStore = await SqliteVectorStore.create({ filePath: storePath, dimension, autoSave: true });
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
