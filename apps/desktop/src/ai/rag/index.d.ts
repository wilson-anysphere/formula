import type { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";

export type DesktopRag = {
  vectorStore: any;
  /**
   * Workbook embedder.
   *
   * Note: In Formula, desktop workbook RAG uses deterministic, offline hash
   * embeddings (`HashEmbedder`) by default (not user-configurable). A future
   * Cursor-managed embedding service can replace this to improve retrieval
   * quality.
   */
  embedder: any;
  contextManager: ContextManager;
  indexWorkbook(workbook: any, params?: any): Promise<any>;
};

/**
 * Desktop-oriented wiring for workbook RAG.
 *
 * Persistence is sqlite-backed and stored via LocalStorage. Embeddings default
 * to deterministic, offline hash embeddings (`HashEmbedder`).
 */
export function createDesktopRagSqlite(opts: any): Promise<DesktopRag>;
export function createDesktopRag(opts: any): Promise<DesktopRag>;
