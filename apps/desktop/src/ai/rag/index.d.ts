import type { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";

export type DesktopRag = {
  vectorStore: any;
  embedder: any;
  contextManager: ContextManager;
  indexWorkbook(workbook: any, params?: any): Promise<any>;
};

export function createDesktopRagSqlite(opts: any): Promise<DesktopRag>;
export function createDesktopRag(opts: any): Promise<DesktopRag>;

