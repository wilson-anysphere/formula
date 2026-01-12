import type { DocumentController } from "../../document/documentController.js";

import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import {
  HashEmbedder,
  LocalStorageBinaryStorage,
  workbookFromSpreadsheetApi,
} from "../../../../../packages/ai-rag/src/index.js";

import { createDesktopRag } from "./index.js";

export type DesktopRagEmbedderConfig = {
  /**
   * Hash embeddings are deterministic and offline. Dimension controls the vector
   * size stored in SQLite (higher = more storage, marginally better recall).
   *
   * Note: Embeddings are intentionally not user-configurable in Formula; this is
   * an internal/testing knob only.
   */
  dimension?: number;

  // Kept for backwards compatibility with earlier config shapes.
  type?: "hash";
};

export interface DesktopRagServiceOptions {
  documentController: DocumentController;
  workbookId: string;

  /**
   * ContextManager configuration.
   */
  tokenBudgetTokens?: number;
  topK?: number;
  sampleRows?: number;

  /**
   * Embedding configuration. Desktop workbook RAG uses `HashEmbedder`
   * (deterministic hash embeddings) by default (offline; no API keys / local models).
   */
  embedder?: DesktopRagEmbedderConfig;

  /**
   * Override the sqlite BinaryStorage namespace (advanced / tests).
   *
   * Note: the default namespace is stable per workbook **and** embedding
   * dimension, so changing the dimension won't brick the store due to mismatches.
   */
  storageNamespace?: string;

  /**
   * sql.js locateFile hook (useful in non-standard bundling environments).
   */
  locateFile?: (file: string, prefix?: string) => string;

  /**
   * Test seam: override the underlying createDesktopRag factory.
   */
  createRag?: typeof createDesktopRag;
}

type DesktopRag = Awaited<ReturnType<typeof createDesktopRag>>;

export interface DesktopRagService {
  getContextManager(): Promise<ContextManager>;
  buildWorkbookContextFromSpreadsheetApi(params: {
    spreadsheet: any;
    workbookId: string;
    query: string;
    attachments?: any[];
    topK?: number;
    dlp?: any;
  }): Promise<any>;
  /**
   * Remove DocumentController listeners and close the underlying sqlite store
   * (if it was ever initialized).
   */
  dispose(): Promise<void>;
}

function embedderIdentity(dimension: number): string {
  return dimension === 384 ? "hash" : `hash-${dimension}`;
}

function storageNamespaceForEmbedder(params: {
  baseNamespace: string;
  dimension: number;
}): string {
  // Preserve the legacy namespace for the default (hash, 384) embedder.
  const id = embedderIdentity(params.dimension);
  if (id === "hash") return params.baseNamespace;
  return `${params.baseNamespace}:${id}`;
}

function resolveEmbedder(
  config: DesktopRagEmbedderConfig | undefined,
): { embedder: HashEmbedder; dimension: number } {
  // Ignore unknown/legacy config types (Formula only supports deterministic hash embeddings).
  const dimension = config?.type && config.type !== "hash" ? 384 : (config?.dimension ?? 384);
  return { embedder: new HashEmbedder({ dimension }), dimension };
}

/**
 * Desktop RAG service:
 * - Uses sqlite-backed vector store persisted in LocalStorage.
 * - Tracks DocumentController mutations and only re-indexes when content changes.
 * - Keeps buildWorkbookContextFromSpreadsheetApi cheap (no workbook scan when index is up to date).
 * - Uses deterministic hash embeddings by design (offline; no user API keys / local models).
 */
export function createDesktopRagService(options: DesktopRagServiceOptions): DesktopRagService {
  const ragFactory = options.createRag ?? createDesktopRag;

  const { embedder, dimension } = resolveEmbedder(options.embedder);

  const storageNamespace = storageNamespaceForEmbedder({
    baseNamespace: options.storageNamespace ?? "formula.desktop.rag.sqlite",
    dimension,
  });

  let ragPromise: Promise<DesktopRag> | null = null;
  let disposed = false;

  let documentVersion = 0;
  let indexedVersion: number | null = null;
  let indexPromise: Promise<unknown> | null = null;
  let lastIndexStats: unknown = null;

  let suppressNextUpdate = false;

  const onChange = () => {
    documentVersion += 1;
    suppressNextUpdate = true;
    queueMicrotask(() => {
      suppressNextUpdate = false;
    });
  };

  const onUpdate = () => {
    if (suppressNextUpdate) return;
    documentVersion += 1;
  };

  const unsubscribeChange = options.documentController.on?.("change", onChange) ?? (() => {});
  const unsubscribeUpdate = options.documentController.on?.("update", onUpdate) ?? (() => {});

  async function getRag(): Promise<DesktopRag> {
    if (disposed) throw new Error("DesktopRagService is disposed");
    if (!ragPromise) {
      ragPromise = ragFactory({
        workbookId: options.workbookId,
        dimension,
        embedder,
        storage: new LocalStorageBinaryStorage({ namespace: storageNamespace, workbookId: options.workbookId }),
        tokenBudgetTokens: options.tokenBudgetTokens,
        topK: options.topK,
        sampleRows: options.sampleRows,
        locateFile: options.locateFile,
      } as any);
    }
    return ragPromise;
  }

  async function ensureIndexed(spreadsheet: any): Promise<void> {
    if (disposed) throw new Error("DesktopRagService is disposed");

    // Avoid concurrent re-indexing (multiple chat messages, tool loops, etc).
    if (indexPromise) await indexPromise;

    const currentVersion = documentVersion;
    if (indexedVersion === currentVersion) return;

    const run = (async () => {
      const rag = await getRag();
      const versionToIndex = documentVersion;
      const workbook = workbookFromSpreadsheetApi({
        spreadsheet,
        workbookId: options.workbookId,
        coordinateBase: "one",
      });

      lastIndexStats = await rag.indexWorkbook(workbook, { sampleRows: options.sampleRows } as any);
      indexedVersion = versionToIndex;
    })();

    indexPromise = run;
    try {
      await run;
    } finally {
      if (indexPromise === run) indexPromise = null;
    }
  }

  async function getContextManager(): Promise<ContextManager> {
    const rag = await getRag();
    return rag.contextManager as ContextManager;
  }

  async function buildWorkbookContextFromSpreadsheetApi(params: {
    spreadsheet: any;
    workbookId: string;
    query: string;
    attachments?: any[];
    topK?: number;
    dlp?: any;
  }): Promise<any> {
    if (params.workbookId !== options.workbookId) {
      throw new Error(
        `DesktopRagService workbookId mismatch: expected "${options.workbookId}", got "${params.workbookId}"`,
      );
    }

    const rag = await getRag();

    // If DLP is enabled, ContextManager forces indexing so it can apply redaction
    // before embedding/persisting content. In that mode we don't try to manage
    // indexing externally (avoid double-indexing).
    const canSkipIndexing = !params.dlp;
    if (canSkipIndexing) await ensureIndexed(params.spreadsheet);

    const ctx = await rag.contextManager.buildWorkbookContextFromSpreadsheetApi({
      ...params,
      // When the index is already up to date, avoid re-scanning the workbook for cells.
      skipIndexing: canSkipIndexing,
    } as any);

    // Preserve last index stats for callers that want to surface them even when
    // buildWorkbookContextFromSpreadsheetApi is in "cheap" mode.
    if (ctx && ctx.indexStats == null) ctx.indexStats = lastIndexStats;
    return ctx;
  }

  async function dispose(): Promise<void> {
    if (disposed) return;
    disposed = true;
    try {
      unsubscribeChange();
    } catch {
      // ignore
    }
    try {
      unsubscribeUpdate();
    } catch {
      // ignore
    }

    const rag = await ragPromise?.catch(() => null);
    try {
      await rag?.vectorStore?.close?.();
    } catch {
      // ignore
    }
  }

  return {
    getContextManager,
    buildWorkbookContextFromSpreadsheetApi,
    dispose,
  };
}
