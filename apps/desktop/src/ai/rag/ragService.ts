import type { DocumentController } from "../../document/documentController.js";

import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import { stableJsonStringify, type TokenEstimator } from "../../../../../packages/ai-context/src/tokenBudget.js";
import {
  HashEmbedder,
  ChunkedLocalStorageBinaryStorage,
  IndexedDBBinaryStorage,
  workbookFromSpreadsheetApi,
} from "../../../../../packages/ai-rag/src/index.js";
import { DlpViolationError } from "../../../../../packages/security/dlp/src/errors.js";

import { createDesktopRag } from "./index.js";
import { computeDlpCacheKey } from "../dlp/dlpCacheKey.js";

function createAbortError(message = "Aborted"): Error {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw createAbortError();
}

async function awaitWithAbort<T>(promise: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return promise;
  if (signal.aborted) throw createAbortError();

  return new Promise<T>((resolve, reject) => {
    const onAbort = () => {
      reject(createAbortError());
    };
    signal.addEventListener("abort", onAbort, { once: true });

    promise.then(
      (value) => {
        signal.removeEventListener("abort", onAbort);
        resolve(value);
      },
      (error) => {
        signal.removeEventListener("abort", onAbort);
        reject(error);
      },
    );
  });
}

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
  tokenEstimator?: TokenEstimator;

  /**
   * Embedding configuration. Desktop workbook RAG uses `HashEmbedder`
   * (deterministic hash embeddings) by default (offline; no user API keys or local model setup).
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

function sheetNamesCacheKeyFor(params: { spreadsheet: any }): string {
  const spreadsheet = params.spreadsheet;
  if (!spreadsheet || typeof spreadsheet.listSheets !== "function") return "sheets:none";
  try {
    const raw = spreadsheet.listSheets();
    const list = Array.isArray(raw) ? raw : [];
    const normalized = list
      .map((name) => String(name ?? "").trim())
      .filter((name) => name.length > 0);
    // Sheet order should not matter for indexing validity; only the set of names.
    normalized.sort((a, b) => a.localeCompare(b));
    return stableJsonStringify(normalized);
  } catch {
    // Never crash AI surfaces due to adapter errors; treat as a changing key so indexing will retry.
    return `sheets:error:${Date.now()}`;
  }
}

export interface DesktopRagService {
  getContextManager(): Promise<ContextManager>;
  buildWorkbookContextFromSpreadsheetApi(params: {
    spreadsheet: any;
    workbookId: string;
    query: string;
    attachments?: any[];
    topK?: number;
    includePromptContext?: boolean;
    signal?: AbortSignal;
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
  if (config?.type && config.type !== "hash") {
    throw new Error(
      `Desktop workbook RAG only supports deterministic hash embeddings (HashEmbedder). Received embedder.type="${config.type}".`,
    );
  }
  const dimension = config?.dimension ?? 384;
  return { embedder: new HashEmbedder({ dimension }), dimension };
}

/**
 * Desktop RAG service:
 * - Uses sqlite-backed vector store persisted in browser storage (IndexedDB when
 *   available; chunked localStorage fallback).
 * - Tracks DocumentController mutations and only re-indexes when content changes.
 * - Keeps buildWorkbookContextFromSpreadsheetApi cheap (no workbook scan when index is up to date).
 * - Uses deterministic hash embeddings by design (offline; no user API keys or local model setup).
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

  const controllerAny = options.documentController as any;
  const supportsContentVersion =
    typeof controllerAny?.contentVersion === "number" && Number.isFinite(controllerAny.contentVersion);
  const supportsUpdateVersion = typeof controllerAny?.updateVersion === "number" && Number.isFinite(controllerAny.updateVersion);

  // Fallback for older controller implementations that don't expose version counters.
  let fallbackVersion = 0;
  let indexedVersion: number | null = null;
  let indexedDlpKey: string | null = null;
  let indexedSheetNamesKey: string | null = null;
  let indexPromise: Promise<unknown> | null = null;
  let lastIndexStats: unknown = null;

  // Legacy controllers may fire both `change` and `update` for a single mutation.
  // Keep behavior aligned with the prior implementation by suppressing the next
  // `update` tick after a `change`.
  let suppressNextUpdate = false;

  function currentVersion(): number {
    const controller = options.documentController as any;
    const contentVersion = controller?.contentVersion;
    if (typeof contentVersion === "number" && Number.isFinite(contentVersion)) return contentVersion;
    const updateVersion = controller?.updateVersion;
    if (typeof updateVersion === "number" && Number.isFinite(updateVersion)) return updateVersion;
    return fallbackVersion;
  }

  const onChange = () => {
    // When the controller exposes monotonic versions, prefer those (they allow
    // us to ignore view-only changes via `contentVersion`).
    if (supportsContentVersion || supportsUpdateVersion) return;

    // Legacy fallback: bump on any change.
    fallbackVersion += 1;
    suppressNextUpdate = true;
    queueMicrotask(() => {
      suppressNextUpdate = false;
    });
  };

  const onUpdate = () => {
    if (supportsContentVersion || supportsUpdateVersion) return;
    if (suppressNextUpdate) return;

    // Legacy fallback: bump on any update. Note: change and update may both
    // fire for a single mutation depending on controller implementation.
    fallbackVersion += 1;
  };

  const shouldSubscribeForFallback = !(supportsContentVersion || supportsUpdateVersion);
  const unsubscribeChange = shouldSubscribeForFallback ? options.documentController.on?.("change", onChange) ?? (() => {}) : () => {};
  const unsubscribeUpdate = shouldSubscribeForFallback ? options.documentController.on?.("update", onUpdate) ?? (() => {}) : () => {};

  async function getRag(): Promise<DesktopRag> {
    if (disposed) throw new Error("DesktopRagService is disposed");
    if (!ragPromise) {
      const idb = (globalThis as any)?.indexedDB as IDBFactory | undefined;
      const hasIndexedDB = idb && typeof idb.open === "function";

      const buildStorage = () => {
        if (!hasIndexedDB) {
          // Restricted environments that disable IndexedDB (or older WebViews).
          // Use chunked localStorage to avoid per-key limits.
          return new ChunkedLocalStorageBinaryStorage({ namespace: storageNamespace, workbookId: options.workbookId });
        }

        const primary = new IndexedDBBinaryStorage({ namespace: storageNamespace, workbookId: options.workbookId, dbName: "formula.desktop.rag.sqlite" });
        // Backwards compatibility:
        // - Older desktop builds used LocalStorageBinaryStorage (single-key base64).
        // - Newer builds can use ChunkedLocalStorageBinaryStorage (multi-key base64).
        // ChunkedLocalStorageBinaryStorage also knows how to load legacy single-key
        // values and migrate them into chunks.
        const fallback = new ChunkedLocalStorageBinaryStorage({ namespace: storageNamespace, workbookId: options.workbookId });

        const bytesEqual = (a: Uint8Array, b: Uint8Array) => {
          if (a.byteLength !== b.byteLength) return false;
          for (let i = 0; i < a.byteLength; i += 1) {
            if (a[i] !== b[i]) return false;
          }
          return true;
        };

        const isStorage = (value: unknown): value is Storage =>
          Boolean(
            value &&
              typeof (value as any).getItem === "function" &&
              typeof (value as any).setItem === "function" &&
              typeof (value as any).removeItem === "function",
          );

        const localStorageHasKey = (key: string) => {
          try {
            const storage = (globalThis as any)?.localStorage as Storage | undefined;
            if (!isStorage(storage)) return false;
            return storage.getItem(key) != null;
          } catch {
            return false;
          }
        };

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

        const saveToPrimaryAndVerify = async (data: Uint8Array) => {
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
            const key = `${storageNamespace}:${options.workbookId}`;
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
          async save(data: Uint8Array) {
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
      };

      ragPromise = ragFactory({
        workbookId: options.workbookId,
        dimension,
        embedder,
        storage: buildStorage(),
        tokenBudgetTokens: options.tokenBudgetTokens,
        tokenEstimator: options.tokenEstimator,
        topK: options.topK,
        sampleRows: options.sampleRows,
        locateFile: options.locateFile,
      } as any);
    }
    return ragPromise;
  }

  async function ensureIndexed(spreadsheet: any, signal?: AbortSignal): Promise<void> {
    if (disposed) throw new Error("DesktopRagService is disposed");
    throwIfAborted(signal);

    // Avoid concurrent re-indexing (multiple chat messages, tool loops, etc).
    if (indexPromise) {
      try {
        await awaitWithAbort(indexPromise, signal);
      } catch (error) {
        // If a previous indexing run was canceled by a different request, allow this call
        // to retry indexing. Only propagate AbortError when *this* call is aborted.
        if (signal?.aborted) throw error;
        if ((error as any)?.name !== "AbortError") throw error;
      }
    }
    throwIfAborted(signal);

    const versionNow = currentVersion();
    const sheetNamesKeyNow = sheetNamesCacheKeyFor({ spreadsheet });
    if (indexedVersion === versionNow && indexedSheetNamesKey === sheetNamesKeyNow) return;

    const run = (async () => {
      throwIfAborted(signal);
      const rag = await awaitWithAbort(getRag(), signal);
      const versionToIndex = currentVersion();
      const sheetNamesKeyToIndex = sheetNamesCacheKeyFor({ spreadsheet });
      const workbook = workbookFromSpreadsheetApi({
        spreadsheet,
        workbookId: options.workbookId,
        coordinateBase: "one",
        signal,
      });

      lastIndexStats = await rag.indexWorkbook(workbook, { sampleRows: options.sampleRows, signal } as any);
      indexedVersion = versionToIndex;
      indexedDlpKey = null;
      indexedSheetNamesKey = sheetNamesKeyToIndex;
    })();

    indexPromise = run;
    run
      .finally(() => {
        if (indexPromise === run) indexPromise = null;
      })
      .catch(() => {});

    await awaitWithAbort(run, signal);
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
    includePromptContext?: boolean;
    signal?: AbortSignal;
    dlp?: any;
  }): Promise<any> {
    const signal = params.signal;
    throwIfAborted(signal);
    if (params.workbookId !== options.workbookId) {
      throw new Error(
        `DesktopRagService workbookId mismatch: expected "${options.workbookId}", got "${params.workbookId}"`,
      );
    }

    const rag = await awaitWithAbort(getRag(), signal);

    const hasDlp = Boolean(params.dlp);

    // Non-DLP mode: we manage incremental indexing externally and always run
    // ContextManager in "cheap" mode to avoid workbook scans.
    if (!hasDlp) {
      await ensureIndexed(params.spreadsheet, signal);
      throwIfAborted(signal);

      const ctx = await rag.contextManager.buildWorkbookContextFromSpreadsheetApi({
        ...params,
        skipIndexing: true,
      } as any);

      // Preserve last index stats for callers that want to surface them even when
      // buildWorkbookContextFromSpreadsheetApi is in "cheap" mode.
      if (ctx && ctx.indexStats == null) ctx.indexStats = lastIndexStats;
      return ctx;
    }

    // DLP mode: only skip workbook scans when both the workbook version AND the DLP inputs
    // (policy/classifications/includeRestrictedContent) match the last indexed state.
    const dlpKey = computeDlpCacheKey(params.dlp);
    const sheetNamesKeyNow = sheetNamesCacheKeyFor({ spreadsheet: params.spreadsheet });

    // Avoid concurrent re-indexing (multiple chat messages, tool loops, etc).
    if (indexPromise) {
      try {
        await awaitWithAbort(indexPromise, signal);
      } catch (error) {
        // If a previous indexing run was canceled by a different request, allow this call
        // to retry indexing. Only propagate AbortError when *this* call is aborted.
        if (signal?.aborted) throw error;
        if ((error as any)?.name !== "AbortError") throw error;
      }
    }
    throwIfAborted(signal);

    const versionNow = currentVersion();
    const shouldIndex = indexedVersion !== versionNow || indexedDlpKey !== dlpKey || indexedSheetNamesKey !== sheetNamesKeyNow;

    if (shouldIndex) {
      const run = (async () => {
        const versionToIndex = currentVersion();
        const sheetNamesKeyToIndex = sheetNamesCacheKeyFor({ spreadsheet: params.spreadsheet });
        try {
          const ctx = await rag.contextManager.buildWorkbookContextFromSpreadsheetApi({
            ...params,
            // DLP-safe full path: force a rescan/index so we can apply redaction before embedding.
            skipIndexing: false,
          } as any);
          lastIndexStats = ctx?.indexStats ?? lastIndexStats;
          indexedVersion = versionToIndex;
          indexedDlpKey = dlpKey;
          indexedSheetNamesKey = sheetNamesKeyToIndex;
          return ctx;
        } catch (error) {
          // If DLP blocks cloud AI processing, ContextManager throws after indexing so we can
          // prevent sending anything to the LLM. In that case we can still treat the index
          // as up-to-date, which avoids expensive rescans on repeated blocked requests.
          if (error instanceof DlpViolationError) {
            indexedVersion = versionToIndex;
            indexedDlpKey = dlpKey;
            indexedSheetNamesKey = sheetNamesKeyToIndex;
          }
          throw error;
        }
      })();

      indexPromise = run;
      // Ensure we don't clear `indexPromise` early if a caller aborts while waiting.
      // The indexing run should still be considered in-flight until it settles, otherwise
      // subsequent calls may kick off concurrent indexing work.
      run
        .finally(() => {
          if (indexPromise === run) indexPromise = null;
        })
        .catch(() => {});

      const ctx = await awaitWithAbort(run, signal);
      if (ctx && ctx.indexStats == null) ctx.indexStats = lastIndexStats;
      return ctx;
    }

    const ctx = await rag.contextManager.buildWorkbookContextFromSpreadsheetApi({
      ...params,
      // When the DLP index is already up to date, avoid re-scanning the workbook for cells.
      skipIndexing: true,
      skipIndexingWithDlp: true,
    } as any);

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
