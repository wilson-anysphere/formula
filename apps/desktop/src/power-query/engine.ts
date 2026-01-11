import { CacheManager } from "../../../../packages/power-query/src/cache/cache.js";
import { IndexedDBCacheStore } from "../../../../packages/power-query/src/cache/indexeddb.js";
import { MemoryCacheStore } from "../../../../packages/power-query/src/cache/memory.js";
import { HttpConnector } from "../../../../packages/power-query/src/connectors/http.js";
import { QueryEngine } from "../../../../packages/power-query/src/engine.js";

import { enforceExternalConnector } from "../dlp/enforceExternalConnector.js";
import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";

type DlpContext = {
  documentId: string;
  sheetId?: string;
  range?: unknown;
  classificationStore: { list: (documentId: string) => Array<{ selector: unknown; classification: unknown }> };
  policy: unknown;
};

export type DesktopQueryEngineOptions = {
  /**
   * Optional DLP context. When present, connector permission requests will be
   * checked against the organization policy before running.
   */
  dlp?: DlpContext;
  /**
   * Optional user prompt for permission requests. Returning `false` denies the request.
   * If omitted, permissions are allowed (after DLP enforcement).
   */
  onPermissionPrompt?: (kind: string, details: unknown) => boolean | Promise<boolean>;
  onCredentialRequest?: (connectorId: string, details: unknown) => unknown | Promise<unknown>;
  /**
   * Overrides for file IO. By default we use the Tauri filesystem API.
   */
  fileAdapter?: {
    readText: (path: string) => Promise<string>;
    readBinary: (path: string) => Promise<Uint8Array>;
  };
  /**
   * Overrides for HTTP requests. Defaults to the global `fetch`.
   */
  fetch?: typeof fetch;
  /**
   * Cache manager override. If omitted, Formula uses IndexedDB in browser contexts
   * (and an in-memory cache as a fallback for non-browser environments).
   */
  cache?: CacheManager;
  defaultCacheTtlMs?: number;
};

const PERMISSION_KIND_TO_DLP_ACTION: Record<string, string> = {
  "file:read": DLP_ACTION.EXTERNAL_CONNECTOR,
  "http:request": DLP_ACTION.EXTERNAL_CONNECTOR,
  "database:query": DLP_ACTION.EXTERNAL_CONNECTOR,
};

function getTauriFs(): any {
  const tauri = (globalThis as any).__TAURI__;
  return tauri?.fs ?? tauri?.plugin?.fs ?? null;
}

function normalizeBinaryPayload(payload: unknown): Uint8Array {
  if (payload instanceof Uint8Array) return payload;
  // Some APIs return plain number arrays.
  if (Array.isArray(payload)) return new Uint8Array(payload);
  // Node Buffer (Uint8Array subclass) or ArrayBuffer.
  if (payload && typeof (payload as any).byteLength === "number") {
    return payload instanceof ArrayBuffer ? new Uint8Array(payload) : new Uint8Array(payload as any);
  }
  throw new Error("Unexpected binary payload returned from filesystem API");
}

function createDefaultFileAdapter(): DesktopQueryEngineOptions["fileAdapter"] {
  const fs = getTauriFs();
  const readTextFile = fs?.readTextFile;
  const readFile = fs?.readFile ?? fs?.readBinaryFile;

  if (typeof readTextFile !== "function" || typeof readFile !== "function") {
    throw new Error("Tauri filesystem API not available (missing readTextFile/readFile)");
  }

  return {
    readText: async (path) => readTextFile(path),
    readBinary: async (path) => normalizeBinaryPayload(await readFile(path)),
  };
}

function createDefaultCacheManager(): CacheManager {
  const store =
    typeof indexedDB !== "undefined" ? new IndexedDBCacheStore({ dbName: "formula-power-query-cache" }) : new MemoryCacheStore();
  return new CacheManager({ store });
}

function defaultPermissionPrompt(kind: string, details: unknown): boolean {
  if (typeof window === "undefined" || typeof window.confirm !== "function") return true;

  const request = (details as any)?.request;
  if (kind === "http:request") {
    const url = typeof request?.url === "string" ? request.url : "an external URL";
    return window.confirm(`Allow this query to make a network request to ${url}?`);
  }
  if (kind === "file:read") {
    const path = typeof request?.path === "string" ? request.path : "a local file";
    return window.confirm(`Allow this query to read ${path}?`);
  }
  if (kind === "database:query") {
    return window.confirm("Allow this query to run a database query?");
  }
  return window.confirm(`Allow this query to access: ${kind}?`);
}

/**
 * Create a QueryEngine configured for the desktop app runtime:
 * - file reads via the Tauri FS bridge
 * - HTTP via `fetch`
 * - IndexedDB-backed cache
 * - optional DLP/permission prompting via `onPermissionRequest`
 */
export function createDesktopQueryEngine(options: DesktopQueryEngineOptions = {}): QueryEngine {
  const cache = options.cache ?? createDefaultCacheManager();
  const fileAdapter = options.fileAdapter ?? createDefaultFileAdapter();

  const http = options.fetch ? new HttpConnector({ fetch: options.fetch }) : undefined;

  return new QueryEngine({
    cache,
    defaultCacheTtlMs: options.defaultCacheTtlMs,
    fileAdapter: {
      readText: fileAdapter.readText,
      readBinary: fileAdapter.readBinary,
    },
    connectors: http ? { http } : undefined,
    onCredentialRequest: options.onCredentialRequest,
    onPermissionRequest: async (kind, details) => {
      const dlpAction = PERMISSION_KIND_TO_DLP_ACTION[kind];
      if (dlpAction === DLP_ACTION.EXTERNAL_CONNECTOR && options.dlp) {
        enforceExternalConnector({
          documentId: options.dlp.documentId,
          sheetId: options.dlp.sheetId,
          range: options.dlp.range,
          classificationStore: options.dlp.classificationStore,
          policy: options.dlp.policy,
        });
      }

      if (options.onPermissionPrompt) {
        return await options.onPermissionPrompt(kind, details);
      }

      return defaultPermissionPrompt(kind, details);
    },
  });
}
