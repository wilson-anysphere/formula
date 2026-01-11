import { CacheManager } from "../../../../packages/power-query/src/cache/cache.js";
import { hashValue } from "../../../../packages/power-query/src/cache/key.js";
import { IndexedDBCacheStore } from "../../../../packages/power-query/src/cache/indexeddb.js";
import { MemoryCacheStore } from "../../../../packages/power-query/src/cache/memory.js";
import { HttpConnector } from "../../../../packages/power-query/src/connectors/http.js";
import { QueryEngine } from "../../../../packages/power-query/src/engine.js";
import type { OAuth2Manager } from "../../../../packages/power-query/src/oauth2/manager.js";

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
    /**
     * Optional stat adapter used for cache validation (mtime-based).
     */
    stat?: (path: string) => Promise<{ mtimeMs: number }>;
  };
  /**
   * Overrides for HTTP requests. Defaults to the global `fetch`.
   */
  fetch?: typeof fetch;
  /**
   * Optional OAuth2 manager. When supplied, HTTP connectors can transparently
   * add bearer tokens for requests that specify `auth: { type: "oauth2", ... }`
   * (or when credential hooks return `{ oauth2: ... }`).
   */
  oauth2Manager?: OAuth2Manager;
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

type TauriInvoke = (cmd: string, args?: Record<string, unknown>) => Promise<unknown>;

function getTauriInvoke(): TauriInvoke {
  const invoke = (globalThis as any).__TAURI__?.core?.invoke as TauriInvoke | undefined;
  if (!invoke) {
    throw new Error("Tauri invoke API not available");
  }
  return invoke;
}

function getTauriFs(): any {
  const tauri = (globalThis as any).__TAURI__;
  return tauri?.fs ?? tauri?.plugin?.fs ?? null;
}

function normalizeBinaryPayload(payload: unknown): Uint8Array {
  if (typeof payload === "string") {
    if (typeof Buffer !== "undefined") {
      // Node (and some bundlers) provide Buffer.
      // eslint-disable-next-line no-undef
      return new Uint8Array(Buffer.from(payload, "base64"));
    }
    if (typeof atob === "function") {
      const binary = atob(payload);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      return bytes;
    }
    throw new Error("Base64 decoding is not available in this environment");
  }
  if (payload instanceof Uint8Array) return payload;
  // Some APIs return plain number arrays.
  if (Array.isArray(payload)) return new Uint8Array(payload);
  // Node Buffer (Uint8Array subclass) or ArrayBuffer.
  if (payload && typeof (payload as any).byteLength === "number") {
    return payload instanceof ArrayBuffer ? new Uint8Array(payload) : new Uint8Array(payload as any);
  }
  throw new Error("Unexpected binary payload returned from filesystem API");
}

function normalizeMtimeMs(payload: unknown): number {
  if (payload == null) {
    throw new Error("Unexpected stat payload returned from filesystem API");
  }

  if (payload instanceof Date) {
    const ms = payload.getTime();
    if (!Number.isNaN(ms)) return ms;
    throw new Error("Unexpected Date payload returned from filesystem API");
  }

  if (typeof payload === "number") {
    if (!Number.isFinite(payload)) throw new Error("Unexpected numeric mtime returned from filesystem API");
    // Heuristic: treat small values as seconds-since-epoch.
    return payload > 0 && payload < 100_000_000_000 ? payload * 1000 : payload;
  }

  if (typeof payload === "string") {
    const numeric = Number(payload);
    if (Number.isFinite(numeric)) {
      return numeric > 0 && numeric < 100_000_000_000 ? numeric * 1000 : numeric;
    }
    const parsed = new Date(payload);
    const ms = parsed.getTime();
    if (!Number.isNaN(ms)) return ms;
    throw new Error("Unexpected string mtime returned from filesystem API");
  }

  if (payload && typeof payload === "object") {
    const obj = payload as any;

    // Common shapes:
    // - Tauri invoke: { mtimeMs: number }
    // - Node fs.Stats: { mtimeMs: number, ... } or { mtime: Date }
    // - Rust/SystemTime serialization: { secs, nanos }
    if (typeof obj.secs === "number" && Number.isFinite(obj.secs)) {
      const nanos = typeof obj.nanos === "number" && Number.isFinite(obj.nanos) ? obj.nanos : 0;
      return obj.secs * 1000 + Math.floor(nanos / 1_000_000);
    }

    const candidate =
      obj.mtimeMs ??
      obj.mtime_ms ??
      obj.mtime ??
      obj.modifiedAtMs ??
      obj.modifiedAt ??
      obj.modified ??
      obj.lastModified ??
      null;
    if (candidate != null) return normalizeMtimeMs(candidate);
  }

  throw new Error("Unexpected stat payload returned from filesystem API");
}

function createDefaultFileAdapter(): DesktopQueryEngineOptions["fileAdapter"] {
  const fs = getTauriFs();
  const readTextFile = fs?.readTextFile;
  const readFile = fs?.readFile ?? fs?.readBinaryFile;
  const statFile = fs?.stat ?? fs?.metadata;

  if (typeof readTextFile === "function" && typeof readFile === "function") {
    return {
      readText: async (path) => readTextFile(path),
      readBinary: async (path) => normalizeBinaryPayload(await readFile(path)),
      stat:
        typeof statFile === "function"
          ? async (path) => ({ mtimeMs: normalizeMtimeMs(await statFile(path)) })
          : undefined,
    };
  }

  // The desktop app does not currently ship with the official Tauri FS plugin enabled.
  // Use our own invoke commands as a fallback.
  const invoke = getTauriInvoke();
  return {
    readText: async (path) => String(await invoke("read_text_file", { path })),
    readBinary: async (path) => normalizeBinaryPayload(await invoke("read_binary_file", { path })),
    stat: async (path) => ({ mtimeMs: normalizeMtimeMs(await invoke("stat_file", { path })) }),
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

  const http =
    options.fetch || options.oauth2Manager
      ? new HttpConnector({ fetch: options.fetch, oauth2Manager: options.oauth2Manager })
      : undefined;

  // Cache permission prompts across executions so previewing the same query
  // doesn't repeatedly ask the user.
  const permissionPromptCache = new Map<string, Promise<boolean>>();

  return new QueryEngine({
    cache,
    defaultCacheTtlMs: options.defaultCacheTtlMs,
    fileAdapter: {
      readText: fileAdapter.readText,
      readBinary: fileAdapter.readBinary,
      stat: fileAdapter.stat,
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

      const cacheKey = `${kind}:${hashValue(details)}`;
      const existing = permissionPromptCache.get(cacheKey);
      if (existing) return await existing;

      const decisionPromise = Promise.resolve().then(async () => {
        if (options.onPermissionPrompt) {
          return await options.onPermissionPrompt(kind, details);
        }
        return defaultPermissionPrompt(kind, details);
      });
      permissionPromptCache.set(cacheKey, decisionPromise);
      return await decisionPromise;
    },
  });
}
