import { CacheManager } from "../../../../packages/power-query/src/cache/cache.js";
import { hashValue } from "../../../../packages/power-query/src/cache/key.js";
import { IndexedDBCacheStore } from "../../../../packages/power-query/src/cache/indexeddb.js";
import { MemoryCacheStore } from "../../../../packages/power-query/src/cache/memory.js";
import { HttpConnector } from "../../../../packages/power-query/src/connectors/http.js";
import { QueryEngine } from "../../../../packages/power-query/src/engine.js";
import { DataTable } from "../../../../packages/power-query/src/table.js";
import { parseA1Range, splitSheetQualifier } from "../../../../packages/search/index.js";
import type { OAuth2Manager } from "../../../../packages/power-query/src/oauth2/manager.js";
import type { QueryExecutionContext } from "../../../../packages/power-query/src/engine.js";

import { enforceExternalConnector } from "../dlp/enforceExternalConnector.js";
import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { effectiveDocumentClassification, effectiveRangeClassification } from "../../../../packages/security/dlp/src/selectors.js";

import type { DocumentController } from "../document/documentController.js";
import { getTableSignatureRegistry } from "./tableSignatures.ts";

type DlpContext = {
  documentId: string;
  sheetId?: string;
  range?: unknown;
  classificationStore: { list: (documentId: string) => Array<{ selector: unknown; classification: unknown }> };
  policy: unknown | (() => unknown | Promise<unknown>);
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
    readBinaryStream?: (path: string, options?: { signal?: AbortSignal }) => AsyncIterable<Uint8Array>;
    openFile?: (path: string, options?: { signal?: AbortSignal }) => Promise<Blob>;
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
  /**
   * Power Query privacy/firewall mode. Defaults to `"ignore"` for backwards compatibility.
   */
  privacyMode?: "ignore" | "enforce" | "warn";
};

const PERMISSION_KIND_TO_DLP_ACTION: Record<string, string> = {
  "file:read": DLP_ACTION.EXTERNAL_CONNECTOR,
  "http:request": DLP_ACTION.EXTERNAL_CONNECTOR,
  "database:query": DLP_ACTION.EXTERNAL_CONNECTOR,
};

function createEphemeralObjectId(): (value: unknown) => string | null {
  const ids = new WeakMap<object, string>();
  let counter = 0;
  return (value) => {
    if (!value) return null;
    const type = typeof value;
    if (type !== "object" && type !== "function") return null;
    const obj = value as any as object;
    const existing = ids.get(obj);
    if (existing) return existing;
    const next = `obj:${++counter}`;
    ids.set(obj, next);
    return next;
  };
}

function permissionPromptCacheKey(
  kind: string,
  details: unknown,
  opts: { getObjectId: (value: unknown) => string | null },
): string | null {
  const getObjectId = opts.getObjectId;
  const record = details && typeof details === "object" && !Array.isArray(details) ? (details as any) : null;
  const request = record?.request;

  try {
    if (kind === "database:query") {
      const req = request && typeof request === "object" && !Array.isArray(request) ? (request as any) : {};
      const sql = typeof req.sql === "string" ? req.sql : "";
      const params = "params" in req ? (req.params ?? null) : null;

      const explicitConnectionId = typeof req.connectionId === "string" && req.connectionId ? req.connectionId : null;
      if (explicitConnectionId) {
        return `${kind}:${hashValue({ connector: "sql", connectionId: explicitConnectionId, sql, params })}`;
      }

      const connection = req.connection;
      if (typeof connection === "string" && connection) {
        return `${kind}:${hashValue({ connector: "sql", connectionId: connection, sql, params })}`;
      }
      if ((typeof connection === "number" && Number.isFinite(connection)) || typeof connection === "boolean") {
        return `${kind}:${hashValue({ connector: "sql", connectionId: hashValue(connection), sql, params })}`;
      }
      if (connection && typeof connection === "object" && !Array.isArray(connection)) {
        const id = (connection as any).id;
        if (typeof id === "string" && id) {
          return `${kind}:${hashValue({ connector: "sql", connectionId: id, sql, params })}`;
        }
      }

      const connectionRefId = getObjectId(connection);
      return `${kind}:${hashValue({
        connector: "sql",
        missingConnectionId: true,
        ...(connectionRefId ? { connectionRefId } : null),
        sql,
        params,
      })}`;
    }

    if (kind === "file:read") {
      const req = request && typeof request === "object" && !Array.isArray(request) ? request : {};
      return `${kind}:${hashValue({ connector: "file", ...(req as any) })}`;
    }

    if (kind === "http:request") {
      const req = request && typeof request === "object" && !Array.isArray(request) ? (request as any) : {};
      const url = typeof req.url === "string" ? req.url : "";
      const method = (typeof req.method === "string" ? req.method : "GET").toUpperCase();
      const headers = req.headers && typeof req.headers === "object" ? req.headers : {};
      const responseType = typeof req.responseType === "string" ? req.responseType : "auto";
      const jsonPath = typeof req.jsonPath === "string" ? req.jsonPath : "";
      const key: any = { connector: "http", url, method, headers, responseType, jsonPath };
      if (req.auth) key.auth = req.auth;
      return `${kind}:${hashValue(key)}`;
    }

    // For unknown permission kinds, fall back to hashing the details object.
    return `${kind}:${hashValue(details)}`;
  } catch {
    return null;
  }
}

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

function normalizeFileSize(payload: unknown): number {
  if (payload == null) {
    throw new Error("Unexpected stat payload returned from filesystem API");
  }
  if (typeof payload === "number") {
    if (!Number.isFinite(payload) || payload < 0) {
      throw new Error("Unexpected file size returned from filesystem API");
    }
    return payload;
  }
  if (typeof payload === "string") {
    const numeric = Number(payload);
    if (Number.isFinite(numeric) && numeric >= 0) {
      return numeric;
    }
  }
  if (payload && typeof payload === "object") {
    const obj = payload as any;
    const candidate =
      obj.sizeBytes ??
      obj.size_bytes ??
      obj.size ??
      obj.length ??
      obj.len ??
      obj.fileSize ??
      obj.file_size ??
      obj.bytes ??
      null;
    if (candidate != null) return normalizeFileSize(candidate);
  }
  throw new Error("Unexpected stat payload returned from filesystem API (missing file size)");
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
    const openFile = async (path: string): Promise<Blob> => {
      // Best-effort: the FS plugin does not currently expose a streamable file handle, so fall back
      // to an in-memory Blob.
      const bytes = normalizeBinaryPayload(await readFile(path));
      return new Blob([bytes]);
    };

    return {
      readText: async (path) => readTextFile(path),
      readBinary: async (path) => normalizeBinaryPayload(await readFile(path)),
      readBinaryStream: async function* (path) {
        yield normalizeBinaryPayload(await readFile(path));
      },
      openFile,
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
    readBinaryStream: async function* (path, options = {}) {
      const signal = options.signal;
      const chunkSize = 1024 * 1024; // 1MiB
      let offset = 0;

      while (true) {
        if (signal?.aborted) {
          const err = new Error("Aborted");
          (err as any).name = "AbortError";
          throw err;
        }
        const payload = await invoke("read_binary_file_range", { path, offset, length: chunkSize });
        const bytes = normalizeBinaryPayload(payload);
        if (bytes.length === 0) break;
        yield bytes;
        offset += bytes.length;
        if (bytes.length < chunkSize) break;
      }
    },
    openFile: async (path, options = {}) => {
      const signal = options.signal;
      if (signal?.aborted) {
        const err = new Error("Aborted");
        (err as any).name = "AbortError";
        throw err;
      }

      const statPayload = await invoke("stat_file", { path });
      const fileSize = normalizeFileSize(statPayload);
      if (signal?.aborted) {
        const err = new Error("Aborted");
        (err as any).name = "AbortError";
        throw err;
      }

      class TauriFileBlob {
        path: string;
        invoke: TauriInvoke;
        start: number;
        end: number;
        size: number;
        type: string;

        constructor(path: string, invoke: TauriInvoke, start: number, end: number) {
          this.path = path;
          this.invoke = invoke;
          this.start = start;
          this.end = end;
          this.size = Math.max(0, end - start);
          this.type = "";
          // Best-effort compatibility with code that checks for Blob-ish objects.
          (this as any)[Symbol.toStringTag] = "Blob";
        }

        slice(start = 0, end = this.size): TauriFileBlob {
          const sliceStart = Math.max(0, Math.min(this.size, start));
          const sliceEnd = Math.max(sliceStart, Math.min(this.size, end));
          return new TauriFileBlob(this.path, this.invoke, this.start + sliceStart, this.start + sliceEnd);
        }

        async arrayBuffer(): Promise<ArrayBuffer> {
          const length = this.size;
          if (length <= 0) return new ArrayBuffer(0);
          const payload = await this.invoke("read_binary_file_range", { path: this.path, offset: this.start, length });
          const bytes = normalizeBinaryPayload(payload);
          return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
        }
      }

      return new TauriFileBlob(path, invoke, 0, fileSize) as unknown as Blob;
    },
    stat: async (path) => ({ mtimeMs: normalizeMtimeMs(await invoke("stat_file", { path })) }),
  };
}

function createDefaultCacheManager(): CacheManager {
  // Keep Power Query caches bounded so long-lived desktop sessions don't accumulate
  // unbounded IndexedDB storage.
  const limits = { maxEntries: 256, maxBytes: 256 * 1024 * 1024 };
  const store =
    typeof indexedDB !== "undefined" ? new IndexedDBCacheStore({ dbName: "formula-power-query-cache" }) : new MemoryCacheStore();
  const cache = new CacheManager({ store, limits });
  // Best-effort: prune on startup to enforce quotas even if the cache directory/db
  // was created by an older version without eviction.
  cache.prune().catch(() => {});
  return cache;
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

type PrivacyLevel = "public" | "organizational" | "private" | "unknown";

function classificationLevelToPrivacy(level: unknown): PrivacyLevel {
  switch (level) {
    case "Public":
      return "public";
    case "Internal":
      return "organizational";
    case "Confidential":
    case "Restricted":
      return "private";
    default:
      return "unknown";
  }
}

function computeWorkbookPrivacyLevel(dlp: DlpContext | undefined): PrivacyLevel {
  if (!dlp) return "unknown";
  try {
    const records = dlp.classificationStore.list(dlp.documentId);
    if (dlp.sheetId && dlp.range) {
      try {
        const selection = effectiveRangeClassification(
          { documentId: dlp.documentId, sheetId: dlp.sheetId, range: dlp.range },
          records,
        );
        return classificationLevelToPrivacy((selection as any)?.level);
      } catch {
        // Fall back to document classification below.
      }
    }
    const doc = effectiveDocumentClassification(dlp.documentId, records);
    return classificationLevelToPrivacy((doc as any)?.level);
  } catch {
    return "unknown";
  }
}

function collectWorkbookTableSourceIds(query: any, queries: Record<string, any> | undefined): Set<string> {
  const out = new Set<string>();
  const visited = new Set<string>();

  const visit = (q: any) => {
    if (!q || typeof q !== "object") return;
    if (typeof q.id === "string") {
      if (visited.has(q.id)) return;
      visited.add(q.id);
    }

    const source = q.source;
    if (source && typeof source === "object") {
      if (source.type === "table" && typeof source.table === "string") {
        out.add(`workbook:table:${source.table}`);
      } else if (source.type === "range") {
        out.add("workbook:range");
      } else if (source.type === "query" && typeof source.queryId === "string") {
        const dep = queries?.[source.queryId];
        if (dep) visit(dep);
      }
    }

    const steps = Array.isArray(q.steps) ? q.steps : [];
    for (const step of steps) {
      const op = step?.operation;
      if (!op || typeof op !== "object") continue;
      if (op.type === "merge" && typeof op.rightQuery === "string") {
        const dep = queries?.[op.rightQuery];
        if (dep) visit(dep);
      } else if (op.type === "append" && Array.isArray(op.queries)) {
        for (const id of op.queries) {
          if (typeof id !== "string") continue;
          const dep = queries?.[id];
          if (dep) visit(dep);
        }
      }
    }
  };

  visit(query);
  return out;
}

class DesktopQueryEngine extends QueryEngine {
  private defaultPrivacyLevelsBySourceId: Record<string, PrivacyLevel>;
  private workbookPrivacyLevel: PrivacyLevel;

  constructor(options: any, privacy: { levelsBySourceId: Record<string, PrivacyLevel>; workbookLevel: PrivacyLevel }) {
    super(options);
    this.defaultPrivacyLevelsBySourceId = privacy.levelsBySourceId;
    this.workbookPrivacyLevel = privacy.workbookLevel;
  }

  private withPrivacyContext(query: any, context: any = {}) {
    const base = { ...this.defaultPrivacyLevelsBySourceId };
    if (this.workbookPrivacyLevel !== "unknown") {
      const ids = collectWorkbookTableSourceIds(query, context?.queries);
      for (const id of ids) base[id] = this.workbookPrivacyLevel;
      base["workbook:range"] = this.workbookPrivacyLevel;

      // Cover host-provided context objects where the query graph is incomplete
      // (e.g. refresh orchestration supplies `queryResults` without `queries`).
      const tables = context?.tables;
      if (tables && typeof tables === "object") {
        for (const name of Object.keys(tables)) {
          if (typeof name === "string" && name.length > 0) {
            base[`workbook:table:${name}`] = this.workbookPrivacyLevel;
          }
        }
      }

      const queryResults = context?.queryResults;
      if (queryResults && typeof queryResults === "object") {
        for (const result of Object.values(queryResults)) {
          const sources = result && typeof result === "object" ? (result as any).meta?.sources : null;
          if (!Array.isArray(sources)) continue;
          for (const source of sources) {
            const provenance = source?.provenance;
            if (!provenance || typeof provenance !== "object") continue;
            const kind = (provenance as any).kind;
            if (kind === "table" && typeof (provenance as any).table === "string") {
              base[`workbook:table:${(provenance as any).table}`] = this.workbookPrivacyLevel;
            } else if (kind === "range") {
              base["workbook:range"] = this.workbookPrivacyLevel;
            }
          }
        }
      }
    }

    const overrides = context?.privacy?.levelsBySourceId;
    const mergedLevels = overrides ? { ...base, ...overrides } : base;
    if (Object.keys(mergedLevels).length === 0) return context;
    return { ...context, privacy: { ...(context?.privacy ?? {}), levelsBySourceId: mergedLevels } };
  }

  async executeQuery(query: any, context: any = {}, options: any = {}) {
    return super.executeQuery(query, this.withPrivacyContext(query, context), options);
  }

  async executeQueryWithMeta(query: any, context: any = {}, options: any = {}) {
    return super.executeQueryWithMeta(query, this.withPrivacyContext(query, context), options);
  }

  async executeQueryWithMetaInSession(query: any, context: any = {}, options: any = {}, session: any) {
    return super.executeQueryWithMetaInSession(query, this.withPrivacyContext(query, context), options, session);
  }

  async executeQueryStreaming(query: any, context: any = {}, options: any) {
    return super.executeQueryStreaming(query, this.withPrivacyContext(query, context), options);
  }

  async getCacheKey(query: any, context: any = {}, options: any = {}) {
    return super.getCacheKey(query, this.withPrivacyContext(query, context), options);
  }

  async invalidateQueryCache(query: any, context: any = {}, options: any = {}) {
    return super.invalidateQueryCache(query, this.withPrivacyContext(query, context), options);
  }
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

  // Prefer a host-provided queued invoke (set by the desktop entrypoint) so reads
  // like `list_tables`/`get_range` cannot race ahead of pending workbook writes.
  const tauriInvoke = ((globalThis as any).__FORMULA_WORKBOOK_INVOKE__ ??
    (globalThis as any).__TAURI__?.core?.invoke) as TauriInvoke | undefined;
  const tableAdapter =
    typeof tauriInvoke === "function"
      ? {
          getTable: async (tableName: string, adapterOptions?: { signal?: AbortSignal }) => {
            if (adapterOptions?.signal?.aborted) {
              const err = new Error("Aborted");
              err.name = "AbortError";
              throw err;
            }

            const tables = (await tauriInvoke("list_tables")) as any;
            const list = Array.isArray(tables) ? tables : [];
            let info = list.find((t) => t && typeof t === "object" && (t as any).name === tableName);
            if (!info) {
              const lower = tableName.toLowerCase();
              info = list.find(
                (t) => t && typeof t === "object" && typeof (t as any).name === "string" && (t as any).name.toLowerCase() === lower,
              );
            }
            if (!info) {
              const definedNames = (await tauriInvoke("list_defined_names")) as any;
              const names = Array.isArray(definedNames) ? definedNames : [];
              let defined = names.find((n) => n && typeof n === "object" && (n as any).name === tableName);
              if (!defined) {
                const lower = tableName.toLowerCase();
                defined = names.find(
                  (n) => n && typeof n === "object" && typeof (n as any).name === "string" && (n as any).name.toLowerCase() === lower,
                );
              }
              if (!defined) {
                throw new Error(`Unknown table or defined name '${tableName}'`);
              }

              const rawRefersTo = typeof (defined as any).refers_to === "string" ? String((defined as any).refers_to) : "";
              const refersTo = rawRefersTo.trim().startsWith("=") ? rawRefersTo.trim().slice(1).trim() : rawRefersTo.trim();
              const { sheetName, ref } = splitSheetQualifier(refersTo);
              const sheetId =
                sheetName ?? (typeof (defined as any).sheet_id === "string" ? String((defined as any).sheet_id) : null);
              if (!sheetId || sheetId.trim() === "") {
                throw new Error(
                  `Defined name '${tableName}' refers to '${rawRefersTo}', but no sheet was provided. Only simple A1 ranges are supported (e.g. Sheet1!$A$1:$B$2)`,
                );
              }

              let range;
              try {
                range = parseA1Range(ref);
              } catch {
                throw new Error(
                  `Defined name '${tableName}' refers to unsupported formula '${rawRefersTo}'. Only simple A1 ranges are supported (e.g. Sheet1!$A$1:$B$2)`,
                );
              }

              const rangePayload = (await tauriInvoke("get_range", {
                sheet_id: sheetId,
                start_row: range.startRow,
                start_col: range.startCol,
                end_row: range.endRow,
                end_col: range.endCol,
              })) as any;

              const rows = Array.isArray(rangePayload?.values) ? rangePayload.values : [];
              /** @type {unknown[][]} */
              const grid = rows.map((row: any) =>
                Array.isArray(row)
                  ? row.map((cell) => (cell && typeof cell === "object" ? (cell as any).value ?? null : null))
                  : [],
              );

              const hasHeaders =
                grid.length >= 2 &&
                Array.isArray(grid[0]) &&
                (grid[0] as any[]).some((v) => typeof v === "string" && v.trim() !== "") &&
                (grid[0] as any[]).every((v) => v == null || typeof v === "string");

              return DataTable.fromGrid(grid, { hasHeaders, inferTypes: true });
            }

            const columns = Array.isArray((info as any).columns) ? ((info as any).columns as string[]) : [];
            const width = columns.length;
            const range = (await tauriInvoke("get_range", {
              sheet_id: (info as any).sheet_id,
              start_row: (info as any).start_row,
              start_col: (info as any).start_col,
              end_row: (info as any).end_row,
              end_col: (info as any).end_col,
            })) as any;

            const rows = Array.isArray(range?.values) ? range.values : [];
            /** @type {unknown[][]} */
            const grid = rows.map((row: any) => {
              const values = Array.isArray(row) ? row.map((cell) => (cell && typeof cell === "object" ? (cell as any).value ?? null : null)) : [];
              if (width === 0) return values;
              const out = values.slice(0, width);
              while (out.length < width) out.push(null);
              return out;
            });

            if (grid.length === 0) {
              grid.push(columns.slice());
            } else if (columns.length > 0) {
              grid[0] = columns.slice();
            }

            return DataTable.fromGrid(grid, { hasHeaders: true, inferTypes: true });
          },
          // Note: table source cache safety is handled via `QueryExecutionContext.getTableSignature`
          // (see `getContextForDocument` in this module), which incorporates a per-table signature
          // into the Power Query cache key.
        }
      : undefined;

  const http =
    options.fetch || options.oauth2Manager
      ? new HttpConnector({ fetch: options.fetch, oauth2Manager: options.oauth2Manager })
      : undefined;

  // Cache permission prompts across executions so previewing the same query
  // doesn't repeatedly ask the user.
  const permissionPromptCache = new Map<string, Promise<boolean>>();
  const getPermissionObjectId = createEphemeralObjectId();

  const workbookPrivacyLevel = computeWorkbookPrivacyLevel(options.dlp);
  /** @type {Record<string, PrivacyLevel>} */
  const defaultPrivacyLevelsBySourceId = {};
  if (workbookPrivacyLevel !== "unknown") {
    defaultPrivacyLevelsBySourceId["workbook:range"] = workbookPrivacyLevel;
    // Provide a fallback for table provenance that does not include a specific table name.
    defaultPrivacyLevelsBySourceId["workbook:table"] = workbookPrivacyLevel;
  }

  return new DesktopQueryEngine(
    {
      cache,
      defaultCacheTtlMs: options.defaultCacheTtlMs,
      fileAdapter: {
        readText: fileAdapter.readText,
        readBinary: fileAdapter.readBinary,
        readBinaryStream: fileAdapter.readBinaryStream,
        openFile: fileAdapter.openFile,
        stat: fileAdapter.stat,
      },
      tableAdapter,
      connectors: http ? { http } : undefined,
      privacyMode: options.privacyMode,
      onCredentialRequest: options.onCredentialRequest,
      onPermissionRequest: async (kind, details) => {
        const dlpAction = PERMISSION_KIND_TO_DLP_ACTION[kind];
        if (dlpAction === DLP_ACTION.EXTERNAL_CONNECTOR && options.dlp) {
          const policy = typeof options.dlp.policy === "function" ? await options.dlp.policy() : options.dlp.policy;
          enforceExternalConnector({
            documentId: options.dlp.documentId,
            sheetId: options.dlp.sheetId,
            range: options.dlp.range,
            classificationStore: options.dlp.classificationStore,
            policy,
          });
        }

        const cacheKey = permissionPromptCacheKey(kind, details, { getObjectId: getPermissionObjectId });
        if (cacheKey) {
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
        }

        if (options.onPermissionPrompt) {
          return await options.onPermissionPrompt(kind, details);
        }
        return defaultPermissionPrompt(kind, details);
      },
    },
    { levelsBySourceId: defaultPrivacyLevelsBySourceId, workbookLevel: workbookPrivacyLevel },
  );
}

export function getContextForDocument(doc: DocumentController): QueryExecutionContext {
  const registry = getTableSignatureRegistry(doc);
  return {
    getTableSignature: (tableName) => registry.getTableSignature(tableName),
  };
}
