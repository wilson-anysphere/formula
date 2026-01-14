import type { AIAuditStore } from "./store.ts";
import type { AIAuditEntry, AIMode, AuditListFilters } from "./types.ts";
import { stableStringify } from "./stable-json.ts";

export interface IndexedDbAIAuditStoreOptions {
  /**
   * IndexedDB database name. Defaults to `formula_ai_audit`.
   */
  db_name?: string;
  /**
   * Object store name within the database. Defaults to `ai_audit_log`.
   */
  store_name?: string;
  /**
   * Maximum number of entries to retain (newest retained). If unset, entries are unbounded.
   *
   * Semantics match `SqliteAIAuditStoreRetention.max_entries`.
   */
  max_entries?: number;
  /**
   * Maximum age in milliseconds. Entries older than (now - max_age_ms) are deleted at
   * write-time. If unset, age-based retention is disabled.
   *
   * Semantics match `SqliteAIAuditStoreRetention.max_age_ms`.
   */
  max_age_ms?: number;
}

const DEFAULT_DB_NAME = "formula_ai_audit";
const DEFAULT_STORE_NAME = "ai_audit_log";
const INDEXEDDB_SCHEMA_VERSION = 1;

/**
 * IndexedDB-backed audit store intended for browser contexts where `localStorage`
 * quota limits and full-array rewrites are undesirable.
 *
 * Note: IndexedDB also has quota limits. For defense-in-depth against oversized
 * single entries, wrap this store with `BoundedAIAuditStore`.
 */
export class IndexedDbAIAuditStore implements AIAuditStore {
  private readonly dbName: string;
  private readonly storeName: string;
  private readonly maxEntries?: number;
  private readonly maxAgeMs?: number;

  private dbPromise: Promise<IDBDatabase> | null = null;

  constructor(options: IndexedDbAIAuditStoreOptions = {}) {
    this.dbName = options.db_name ?? DEFAULT_DB_NAME;
    this.storeName = options.store_name ?? DEFAULT_STORE_NAME;
    this.maxEntries = options.max_entries;
    this.maxAgeMs = options.max_age_ms;
  }

  async logEntry(entry: AIAuditEntry): Promise<void> {
    const db = await this.getDb();
    const record = normalizeEntry(entry);

    await withTransaction(db, this.storeName, "readwrite", async (store) => {
      const request = addWithCloneFallback(store, record);
      await requestToPromise(request);
    });

    await this.enforceRetention(db);
  }

  async listEntries(filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    const db = await this.getDb();

    // Retention is primarily enforced on write, but list-time enforcement keeps
    // long-lived stores bounded even when no new writes occur.
    try {
      await this.enforceRetention(db);
    } catch {
      // Best-effort: failures should not prevent reading existing audit logs.
    }

    const sessionId = filters.session_id;
    const workbookId = filters.workbook_id;
    const modeFilter = normalizeModes(filters.mode);
    const limit = normalizeLimit(filters.limit);
    if (limit === 0) return [];

    const maxAgeCutoff = computeCutoffMs(this.maxAgeMs);
    const afterTimestampMs = normalizeFiniteNumber(filters.after_timestamp_ms);
    const beforeTimestampMs = normalizeFiniteNumber(filters.before_timestamp_ms);
    const cursorTs = normalizeFiniteNumber(filters.cursor?.before_timestamp_ms);
    const cursorId = typeof filters.cursor?.before_id === "string" ? filters.cursor.before_id : undefined;

    const minTimestampInclusive = maxOfFinite(maxAgeCutoff, afterTimestampMs);
    const upperBound = computeUpperBound(beforeTimestampMs, cursorTs);

    return withTransaction(db, this.storeName, "readonly", async (store) => {
      const index = store.index("timestamp_ms");
      const out: AIAuditEntry[] = [];

      const range = upperBound ? createUpperBoundRange(upperBound.value, upperBound.open) : null;

      await cursorToPromise(index.openCursor(range, "prev"), (cursor) => {
        const value = cursor.value as AIAuditEntry;
        if (!value || typeof value !== "object") {
          cursor.continue();
          return;
        }

        if (typeof minTimestampInclusive === "number" && value.timestamp_ms < minTimestampInclusive) {
          // Cursor is descending by timestamp, so once we hit the lower bound we can stop.
          return false;
        }

        if (typeof beforeTimestampMs === "number" && value.timestamp_ms >= beforeTimestampMs) {
          cursor.continue();
          return;
        }

        if (typeof cursorTs === "number") {
          if (value.timestamp_ms > cursorTs) {
            cursor.continue();
            return;
          }
          if (value.timestamp_ms === cursorTs) {
            if (!cursorId) {
              cursor.continue();
              return;
            }
            if (!(String(value.id) < cursorId)) {
              cursor.continue();
              return;
            }
          }
        }

        if (sessionId && value.session_id !== sessionId) {
          cursor.continue();
          return;
        }

        if (workbookId && !matchesWorkbookFilter(value, workbookId)) {
          cursor.continue();
          return;
        }

        if (modeFilter && modeFilter.length > 0 && !modeFilter.includes(value.mode as AIMode)) {
          cursor.continue();
          return;
        }

        out.push(value);
        if (limit !== undefined && out.length >= limit) return false;
        cursor.continue();
      });

      return out;
    });
  }

  private async getDb(): Promise<IDBDatabase> {
    if (!this.dbPromise) {
      this.dbPromise = openDatabase(this.dbName, this.storeName).catch((err) => {
        // Ensure subsequent calls retry if open failed.
        this.dbPromise = null;
        throw err;
      });
    }
    return this.dbPromise;
  }

  private async enforceRetention(db: IDBDatabase): Promise<void> {
    const maxAgeMs = normalizeFinitePositive(this.maxAgeMs);
    const maxEntries = normalizeFinitePositive(this.maxEntries);
    if (!maxAgeMs && !maxEntries) return;

    await withTransaction(db, this.storeName, "readwrite", async (store) => {
      if (maxAgeMs) {
        const cutoff = Date.now() - maxAgeMs;
        const index = store.index("timestamp_ms");
        const range = createUpperBoundRange(cutoff, true);

        if (range) {
          await cursorToPromise(index.openCursor(range, "next"), (cursor) => {
            cursor.delete();
            cursor.continue();
          });
        } else {
          // Fallback for environments without `IDBKeyRange`.
          await cursorToPromise(index.openCursor(null, "next"), (cursor) => {
            const key = cursor.key;
            const ts = typeof key === "number" ? key : Number(key);
            if (!Number.isFinite(ts) || ts >= cutoff) return false;
            cursor.delete();
            cursor.continue();
          });
        }
      }

      if (maxEntries) {
        const cap = Math.floor(maxEntries);
        const total = await requestToPromise(store.count());
        const over = total - cap;
        if (over <= 0) return;

        const index = store.index("timestamp_ms");
        let deleted = 0;
        await cursorToPromise(index.openCursor(null, "next"), (cursor) => {
          if (deleted >= over) return false;
          cursor.delete();
          deleted += 1;
          cursor.continue();
        });
      }
    });
  }
}

function addWithCloneFallback(store: IDBObjectStore, value: AIAuditEntry): IDBRequest<IDBValidKey> {
  try {
    return store.add(value);
  } catch (err) {
    // IndexedDB uses structured clone for values; some inputs (functions, some BigInt
    // implementations, etc) can throw `DataCloneError`. Fall back to a JSON-safe
    // clone so audit logging doesn't crash.
    if (!isDataCloneError(err)) throw err;
    return store.add(safeJsonClone(value));
  }
}

function isDataCloneError(err: unknown): boolean {
  if (!err || typeof err !== "object") return false;
  // `DataCloneError` is usually surfaced as a DOMException.
  return (err as { name?: unknown }).name === "DataCloneError";
}

function safeJsonClone(entry: AIAuditEntry): AIAuditEntry {
  try {
    const parsed = JSON.parse(stableStringify(entry)) as unknown;
    return parsed && typeof parsed === "object" ? (parsed as AIAuditEntry) : entry;
  } catch {
    return entry;
  }
}

function openDatabase(dbName: string, storeName: string): Promise<IDBDatabase> {
  const idb: IDBFactory | undefined = (globalThis as any).indexedDB;
  if (!idb?.open) {
    return Promise.reject(new Error("IndexedDbAIAuditStore: indexedDB.open is unavailable in this environment"));
  }

  return new Promise((resolve, reject) => {
    const request = idb.open(dbName, INDEXEDDB_SCHEMA_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      const store = db.objectStoreNames.contains(storeName)
        ? request.transaction!.objectStore(storeName)
        : db.createObjectStore(storeName, { keyPath: "id" });

      ensureIndex(store, "timestamp_ms", "timestamp_ms");
      ensureIndex(store, "session_id", "session_id");
      ensureIndex(store, "workbook_id", "workbook_id");
      ensureIndex(store, "mode", "mode");
    };

    request.onsuccess = () => {
      const db = request.result;
      db.onversionchange = () => {
        try {
          db.close();
        } catch {
          // ignore
        }
      };
      resolve(db);
    };
    request.onerror = () => reject(request.error ?? new Error("IndexedDbAIAuditStore: indexedDB.open failed"));
    request.onblocked = () => reject(new Error("IndexedDbAIAuditStore: indexedDB.open was blocked"));
  });
}

function ensureIndex(store: IDBObjectStore, name: string, keyPath: string): void {
  if (store.indexNames.contains(name)) return;
  store.createIndex(name, keyPath, { unique: false });
}

async function withTransaction<T>(
  db: IDBDatabase,
  storeName: string,
  mode: IDBTransactionMode,
  fn: (store: IDBObjectStore) => Promise<T>
): Promise<T> {
  let tx: IDBTransaction;
  try {
    tx = db.transaction(storeName, mode);
  } catch (err) {
    const name = typeof storeName === "string" ? storeName : String(storeName);
    const details = err instanceof Error ? err.message : String(err);
    throw new Error(`IndexedDbAIAuditStore: object store "${name}" is unavailable (${details})`);
  }

  const store = tx.objectStore(storeName);
  const result = await fn(store);
  await transactionDone(tx);
  return result;
}

function transactionDone(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDbAIAuditStore: transaction aborted"));
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDbAIAuditStore: transaction failed"));
  });
}

function requestToPromise<T = unknown>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("IndexedDbAIAuditStore: request failed"));
  });
}

type CursorHandler = (cursor: IDBCursorWithValue) => void | false;

function cursorToPromise(request: IDBRequest<IDBCursorWithValue | null>, onCursor: CursorHandler): Promise<void> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => {
      const cursor = request.result;
      if (!cursor) {
        resolve();
        return;
      }
      try {
        const res = onCursor(cursor);
        if (res === false) resolve();
      } catch (err) {
        reject(err);
      }
    };
    request.onerror = () => reject(request.error ?? new Error("IndexedDbAIAuditStore: cursor failed"));
  });
}

function normalizeModes(mode: AuditListFilters["mode"]): AIMode[] | undefined {
  if (!mode) return undefined;
  const modes = Array.isArray(mode) ? mode : [mode];
  return modes.length > 0 ? modes : undefined;
}

function normalizeLimit(limit: AuditListFilters["limit"]): number | undefined {
  if (typeof limit !== "number") return undefined;
  if (!Number.isFinite(limit)) return undefined;
  return Math.max(0, Math.trunc(limit));
}

function normalizeFinitePositive(value: number | undefined): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) return undefined;
  return value;
}

function normalizeFiniteNumber(value: number | undefined): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  return value;
}

function maxOfFinite(...values: Array<number | undefined>): number | undefined {
  let out: number | undefined;
  for (const value of values) {
    if (typeof value !== "number" || !Number.isFinite(value)) continue;
    out = out === undefined ? value : Math.max(out, value);
  }
  return out;
}

type UpperBound = { value: number; open: boolean };

function computeUpperBound(beforeTimestampMs: number | undefined, cursorTimestampMs: number | undefined): UpperBound | undefined {
  let out: UpperBound | undefined;

  if (typeof beforeTimestampMs === "number") {
    out = { value: beforeTimestampMs, open: true };
  }

  if (typeof cursorTimestampMs === "number") {
    if (!out || cursorTimestampMs < out.value) {
      out = { value: cursorTimestampMs, open: false };
    } else if (cursorTimestampMs === out.value && out.open === false) {
      out = { value: cursorTimestampMs, open: false };
    }
  }

  return out;
}

function computeCutoffMs(maxAgeMs: number | undefined): number | undefined {
  const maxAge = normalizeFinitePositive(maxAgeMs);
  if (!maxAge) return undefined;
  return Date.now() - maxAge;
}

function createUpperBoundRange(value: number, open: boolean): IDBKeyRange | null {
  const ctor: typeof IDBKeyRange | undefined = (globalThis as any).IDBKeyRange;
  if (ctor?.upperBound) {
    try {
      return ctor.upperBound(value, open);
    } catch {
      return null;
    }
  }
  return null;
}

function normalizeEntry(entry: AIAuditEntry): AIAuditEntry {
  // Ensure workbook_id is persisted for efficient filtering via IndexedDB indexes,
  // even when older integrations omitted the field.
  const workbookIdRaw = entry.workbook_id ?? extractWorkbookIdFromInput(entry.input) ?? extractWorkbookIdFromSessionId(entry.session_id);
  const workbookId = typeof workbookIdRaw === "string" ? workbookIdRaw.trim() : "";
  if (!workbookId) return entry;
  if (entry.workbook_id === workbookId) return entry;
  return { ...entry, workbook_id: workbookId };
}

function extractWorkbookIdFromInput(input: unknown): string | null {
  if (!input || typeof input !== "object") return null;
  const obj = input as Record<string, unknown>;
  const workbookId = obj.workbook_id ?? obj.workbookId;
  const trimmed = typeof workbookId === "string" ? workbookId.trim() : "";
  return trimmed ? trimmed : null;
}

function extractWorkbookIdFromSessionId(sessionId: string): string | null {
  const match = sessionId.match(/^([^:]+):([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})$/);
  if (!match) return null;
  const workbookId = match[1];
  const trimmed = workbookId?.trim() ?? "";
  return trimmed ? trimmed : null;
}

function matchesWorkbookFilter(entry: AIAuditEntry, workbookId: string): boolean {
  const filterId = typeof workbookId === "string" ? workbookId.trim() : "";
  if (!filterId) return false;

  if (typeof entry.workbook_id === "string") {
    const entryId = entry.workbook_id.trim();
    if (entryId) return entryId === filterId;
  }

  const input = entry.input as unknown;
  if (!input || typeof input !== "object") return false;
  const obj = input as Record<string, unknown>;
  const legacyWorkbookId = obj.workbook_id ?? obj.workbookId;
  const legacyId = typeof legacyWorkbookId === "string" ? legacyWorkbookId.trim() : "";
  return legacyId ? legacyId === filterId : false;
}
