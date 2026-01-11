import type { AIAuditEntry, AIAuditStore, AuditListFilters } from "@formula/ai-audit/browser";
import { LocalStorageAIAuditStore, LocalStorageBinaryStorage } from "@formula/ai-audit/browser";

import sqlWasmUrl from "sql.js/dist/sql-wasm.wasm?url";

export const DESKTOP_AI_AUDIT_DB_STORAGE_KEY = "formula:ai_audit_db:v1";

export interface DesktopAIAuditStoreOptions {
  /**
   * Overrides the localStorage key used to persist the sqlite database.
   * Defaults to `DESKTOP_AI_AUDIT_DB_STORAGE_KEY`.
   */
  storageKey?: string;
  /**
   * Maximum number of audit entries to retain in the sqlite-backed store.
   *
   * Defaults to 10k entries.
   */
  retentionMaxEntries?: number;
  /**
   * Maximum age (in ms) to retain in the sqlite-backed store.
   *
   * Defaults to 30 days.
   */
  retentionMaxAgeMs?: number;
}

const DEFAULT_RETENTION_MAX_ENTRIES = 10_000;
const DEFAULT_RETENTION_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;

const storePromiseByKey = new Map<string, Promise<AIAuditStore>>();

async function createSqliteBackedStore(params: { storageKey: string; retentionMaxEntries: number; retentionMaxAgeMs: number }) {
  const { SqliteAIAuditStore } = await import("@formula/ai-audit/sqlite");
  return SqliteAIAuditStore.create({
    storage: new LocalStorageBinaryStorage(params.storageKey),
    // Ensure the sql.js WASM file is bundled by Vite and can be fetched at runtime.
    locateFile: (file: string) => (file.endsWith(".wasm") ? sqlWasmUrl : file),
    retention: { max_entries: params.retentionMaxEntries, max_age_ms: params.retentionMaxAgeMs },
  });
}

async function resolveDesktopAIAuditStore(options: DesktopAIAuditStoreOptions = {}): Promise<AIAuditStore> {
  const storageKey = options.storageKey ?? DESKTOP_AI_AUDIT_DB_STORAGE_KEY;
  const retentionMaxEntries = options.retentionMaxEntries ?? DEFAULT_RETENTION_MAX_ENTRIES;
  const retentionMaxAgeMs = options.retentionMaxAgeMs ?? DEFAULT_RETENTION_MAX_AGE_MS;

  const cached = storePromiseByKey.get(storageKey);
  if (cached) return cached;

  const promise = createSqliteBackedStore({ storageKey, retentionMaxEntries, retentionMaxAgeMs }).catch((_err) => {
    // Best-effort fallback: keep audit logging functional even if sql.js fails to load
    // (e.g. blocked WASM fetch).
    return new LocalStorageAIAuditStore();
  });
  storePromiseByKey.set(storageKey, promise);
  return promise;
}

class LazyAIAuditStore implements AIAuditStore {
  private resolved: AIAuditStore | null = null;
  private resolving: Promise<AIAuditStore> | null = null;

  constructor(private readonly options: DesktopAIAuditStoreOptions) {}

  private async getStore(): Promise<AIAuditStore> {
    if (this.resolved) return this.resolved;
    if (!this.resolving) {
      this.resolving = resolveDesktopAIAuditStore(this.options).then((store) => {
        this.resolved = store;
        return store;
      });
    }
    return this.resolving;
  }

  async logEntry(entry: AIAuditEntry): Promise<void> {
    const store = await this.getStore();
    await store.logEntry(entry);
  }

  async listEntries(filters?: AuditListFilters): Promise<AIAuditEntry[]> {
    const store = await this.getStore();
    return store.listEntries(filters);
  }
}

/**
 * Returns an `AIAuditStore` suitable for the desktop app:
 * - sqlite-backed (sql.js) storage persisted via `LocalStorageBinaryStorage`
 * - falls back to JSON localStorage on initialization failures
 *
 * The returned store is safe to construct synchronously; it lazily initializes
 * the underlying sqlite store on first use.
 */
export function getDesktopAIAuditStore(options: DesktopAIAuditStoreOptions = {}): AIAuditStore {
  return new LazyAIAuditStore(options);
}
