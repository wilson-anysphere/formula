import type { AIAuditEntry, AIAuditStore, AuditListFilters } from "@formula/ai-audit/browser";
import { BoundedAIAuditStore, LocalStorageAIAuditStore, LocalStorageBinaryStorage } from "@formula/ai-audit/browser";

import sqlWasmUrl from "sql.js/dist/sql-wasm.wasm?url";

export const DESKTOP_AI_AUDIT_DB_STORAGE_KEY = "formula:ai_audit_db:v1";

export interface DesktopAIAuditStoreOptions {
  /**
   * Overrides the localStorage key used to persist the sqlite database.
   * Defaults to `DESKTOP_AI_AUDIT_DB_STORAGE_KEY`.
   */
  storageKey?: string;
  /**
   * Hard cap (in characters) for the serialized size of a single audit entry.
   *
   * This is enforced via `BoundedAIAuditStore` before writing to the underlying
   * sqlite/localStorage store to avoid persistence failures when entries grow
   * unexpectedly large (quota limits).
   *
   * Defaults to `200_000`.
   */
  maxEntryChars?: number;
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

function isNodeRuntime(): boolean {
  // Avoid dot-accessing Node version info on `process.versions` so the desktop/WebView bundle stays Node-free.
  // We only need a best-effort Node detector for Vitest/Node environments where
  // sql.js uses `fs.readFileSync` for WASM loading.
  const proc = (globalThis as any).process as any;
  if (!proc || typeof proc !== "object") return false;
  if (proc.release && typeof proc.release === "object" && proc.release.name === "node") return true;
  // Fallback for odd environments: `process.version` is a Node-only string like "v22.10.0".
  return typeof proc.version === "string" && proc.version.startsWith("v");
}

function coerceViteUrlToNodeFileUrl(href: string): string {
  if (!href) return href;
  if (href.startsWith("file://")) return href;
  if (!isNodeRuntime()) return href;

  // Vite asset URLs are typically root-relative (`/node_modules/...` or `/assets/...`).
  // In Node, sql.js uses `fs.readFileSync` for wasm loading, so convert these to a
  // file:// URL rooted at the repository cwd.
  if (href.startsWith("/")) {
    const cwd = typeof (globalThis as any).process?.cwd === "function" ? (globalThis as any).process.cwd() : "";
    if (cwd) return `file://${cwd}${href}`;
  }

  return href;
}

async function createSqliteBackedStore(params: { storageKey: string; retentionMaxEntries: number; retentionMaxAgeMs: number }) {
  const { SqliteAIAuditStore } = await import("@formula/ai-audit/sqlite");
  return SqliteAIAuditStore.create({
    storage: new LocalStorageBinaryStorage(params.storageKey),
    // Ensure the sql.js WASM file is bundled by Vite and can be fetched at runtime.
    locateFile: (file: string, prefix: string = "") => {
      if (file.endsWith(".wasm")) {
        // In Node-based test runners (Vitest), Vite's `?url` import resolves to a
        // root-relative URL (e.g. `/node_modules/.../sql-wasm.wasm`). sql.js detects
        // Node and tries to load wasm via `fs.readFileSync`, so we must provide a
        // file:// URL instead of a server path.
        let resolved: string | null = null;
        try {
          if (typeof import.meta.resolve === "function") {
            resolved = import.meta.resolve(`sql.js/dist/${file}`);
          }
        } catch {
          // ignore
        }

        const candidate = resolved || sqlWasmUrl;
        return coerceViteUrlToNodeFileUrl(candidate);
      }

      // Preserve Emscripten's default locateFile behaviour.
      return prefix ? `${prefix}${file}` : file;
    },
    retention: { max_entries: params.retentionMaxEntries, max_age_ms: params.retentionMaxAgeMs },
  });
}

async function resolveDesktopAIAuditStore(options: DesktopAIAuditStoreOptions = {}): Promise<AIAuditStore> {
  const storageKey = options.storageKey ?? DESKTOP_AI_AUDIT_DB_STORAGE_KEY;
  const maxEntryChars = options.maxEntryChars;
  const retentionMaxEntries = options.retentionMaxEntries ?? DEFAULT_RETENTION_MAX_ENTRIES;
  const retentionMaxAgeMs = options.retentionMaxAgeMs ?? DEFAULT_RETENTION_MAX_AGE_MS;

  const cached = storePromiseByKey.get(storageKey);
  if (cached) return cached;

  const promise = createSqliteBackedStore({ storageKey, retentionMaxEntries, retentionMaxAgeMs }).catch((_err) => {
    // Best-effort fallback: keep audit logging functional even if sql.js fails to load
    // (e.g. blocked WASM fetch).
    return new LocalStorageAIAuditStore();
  }).then((store) => new BoundedAIAuditStore(store, maxEntryChars ? { max_entry_chars: maxEntryChars } : undefined));
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
