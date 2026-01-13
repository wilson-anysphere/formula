import type { AIAuditStore } from "./store.ts";
import { LocalStorageAIAuditStore } from "./local-storage-store.ts";
import { MemoryAIAuditStore } from "./memory-store.ts";
import type { CreateDefaultAIAuditStoreOptions } from "./factory.ts";

/**
 * Node entrypoint implementation for `createDefaultAIAuditStore`.
 *
 * The Node entrypoint is allowed to create the sql.js-backed store on demand,
 * but still defaults to memory to avoid unnecessary overhead in hosts that
 * don't need persistence.
 */
export async function createDefaultAIAuditStore(options: CreateDefaultAIAuditStoreOptions = {}): Promise<AIAuditStore> {
  const retention = options.retention ?? {};
  const prefer = options.prefer ?? "memory";

  if (prefer === "memory") {
    return new MemoryAIAuditStore({ max_entries: retention.max_entries, max_age_ms: retention.max_age_ms });
  }

  if (prefer === "localstorage") {
    // Node runtimes rarely have a real localStorage, but allow explicit opt-in
    // (e.g. jsdom / experimental webstorage). If localStorage is unavailable,
    // fall back to memory.
    if (!isLocalStorageAvailable()) {
      return new MemoryAIAuditStore({ max_entries: retention.max_entries, max_age_ms: retention.max_age_ms });
    }
    return new LocalStorageAIAuditStore({ max_entries: retention.max_entries, max_age_ms: retention.max_age_ms });
  }

  // `prefer: "sqlite"` is explicit opt-in.
  const { SqliteAIAuditStore } = await import("./sqlite-store.ts");
  return SqliteAIAuditStore.create({
    storage: options.sqlite_storage,
    retention
  });
}

function isLocalStorageAvailable(): boolean {
  const storage = getLocalStorageOrNull();
  if (!storage) return false;

  const key = "__formula_ai_audit_probe__";
  try {
    const existing = storage.getItem(key);
    storage.setItem(key, "1");
    if (existing === null) storage.removeItem(key);
    else storage.setItem(key, existing);
    return true;
  } catch {
    return false;
  }
}

function getLocalStorageOrNull(): Storage | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const storage = (globalThis as any).localStorage as Storage | undefined;
    if (!storage) return null;
    if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
    return storage;
  } catch {
    return null;
  }
}
