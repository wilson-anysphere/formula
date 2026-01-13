import type { AIAuditStore } from "./store.ts";
import { BoundedAIAuditStore, type BoundedAIAuditStoreOptions } from "./bounded-store.ts";
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
  const max_entries = options.max_entries ?? retention.max_entries;
  const max_age_ms = options.max_age_ms ?? retention.max_age_ms;
  const prefer = options.prefer ?? "memory";
  const bounded = options.bounded;

  const wrap = (store: AIAuditStore): AIAuditStore => {
    if (bounded === false) return store;
    const boundedOptions: BoundedAIAuditStoreOptions | undefined = typeof bounded === "object" && bounded ? bounded : undefined;
    return new BoundedAIAuditStore(store, boundedOptions);
  };

  if (prefer === "memory") {
    return wrap(new MemoryAIAuditStore({ max_entries, max_age_ms }));
  }

  if (prefer === "indexeddb") {
    // Node runtimes do not have IndexedDB. Treat as unsupported and fall back to
    // localStorage if explicitly available, otherwise memory.
    if (isLocalStorageAvailable()) {
      return wrap(new LocalStorageAIAuditStore({ max_entries, max_age_ms }));
    }
    return wrap(new MemoryAIAuditStore({ max_entries, max_age_ms }));
  }

  if (prefer === "localstorage") {
    // Node runtimes rarely have a real localStorage, but allow explicit opt-in
    // (e.g. jsdom / experimental webstorage). If localStorage is unavailable,
    // fall back to memory.
    if (!isLocalStorageAvailable()) {
      return wrap(new MemoryAIAuditStore({ max_entries, max_age_ms }));
    }
    return wrap(new LocalStorageAIAuditStore({ max_entries, max_age_ms }));
  }

  // `prefer: "sqlite"` is explicit opt-in.
  const { SqliteAIAuditStore } = await import("./sqlite-store.ts");
  const sqlite = await SqliteAIAuditStore.create({
    storage: options.sqlite_storage,
    retention: { max_entries, max_age_ms }
  });
  return wrap(sqlite);
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

