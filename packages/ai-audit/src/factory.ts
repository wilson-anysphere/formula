import type { AIAuditStore } from "./store.ts";
import { BoundedAIAuditStore, type BoundedAIAuditStoreOptions } from "./bounded-store.ts";
import { IndexedDbAIAuditStore } from "./indexeddb-store.ts";
import { LocalStorageAIAuditStore } from "./local-storage-store.ts";
import { MemoryAIAuditStore } from "./memory-store.ts";
import type { SqliteBinaryStorage } from "./storage.ts";

export interface AIAuditStoreRetentionOptions {
  max_entries?: number;
  max_age_ms?: number;
}

export interface CreateDefaultAIAuditStoreOptions {
  /**
   * Prefer a specific backing store. When unset, a reasonable default is selected
   * based on the runtime (browser/webview vs Node).
   */
  prefer?: "sqlite" | "indexeddb" | "localstorage" | "memory";
  /**
   * Wrap the chosen store in `BoundedAIAuditStore` (defense-in-depth against
   * oversized entries / quota overruns).
   *
   * Defaults to enabled.
   */
  bounded?: boolean | BoundedAIAuditStoreOptions;
  /**
   * Retention options forwarded to the underlying store when supported.
   *
   * Deprecated: prefer specifying `max_entries` / `max_age_ms` at the top level.
   */
  retention?: AIAuditStoreRetentionOptions;
  /**
   * Maximum number of entries to retain (newest retained).
   */
  max_entries?: number;
  /**
   * Maximum age in milliseconds. Entries older than (now - max_age_ms) are dropped
   * or deleted at write-time (depending on store implementation).
   */
  max_age_ms?: number;
  /**
   * Optional sqlite storage implementation used when `prefer: "sqlite"` in
   * Node runtimes (via `@formula/ai-audit/node`).
   *
   * This option is accepted here so callsites can share a single options object
   * across browser + node entrypoints.
   */
  sqlite_storage?: SqliteBinaryStorage;
}

/**
 * Create an audit store with sensible defaults for the current runtime.
 *
 * Browser-like runtimes prefer `IndexedDbAIAuditStore` when available, falling
 * back to `LocalStorageAIAuditStore` and finally `MemoryAIAuditStore`.
 *
 * Node runtimes default to `MemoryAIAuditStore` to avoid pulling in sql.js unless
 * explicitly requested (via the Node entrypoint).
 */
export async function createDefaultAIAuditStore(options: CreateDefaultAIAuditStoreOptions = {}): Promise<AIAuditStore> {
  const retention = options.retention ?? {};
  const max_entries = options.max_entries ?? retention.max_entries;
  const max_age_ms = options.max_age_ms ?? retention.max_age_ms;
  const prefer = options.prefer;
  const bounded = options.bounded;

  const wrap = (store: AIAuditStore): AIAuditStore => {
    if (bounded === false) return store;
    const boundedOptions = typeof bounded === "object" && bounded ? bounded : undefined;
    return new BoundedAIAuditStore(store, boundedOptions);
  };

  const createMemory = (): AIAuditStore => new MemoryAIAuditStore({ max_entries, max_age_ms });
  const createLocalStorage = (): AIAuditStore => new LocalStorageAIAuditStore({ max_entries, max_age_ms });

  const createIndexedDb = async (): Promise<AIAuditStore | null> => {
    if (!isIndexedDbAvailable()) return null;
    try {
      const store = new IndexedDbAIAuditStore({ max_entries, max_age_ms });
      // Probe once so we fail fast (blocked/denied IndexedDB) and can fall back.
      await store.listEntries({ limit: 0 });
      return store;
    } catch {
      return null;
    }
  };

  if (prefer === "memory") {
    return wrap(createMemory());
  }

  if (prefer === "localstorage") {
    if (!isLocalStorageAvailable()) return wrap(createMemory());
    return wrap(createLocalStorage());
  }

  if (prefer === "indexeddb") {
    const indexed = await createIndexedDb();
    if (indexed) return wrap(indexed);
    if (isLocalStorageAvailable()) return wrap(createLocalStorage());
    return wrap(createMemory());
  }

  if (prefer === "sqlite") {
    // Keep the browser-safe entrypoint free of sql.js. Consumers that want sqlite
    // in browser-like environments should import `@formula/ai-audit/sqlite` and
    // construct `SqliteAIAuditStore` directly.
    throw new Error(
      'createDefaultAIAuditStore(prefer: "sqlite") is not available in the default/browser entrypoint. ' +
        'Import SqliteAIAuditStore from "@formula/ai-audit/sqlite" instead.'
    );
  }

  // Automatic defaults.
  //
  // Prefer IndexedDB when it's available (even in Node test environments via
  // `fake-indexeddb`), otherwise fall back to a browser-appropriate localStorage
  // store when usable.
  const indexed = await createIndexedDb();
  if (indexed) return wrap(indexed);
  if (isNodeRuntime()) return wrap(createMemory());
  if (isLocalStorageAvailable()) return wrap(createLocalStorage());
  return wrap(createMemory());
}

function isNodeRuntime(): boolean {
  // Treat webviews/Electron renderers as browser-like when `window` exists.
  if (typeof window !== "undefined") return false;

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const versions = typeof process !== "undefined" ? ((process as any).versions as any) : undefined;
    return !!versions?.node;
  } catch {
    return false;
  }
}

function isIndexedDbAvailable(): boolean {
  try {
    if (typeof globalThis === "undefined") return false;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const idb = (globalThis as any).indexedDB as IDBFactory | undefined;
    return !!idb && typeof idb.open === "function";
  } catch {
    return false;
  }
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
  // Prefer `window.localStorage` when available (standard browser case).
  if (typeof window !== "undefined") {
    try {
      const storage = window.localStorage;
      if (!storage) return null;
      if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
      return storage;
    } catch {
      // If `window.localStorage` exists but is inaccessible (Safari private mode,
      // restricted webviews), treat it as unavailable. Do not fall back to
      // `globalThis.localStorage` because browsers alias it to `window.localStorage`,
      // and Node environments can expose an unrelated implementation.
      return null;
    }
  }

  try {
    if (typeof globalThis === "undefined") return null;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const storage = (globalThis as any).localStorage as Storage | undefined;
    if (!storage) return null;
    if (typeof storage.getItem !== "function" || typeof storage.setItem !== "function") return null;
    return storage;
  } catch {
    return null;
  }
}
