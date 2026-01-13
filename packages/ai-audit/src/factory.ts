import type { AIAuditStore } from "./store.ts";
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
   * based on the runtime (browser vs Node).
   */
  prefer?: "sqlite" | "localstorage" | "memory";
  /**
   * Retention options forwarded to the underlying store when supported.
   */
  retention?: AIAuditStoreRetentionOptions;
  /**
   * Optional sqlite storage implementation used when `prefer: "sqlite"` in
   * Node runtimes (via `@formula/ai-audit/node`).
   */
  sqlite_storage?: SqliteBinaryStorage;
}

/**
 * Create an audit store with sensible defaults for the current runtime.
 *
 * Browser-like runtimes prefer `LocalStorageAIAuditStore` (fast + simple),
 * falling back to `MemoryAIAuditStore` when localStorage is unavailable.
 *
 * Node runtimes default to `MemoryAIAuditStore` to avoid pulling in sql.js unless
 * explicitly requested.
 */
export async function createDefaultAIAuditStore(options: CreateDefaultAIAuditStoreOptions = {}): Promise<AIAuditStore> {
  const retention = options.retention ?? {};
  const prefer = options.prefer ?? (isBrowserRuntime() ? "localstorage" : "memory");

  if (prefer === "memory") {
    return new MemoryAIAuditStore({ max_entries: retention.max_entries, max_age_ms: retention.max_age_ms });
  }

  if (prefer === "localstorage") {
    if (!isLocalStorageAvailable()) {
      return new MemoryAIAuditStore({ max_entries: retention.max_entries, max_age_ms: retention.max_age_ms });
    }
    return new LocalStorageAIAuditStore({ max_entries: retention.max_entries, max_age_ms: retention.max_age_ms });
  }

  // Keep the browser-safe entrypoint free of sql.js. Consumers that want sqlite
  // in browser-like environments should import `@formula/ai-audit/sqlite` and
  // construct `SqliteAIAuditStore` directly.
  throw new Error(
    'createDefaultAIAuditStore(prefer: "sqlite") is not available in the default/browser entrypoint. ' +
      'Import SqliteAIAuditStore from "@formula/ai-audit/sqlite" instead.'
  );
}

function isBrowserRuntime(): boolean {
  // Most reliable signal: `window` exists.
  // Avoid touching Node-only globals (process, Buffer) so the browser entrypoint
  // can be imported in constrained environments.
  return typeof window !== "undefined" && typeof window === "object";
}

function isLocalStorageAvailable(): boolean {
  const storage = getLocalStorageOrNull();
  if (!storage) return false;

  // Probe write access; some environments expose localStorage but throw on use
  // (e.g. private mode, restricted webviews).
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
