import { LocalStorageAIAuditStore } from "./local-storage-store.ts";
import type { SqliteAIAuditStore } from "./sqlite-store.ts";
import type { AuditListFilters } from "./types.ts";

export interface MigrateLocalStorageAuditEntriesToSqliteOptions {
  /**
   * Source audit entries stored as a JSON array in localStorage.
   *
   * Pass an existing LocalStorageAIAuditStore instance, or a `{ key }` pointing
   * at legacy entries stored under a custom key.
   */
  source: LocalStorageAIAuditStore | { key: string };
  /**
   * Destination sqlite-backed audit store.
   */
  destination: SqliteAIAuditStore;
  /**
   * When true, remove the source localStorage key only after a successful
   * migration pass.
   *
   * Default: false.
   */
  delete_source?: boolean;
  /**
   * Optional safety cap on how many entries to migrate (newest first, matching
   * the store's `listEntries()` semantics).
   *
   * Default: migrate all available entries.
   */
  max_entries?: number;
}

/**
 * Migrate audit entries from {@link LocalStorageAIAuditStore} (stored as a JSON
 * array under a localStorage key) into a {@link SqliteAIAuditStore}.
 *
 * The migration is idempotent: re-running it will not duplicate rows. Entries
 * that already exist in the destination store (same primary key `id`) are
 * skipped.
 */
export async function migrateLocalStorageAuditEntriesToSqlite(
  opts: MigrateLocalStorageAuditEntriesToSqliteOptions
): Promise<void> {
  const deleteSource = opts.delete_source ?? false;
  const maxEntries = opts.max_entries;

  const sourceKey = opts.source.key;
  const sourceStore =
    isLocalStorageAIAuditStore(opts.source) ? opts.source : new LocalStorageAIAuditStore({ key: sourceKey });

  const filters: AuditListFilters = {};
  if (typeof maxEntries === "number") filters.limit = maxEntries;

  const entries = await sourceStore.listEntries(filters);
  // Insert oldest-to-newest to preserve a stable write order (especially if the
  // destination store enforces retention policies at write-time).
  entries.sort((a, b) => a.timestamp_ms - b.timestamp_ms);

  for (const entry of entries) {
    try {
      await opts.destination.logEntry(entry);
    } catch (err) {
      if (isDuplicatePrimaryKeyError(err)) continue;
      throw err;
    }
  }

  if (deleteSource) {
    clearLocalStorageKey(sourceKey);
  }
}

function isLocalStorageAIAuditStore(source: LocalStorageAIAuditStore | { key: string }): source is LocalStorageAIAuditStore {
  return typeof (source as LocalStorageAIAuditStore).listEntries === "function";
}

function isDuplicatePrimaryKeyError(err: unknown): boolean {
  const message = getErrorMessage(err);
  // sql.js uses SQLite's default error strings (e.g. "UNIQUE constraint failed: ai_audit_log.id").
  return /UNIQUE constraint failed:\s*ai_audit_log\.id/i.test(message);
}

function getErrorMessage(err: unknown): string {
  if (!err) return "";
  if (typeof err === "string") return err;
  if (typeof err === "object" && "message" in err && typeof (err as any).message === "string") {
    return (err as any).message as string;
  }
  try {
    return String(err);
  } catch {
    return "";
  }
}

function clearLocalStorageKey(key: string): void {
  const storage = safeLocalStorage();
  if (!storage) return;

  try {
    storage.removeItem(key);
    return;
  } catch {
    // Some environments throw on removeItem (e.g. restricted storage). Fall back
    // to overwriting with an empty log array.
  }

  try {
    storage.setItem(key, "[]");
  } catch {
    // Ignore failures: migration succeeded and is idempotent even if we cannot
    // delete the source key.
  }
}

function safeLocalStorage(): Storage | null {
  try {
    const storage = globalThis.localStorage;
    if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") return storage;
  } catch {
    // ignore
  }

  try {
    const storage = (globalThis as any).window?.localStorage as Storage | undefined;
    if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") return storage;
  } catch {
    // ignore
  }

  return null;
}

