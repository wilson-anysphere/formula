import type { AIAuditStore } from "./store.ts";
import type { AIAuditEntry, AuditListFilters } from "./types.ts";
import { stableStringify } from "./stable-json.ts";

export interface LocalStorageAIAuditStoreOptions {
  /**
   * localStorage key used to persist entries. Defaults to `formula_ai_audit_log_entries`.
   */
  key?: string;
  /**
   * Cap the number of stored entries (oldest dropped). Defaults to 1000.
   *
   * Note: this caps the *count* of entries, but does not enforce a strict per-entry
   * size limit. For defense-in-depth against LocalStorage quota write failures,
   * wrap this store with `BoundedAIAuditStore`.
   */
  max_entries?: number;
  /**
   * Maximum age in milliseconds. Entries older than (now - max_age_ms) are dropped
   * at write-time and opportunistically on reads.
   *
   * If unset, age-based retention is disabled.
   */
  max_age_ms?: number;
}

export class LocalStorageAIAuditStore implements AIAuditStore {
  readonly key: string;
  readonly maxEntries: number;
  readonly maxAgeMs?: number;
  private readonly memoryFallback: AIAuditEntry[] = [];
  private localStorageUnavailable: boolean = false;

  constructor(options: LocalStorageAIAuditStoreOptions = {}) {
    this.key = options.key ?? "formula_ai_audit_log_entries";
    this.maxEntries = options.max_entries ?? 1000;
    this.maxAgeMs = options.max_age_ms;
  }

  async logEntry(entry: AIAuditEntry): Promise<void> {
    const nowMs = Date.now();
    let entries = this.loadEntries();
    entries.push(cloneAuditEntry(entry));
    entries.sort((a, b) => {
      if (a.timestamp_ms !== b.timestamp_ms) return a.timestamp_ms - b.timestamp_ms;
      return compareIdsAsc(a.id, b.id);
    });
    entries = this.enforceAgeRetention(entries, nowMs);
    while (entries.length > this.maxEntries) entries.shift();
    this.saveEntries(entries);
  }

  async listEntries(filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    const { session_id, mode } = filters;
    const workbook_id = typeof filters.workbook_id === "string" ? filters.workbook_id.trim() : "";
    const after_timestamp_ms =
      typeof filters.after_timestamp_ms === "number" && Number.isFinite(filters.after_timestamp_ms)
        ? filters.after_timestamp_ms
        : undefined;
    const before_timestamp_ms =
      typeof filters.before_timestamp_ms === "number" && Number.isFinite(filters.before_timestamp_ms)
        ? filters.before_timestamp_ms
        : undefined;
    const cursor =
      filters.cursor && typeof filters.cursor.before_timestamp_ms === "number" && Number.isFinite(filters.cursor.before_timestamp_ms)
        ? filters.cursor
        : undefined;
    const limit =
      typeof filters.limit === "number" && Number.isFinite(filters.limit) ? Math.max(0, Math.trunc(filters.limit)) : undefined;
    const nowMs = Date.now();
    const loaded = this.loadEntries();
    const entries = this.enforceAgeRetention(loaded, nowMs);
    if (entries.length !== loaded.length) {
      // Best-effort: persist purged entries so old data doesn't resurface if the app never logs again.
      this.saveEntries(entries);
    }
    let filtered = session_id ? entries.filter((entry) => entry.session_id === session_id) : entries.slice();
    if (workbook_id) {
      filtered = filtered.filter((entry) => matchesWorkbookFilter(entry, workbook_id));
    }
    if (mode) {
      const modes = Array.isArray(mode) ? mode : [mode];
      if (modes.length > 0) {
        filtered = filtered.filter((entry) => modes.includes(entry.mode));
      }
    }
    if (typeof after_timestamp_ms === "number") {
      filtered = filtered.filter((entry) => entry.timestamp_ms >= after_timestamp_ms);
    }
    if (typeof before_timestamp_ms === "number") {
      filtered = filtered.filter((entry) => entry.timestamp_ms < before_timestamp_ms);
    }
    if (cursor) {
      filtered = filtered.filter((entry) => isEntryBeforeCursor(entry, cursor));
    }
    filtered.sort(compareEntriesNewestFirst);
    return typeof limit === "number" ? filtered.slice(0, limit) : filtered;
  }

  private enforceAgeRetention(entries: AIAuditEntry[], nowMs: number): AIAuditEntry[] {
    const maxAgeMs = this.maxAgeMs;
    if (!(typeof maxAgeMs === "number" && Number.isFinite(maxAgeMs) && maxAgeMs > 0)) return entries;
    const cutoff = nowMs - maxAgeMs;
    if (!Number.isFinite(cutoff)) return entries;
    return entries.filter((entry) => entry.timestamp_ms >= cutoff);
  }

  private loadEntries(): AIAuditEntry[] {
    const storage = this.getLocalStorage();
    if (!storage) return this.memoryFallback.slice();
    try {
      const raw = storage.getItem(this.key);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? (parsed as AIAuditEntry[]) : [];
    } catch {
      // localStorage can throw in some environments (e.g. Node webstorage without a file,
      // Safari private mode). Fall back to in-memory storage.
      this.localStorageUnavailable = true;
      return this.memoryFallback.slice();
    }
  }

  private saveEntries(entries: AIAuditEntry[]): void {
    const storage = this.getLocalStorage();
    const snapshot = entries === this.memoryFallback ? entries.slice() : entries;
    if (!storage) {
      this.memoryFallback.length = 0;
      this.memoryFallback.push(...snapshot);
      return;
    }
    try {
      storage.setItem(this.key, stableStringify(snapshot));
    } catch {
      // If persistence fails, at least keep the latest entries in memory.
      this.localStorageUnavailable = true;
      this.memoryFallback.length = 0;
      this.memoryFallback.push(...snapshot);
    }
  }

  private getLocalStorage(): Storage | null {
    if (this.localStorageUnavailable) return null;
    const storage = getLocalStorageOrNull();
    // If localStorage exists but is inaccessible (e.g. Node's experimental webstorage
    // without a file path), avoid retrying (and throwing) on every call.
    if (!storage && typeof globalThis !== "undefined" && "localStorage" in globalThis) {
      this.localStorageUnavailable = true;
    }
    return storage;
  }
}

function compareEntriesNewestFirst(a: AIAuditEntry, b: AIAuditEntry): number {
  if (a.timestamp_ms !== b.timestamp_ms) return b.timestamp_ms - a.timestamp_ms;
  return compareIdsDesc(a.id, b.id);
}

function compareIdsDesc(aId: string | undefined, bId: string | undefined): number {
  const aVal = aId ?? "";
  const bVal = bId ?? "";
  if (aVal === bVal) return 0;
  return aVal < bVal ? 1 : -1;
}

function compareIdsAsc(aId: string | undefined, bId: string | undefined): number {
  const aVal = aId ?? "";
  const bVal = bId ?? "";
  if (aVal === bVal) return 0;
  return aVal < bVal ? -1 : 1;
}

function isEntryBeforeCursor(
  entry: AIAuditEntry,
  cursor: NonNullable<AuditListFilters["cursor"]>,
): boolean {
  if (entry.timestamp_ms < cursor.before_timestamp_ms) return true;
  if (entry.timestamp_ms > cursor.before_timestamp_ms) return false;
  const beforeId = typeof cursor.before_id === "string" ? cursor.before_id : undefined;
  if (!beforeId) return false;
  return compareIdsDesc(entry.id, beforeId) > 0;
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
      // ignore
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

function cloneAuditEntry(entry: AIAuditEntry): AIAuditEntry {
  const structuredCloneFn =
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    typeof (globalThis as any)?.structuredClone === "function" ? ((globalThis as any).structuredClone as any) : null;
  if (structuredCloneFn) {
    try {
      return structuredCloneFn(entry);
    } catch {
      // Fall back to JSON-based cloning.
    }
  }

  try {
    return JSON.parse(stableStringify(entry)) as AIAuditEntry;
  } catch {
    // Last resort: return as-is (audit entries should be JSON-serializable in practice).
    return entry;
  }
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
