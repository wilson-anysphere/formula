import type { AIAuditEntry, AuditListFilters } from "./types.ts";
import type { AIAuditStore } from "./store.ts";
import { stableStringify } from "./stable-json.ts";

export interface MemoryAIAuditStoreOptions {
  /**
   * Maximum number of entries to keep (newest retained). If unset, entries are unbounded.
   */
  max_entries?: number;
  /**
   * Maximum age in milliseconds. Entries older than (now - max_age_ms) are dropped
   * at write-time. If unset, age-based retention is disabled.
   */
  max_age_ms?: number;
}

export class MemoryAIAuditStore implements AIAuditStore {
  private readonly entries: AIAuditEntry[] = [];
  private readonly maxEntries?: number;
  private readonly maxAgeMs?: number;

  constructor(options: MemoryAIAuditStoreOptions = {}) {
    const maxEntries = options.max_entries;
    if (typeof maxEntries === "number" && Number.isFinite(maxEntries) && maxEntries > 0) {
      this.maxEntries = Math.floor(maxEntries);
    }

    const maxAgeMs = options.max_age_ms;
    if (typeof maxAgeMs === "number" && Number.isFinite(maxAgeMs) && maxAgeMs > 0) {
      this.maxAgeMs = maxAgeMs;
    }
  }

  async logEntry(entry: AIAuditEntry): Promise<void> {
    this.entries.push(cloneAuditEntry(entry));
    this.enforceRetention();
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
    let results = session_id ? this.entries.filter((entry) => entry.session_id === session_id) : [...this.entries];
    if (workbook_id) {
      results = results.filter((entry) => matchesWorkbookFilter(entry, workbook_id));
    }
    if (mode) {
      const modes = Array.isArray(mode) ? mode : [mode];
      if (modes.length > 0) {
        results = results.filter((entry) => modes.includes(entry.mode));
      }
    }
    if (typeof after_timestamp_ms === "number") {
      results = results.filter((entry) => entry.timestamp_ms >= after_timestamp_ms);
    }
    if (typeof before_timestamp_ms === "number") {
      results = results.filter((entry) => entry.timestamp_ms < before_timestamp_ms);
    }
    if (cursor) {
      results = results.filter((entry) => isEntryBeforeCursor(entry, cursor));
    }
    results.sort(compareEntriesNewestFirst);
    return typeof limit === "number" ? results.slice(0, limit) : results;
  }

  private enforceRetention(): void {
    const maxAgeMs = this.maxAgeMs;
    if (typeof maxAgeMs === "number" && Number.isFinite(maxAgeMs) && maxAgeMs > 0) {
      const cutoff = Date.now() - maxAgeMs;
      let writeIndex = 0;
      for (const entry of this.entries) {
        if (entry.timestamp_ms >= cutoff) {
          this.entries[writeIndex++] = entry;
        }
      }
      this.entries.length = writeIndex;
    }

    const maxEntries = this.maxEntries;
    if (
      typeof maxEntries === "number" &&
      Number.isFinite(maxEntries) &&
      maxEntries > 0 &&
      this.entries.length > maxEntries
    ) {
      this.entries.sort((a, b) => {
        if (a.timestamp_ms !== b.timestamp_ms) return a.timestamp_ms - b.timestamp_ms;
        return compareIdsAsc(a.id, b.id);
      });
      this.entries.splice(0, this.entries.length - maxEntries);
    }
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
