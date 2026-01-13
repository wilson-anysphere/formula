import type { AIAuditEntry, AuditListFilters } from "./types.ts";
import type { AIAuditStore } from "./store.ts";

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
    const { session_id, workbook_id, mode, limit } = filters;
    let results = session_id ? this.entries.filter((entry) => entry.session_id === session_id) : [...this.entries];
    if (workbook_id) {
      results = results.filter((entry) => matchesWorkbookFilter(entry, workbook_id));
    }
    if (mode) {
      const modes = Array.isArray(mode) ? mode : [mode];
      results = results.filter((entry) => modes.includes(entry.mode));
    }
    results.sort((a, b) => b.timestamp_ms - a.timestamp_ms);
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
      this.entries.sort((a, b) => a.timestamp_ms - b.timestamp_ms);
      this.entries.splice(0, this.entries.length - maxEntries);
    }
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
    return JSON.parse(JSON.stringify(entry)) as AIAuditEntry;
  } catch {
    return entry;
  }
}

function matchesWorkbookFilter(entry: AIAuditEntry, workbookId: string): boolean {
  if (typeof entry.workbook_id === "string" && entry.workbook_id.trim()) {
    return entry.workbook_id === workbookId;
  }

  const input = entry.input as unknown;
  if (!input || typeof input !== "object") return false;
  const obj = input as Record<string, unknown>;
  const legacyWorkbookId = obj.workbook_id ?? obj.workbookId;
  return typeof legacyWorkbookId === "string" ? legacyWorkbookId === workbookId : false;
}
