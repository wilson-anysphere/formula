import type { AIAuditEntry, AuditListFilters } from "./types.js";
import type { AIAuditStore } from "./store.js";

export class MemoryAIAuditStore implements AIAuditStore {
  private readonly entries: AIAuditEntry[] = [];

  async logEntry(entry: AIAuditEntry): Promise<void> {
    this.entries.push(cloneAuditEntry(entry));
  }

  async listEntries(filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    const { session_id, workbook_id, mode, limit } = filters;
    let results = session_id ? this.entries.filter((entry) => entry.session_id === session_id) : [...this.entries];
    if (workbook_id) {
      results = results.filter((entry) => entry.workbook_id === workbook_id);
    }
    if (mode) {
      const modes = Array.isArray(mode) ? mode : [mode];
      results = results.filter((entry) => modes.includes(entry.mode));
    }
    results.sort((a, b) => b.timestamp_ms - a.timestamp_ms);
    return typeof limit === "number" ? results.slice(0, limit) : results;
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
