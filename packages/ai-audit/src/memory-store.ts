import type { AIAuditEntry, AuditListFilters } from "./types.js";
import type { AIAuditStore } from "./store.js";

export class MemoryAIAuditStore implements AIAuditStore {
  private readonly entries: AIAuditEntry[] = [];

  async logEntry(entry: AIAuditEntry): Promise<void> {
    this.entries.push(cloneAuditEntry(entry));
  }

  async listEntries(filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    const { session_id, limit } = filters;
    const results = session_id ? this.entries.filter((entry) => entry.session_id === session_id) : [...this.entries];
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
