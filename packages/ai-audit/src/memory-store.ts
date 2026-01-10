import type { AIAuditEntry, AuditListFilters } from "./types.js";
import type { AIAuditStore } from "./store.js";

export class MemoryAIAuditStore implements AIAuditStore {
  private readonly entries: AIAuditEntry[] = [];

  async logEntry(entry: AIAuditEntry): Promise<void> {
    this.entries.push(structuredClone(entry));
  }

  async listEntries(filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    const { session_id, limit } = filters;
    const results = session_id ? this.entries.filter((entry) => entry.session_id === session_id) : [...this.entries];
    results.sort((a, b) => b.timestamp_ms - a.timestamp_ms);
    return typeof limit === "number" ? results.slice(0, limit) : results;
  }
}

