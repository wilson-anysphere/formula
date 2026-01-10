import type { AIAuditEntry, AuditListFilters } from "./types.js";

export interface AIAuditStore {
  logEntry(entry: AIAuditEntry): Promise<void>;
  listEntries(filters?: AuditListFilters): Promise<AIAuditEntry[]>;
}
