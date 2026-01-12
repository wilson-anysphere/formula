import type { AIAuditEntry, AuditListFilters } from "./types.ts";

export interface AIAuditStore {
  logEntry(entry: AIAuditEntry): Promise<void>;
  listEntries(filters?: AuditListFilters): Promise<AIAuditEntry[]>;
}
