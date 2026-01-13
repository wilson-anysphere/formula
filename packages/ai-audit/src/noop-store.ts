import type { AIAuditEntry, AuditListFilters } from "./types.ts";
import type { AIAuditStore } from "./store.ts";

/**
 * An audit store that intentionally does nothing.
 *
 * Useful for hosts that want to disable persistence explicitly without sprinkling
 * null checks throughout their code, and for deterministic tests that need an
 * `AIAuditStore` instance but don't care about recorded output.
 */
export class NoopAIAuditStore implements AIAuditStore {
  async logEntry(_entry: AIAuditEntry): Promise<void> {
    // Intentionally noop.
  }

  async listEntries(_filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    return [];
  }
}
