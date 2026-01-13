import type { AIAuditEntry, AuditListFilters } from "./types.ts";
import type { AIAuditStore } from "./store.ts";

/**
 * An audit store that always fails.
 *
 * Intended for tests that need to validate best-effort audit logging behavior.
 */
export class FailingAIAuditStore implements AIAuditStore {
  private readonly error: Error;

  constructor(errorOrMessage: Error | string) {
    this.error = errorOrMessage instanceof Error ? errorOrMessage : new Error(errorOrMessage);
  }

  async logEntry(_entry: AIAuditEntry): Promise<void> {
    throw this.error;
  }

  async listEntries(_filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    throw this.error;
  }
}

