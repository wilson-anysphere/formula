import type { AIAuditStore } from "./store.ts";
import type { AIAuditEntry, AuditListFilters } from "./types.ts";

export interface FailingAIAuditStoreOptions {
  error?: unknown;
}

/**
 * An audit store that always fails.
 *
 * Intended for tests that need to validate best-effort audit logging behavior.
 */
export class FailingAIAuditStore implements AIAuditStore {
  readonly error: unknown;

  constructor(errorOrOptions: Error | string | FailingAIAuditStoreOptions = {}) {
    const candidate =
      typeof errorOrOptions === "string" || errorOrOptions instanceof Error ? errorOrOptions : errorOrOptions.error;
    if (candidate === undefined) {
      this.error = new Error("FailingAIAuditStore: operation failed");
      return;
    }

    this.error = typeof candidate === "string" ? new Error(candidate) : candidate;
  }

  async logEntry(_entry: AIAuditEntry): Promise<void> {
    throw this.error;
  }

  async listEntries(_filters: AuditListFilters = {}): Promise<AIAuditEntry[]> {
    throw this.error;
  }
}
