import type { AIAuditStore } from "../src/store.js";
import type { AIAuditEntry } from "../src/types.js";

/**
 * Test helper that always fails writes.
 */
export class FailingAIAuditStore implements AIAuditStore {
  readonly error: Error;

  constructor(error: Error = new Error("boom")) {
    this.error = error;
  }

  async logEntry(_entry: AIAuditEntry): Promise<void> {
    throw this.error;
  }

  async listEntries(): Promise<AIAuditEntry[]> {
    return [];
  }
}

