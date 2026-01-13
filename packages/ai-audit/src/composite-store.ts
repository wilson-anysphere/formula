import type { AIAuditEntry, AuditListFilters } from "./types.ts";
import type { AIAuditStore } from "./store.ts";

export type CompositeAIAuditStoreMode = "all" | "best_effort";

export interface CompositeAIAuditStoreOptions {
  mode?: CompositeAIAuditStoreMode;
}

/**
 * AIAuditStore implementation that fans out writes to multiple stores.
 *
 * Modes:
 * - `all`: all stores must succeed, otherwise `logEntry` rejects.
 * - `best_effort` (default): attempt all stores; ignore partial failures
 *   and only reject if *every* store fails.
 */
export class CompositeAIAuditStore implements AIAuditStore {
  private readonly stores: readonly AIAuditStore[];
  private readonly mode: CompositeAIAuditStoreMode;

  constructor(stores: AIAuditStore[], opts: CompositeAIAuditStoreOptions = {}) {
    this.stores = stores;
    this.mode = opts.mode ?? "best_effort";
  }

  async logEntry(entry: AIAuditEntry): Promise<void> {
    if (this.stores.length === 0) {
      throw new Error("CompositeAIAuditStore: no underlying stores configured");
    }

    const results = await Promise.allSettled(this.stores.map(async (store) => store.logEntry(entry)));
    const errors = results
      .map((result) => (result.status === "rejected" ? result.reason : null))
      .filter((reason): reason is unknown => reason !== null);

    if (this.mode === "all") {
      if (errors.length > 0) {
        throw new AggregateError(errors, "CompositeAIAuditStore: failed to write entry to one or more stores");
      }
      return;
    }

    // best_effort: only surface a combined error if *every* store failed.
    if (errors.length === this.stores.length) {
      throw new AggregateError(errors, "CompositeAIAuditStore: failed to write entry to all stores");
    }
  }

  async listEntries(filters?: AuditListFilters): Promise<AIAuditEntry[]> {
    const primaryStore = this.stores[0];
    if (!primaryStore) return [];
    return primaryStore.listEntries(filters);
  }
}

