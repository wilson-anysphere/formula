import { describe, expect, it, vi } from "vitest";

import { MemoryAIAuditStore } from "../src/memory-store.js";
import type { AIAuditEntry } from "../src/types.js";

function makeEntry(id: string, timestamp_ms: number): AIAuditEntry {
  return {
    id,
    timestamp_ms,
    session_id: "session-1",
    mode: "chat",
    input: { prompt: "hello" },
    model: "unit-test-model",
    tool_calls: []
  };
}

describe("MemoryAIAuditStore retention", () => {
  it("enforces max_entries (keeps newest entries by timestamp)", async () => {
    const store = new MemoryAIAuditStore({ max_entries: 2 });

    await store.logEntry(makeEntry("entry-3-first", 300));
    await store.logEntry(makeEntry("entry-1-second", 100));
    await store.logEntry(makeEntry("entry-2-third", 200));

    const entries = await store.listEntries();
    expect(entries.map((entry) => entry.id)).toEqual(["entry-3-first", "entry-2-third"]);
  });

  it("enforces max_entries deterministically when timestamps tie (id tiebreaker)", async () => {
    const store = new MemoryAIAuditStore({ max_entries: 2 });

    // Insert in an order that would previously make retention depend on insertion order.
    await store.logEntry(makeEntry("c", 1000));
    await store.logEntry(makeEntry("a", 1000));
    await store.logEntry(makeEntry("b", 1000));

    const entries = await store.listEntries();
    expect(entries.map((entry) => entry.id)).toEqual(["c", "b"]);
  });

  it("enforces max_age_ms (drops entries older than Date.now() - max_age_ms)", async () => {
    vi.useFakeTimers();
    try {
      vi.setSystemTime(100_000);

      const store = new MemoryAIAuditStore({ max_age_ms: 10_000 });
      await store.logEntry(makeEntry("recent", 100_000));

      // Advance beyond the max age and log a new entry to trigger retention.
      vi.setSystemTime(115_000);
      await store.logEntry(makeEntry("new", 115_000));

      const entries = await store.listEntries();
      expect(entries.map((entry) => entry.id)).toEqual(["new"]);
    } finally {
      vi.useRealTimers();
    }
  });
});
