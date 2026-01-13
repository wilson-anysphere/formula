import { describe, expect, it } from "vitest";

import { CompositeAIAuditStore } from "../src/composite-store.js";
import { MemoryAIAuditStore } from "../src/memory-store.js";
import type { AIAuditEntry } from "../src/types.js";
import type { AIAuditStore } from "../src/store.js";

class ThrowingAIAuditStore implements AIAuditStore {
  calls = 0;

  async logEntry(_entry: AIAuditEntry): Promise<void> {
    this.calls += 1;
    throw new Error("boom");
  }

  async listEntries(): Promise<AIAuditEntry[]> {
    return [];
  }
}

describe("CompositeAIAuditStore", () => {
  const entry: AIAuditEntry = {
    id: "entry-1",
    timestamp_ms: Date.now(),
    session_id: "session-1",
    mode: "chat",
    input: { prompt: "hello" },
    model: "unit-test-model",
    tool_calls: []
  };

  it("fans out logEntry to all stores", async () => {
    const store1 = new MemoryAIAuditStore();
    const store2 = new MemoryAIAuditStore();
    const composite = new CompositeAIAuditStore([store1, store2]);

    await composite.logEntry(entry);

    const entries1 = await store1.listEntries({ session_id: "session-1" });
    const entries2 = await store2.listEntries({ session_id: "session-1" });

    expect(entries1.map((e) => e.id)).toEqual(["entry-1"]);
    expect(entries2.map((e) => e.id)).toEqual(["entry-1"]);
  });

  it("best_effort mode swallows individual failures when at least one store succeeds", async () => {
    const okStore = new MemoryAIAuditStore();
    const badStore = new ThrowingAIAuditStore();
    const composite = new CompositeAIAuditStore([badStore, okStore]);

    await expect(composite.logEntry(entry)).resolves.toBeUndefined();
    expect(badStore.calls).toBe(1);

    const okEntries = await okStore.listEntries({ session_id: "session-1" });
    expect(okEntries.map((e) => e.id)).toEqual(["entry-1"]);
  });
});

