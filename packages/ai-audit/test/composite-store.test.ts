import { describe, expect, it, vi } from "vitest";

import { CompositeAIAuditStore } from "../src/composite-store.js";
import { FailingAIAuditStore } from "../src/failing-store.js";
import { MemoryAIAuditStore } from "../src/memory-store.js";
import type { AIAuditEntry } from "../src/types.js";

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
    const badStore = new FailingAIAuditStore(new Error("boom"));
    const logSpy = vi.spyOn(badStore, "logEntry");
    const composite = new CompositeAIAuditStore([badStore, okStore]);

    await expect(composite.logEntry(entry)).resolves.toBeUndefined();
    expect(logSpy).toHaveBeenCalledTimes(1);

    const okEntries = await okStore.listEntries({ session_id: "session-1" });
    expect(okEntries.map((e) => e.id)).toEqual(["entry-1"]);
  });
});
