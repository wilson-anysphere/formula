import { describe, expect, it } from "vitest";

import { FailingAIAuditStore, NoopAIAuditStore } from "../src/index.js";
import type { AIAuditEntry } from "../src/types.js";

function makeEntry(overrides: Partial<AIAuditEntry> = {}): AIAuditEntry {
  return {
    id: "entry-1",
    timestamp_ms: Date.now(),
    session_id: "session-1",
    mode: "chat",
    input: { prompt: "hello" },
    model: "unit-test-model",
    tool_calls: [],
    ...overrides,
  };
}

describe("NoopAIAuditStore", () => {
  it("accepts writes and always returns an empty list", async () => {
    const store = new NoopAIAuditStore();

    await expect(store.logEntry(makeEntry())).resolves.toBeUndefined();
    await expect(store.listEntries()).resolves.toEqual([]);
    await expect(store.listEntries({ session_id: "session-1" })).resolves.toEqual([]);
  });
});

describe("FailingAIAuditStore", () => {
  it("rejects both logEntry and listEntries with the configured error", async () => {
    const err = new Error("boom");
    const store = new FailingAIAuditStore({ error: err });

    await expect(store.logEntry(makeEntry())).rejects.toBe(err);
    await expect(store.listEntries()).rejects.toBe(err);
  });
});

