import { describe, expect, it } from "vitest";

import { FailingAIAuditStore, NoopAIAuditStore } from "../src/index.js";
import type { AIAuditEntry } from "../src/types.js";

function makeEntry(): AIAuditEntry {
  return {
    id: "entry-1",
    timestamp_ms: Date.now(),
    session_id: "session-1",
    mode: "chat",
    input: { prompt: "hello" },
    model: "unit-test-model",
    tool_calls: []
  };
}

describe("NoopAIAuditStore", () => {
  it("resolves logEntry and always returns an empty list", async () => {
    const store = new NoopAIAuditStore();
    const entry = makeEntry();

    await expect(store.logEntry(entry)).resolves.toBeUndefined();
    await expect(store.listEntries()).resolves.toEqual([]);
    await expect(store.listEntries({ session_id: entry.session_id })).resolves.toEqual([]);
  });
});

describe("FailingAIAuditStore", () => {
  it("rejects logEntry and listEntries with the provided Error", async () => {
    const error = new Error("boom");
    const store = new FailingAIAuditStore(error);
    const entry = makeEntry();

    await expect(store.logEntry(entry)).rejects.toBe(error);
    await expect(store.listEntries()).rejects.toBe(error);
  });

  it("accepts an error message and rejects with it", async () => {
    const store = new FailingAIAuditStore("audit store is failing");

    await expect(store.listEntries()).rejects.toThrow("audit store is failing");
  });
});
