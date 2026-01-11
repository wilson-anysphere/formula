import { randomUUID } from "node:crypto";
import { describe, expect, it } from "vitest";
import { InMemoryBinaryStorage } from "../src/storage.js";
import { SqliteAIAuditStore } from "../src/sqlite-store.js";
import type { AIAuditEntry } from "../src/types.js";

describe("SqliteAIAuditStore", () => {
  it("persists entries to sqlite via a binary storage adapter", async () => {
    const storage = new InMemoryBinaryStorage();
    const store = await SqliteAIAuditStore.create({ storage });

    const entry: AIAuditEntry = {
      id: randomUUID(),
      timestamp_ms: Date.now(),
      session_id: "session-1",
      mode: "chat",
      input: { prompt: "hello" },
      model: "unit-test-model",
      token_usage: { prompt_tokens: 10, completion_tokens: 5 },
      latency_ms: 123,
      tool_calls: [
        {
          name: "write_cell",
          parameters: { cell: "Sheet1!A1", value: 1 },
          requires_approval: true,
          approved: true,
          ok: true,
          duration_ms: 7
        }
      ],
      verification: {
        needs_tools: true,
        used_tools: true,
        verified: false,
        confidence: 0.25,
        warnings: ["No data tools were used; answer may be a guess."]
      },
      user_feedback: "accepted"
    };

    await store.logEntry(entry);

    const roundTrip = await SqliteAIAuditStore.create({ storage });
    const entries = await roundTrip.listEntries({ session_id: "session-1" });

    expect(entries.length).toBe(1);
    expect(entries[0]!.id).toBe(entry.id);
    expect(entries[0]!.model).toBe("unit-test-model");
    expect(entries[0]!.token_usage?.total_tokens).toBe(15);
    expect(entries[0]!.tool_calls[0]?.name).toBe("write_cell");
    expect(entries[0]!.tool_calls[0]?.approved).toBe(true);
    expect(entries[0]!.verification).toEqual(entry.verification);
    expect(entries[0]!.user_feedback).toBe("accepted");
  });

  it("supports workbook/mode filtering and enforces retention at write-time", async () => {
    const storage = new InMemoryBinaryStorage();
    const store = await SqliteAIAuditStore.create({
      storage,
      retention: { max_entries: 2 }
    });

    const baseTime = Date.now();
    await store.logEntry({
      id: randomUUID(),
      timestamp_ms: baseTime - 3000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { prompt: "older" },
      model: "unit-test-model",
      tool_calls: []
    });

    const keptInlineEditId = randomUUID();
    await store.logEntry({
      id: keptInlineEditId,
      timestamp_ms: baseTime - 2000,
      session_id: "session-2",
      workbook_id: "workbook-1",
      mode: "inline_edit",
      input: { prompt: "mid" },
      model: "unit-test-model",
      tool_calls: []
    });

    const keptChatId = randomUUID();
    await store.logEntry({
      id: keptChatId,
      timestamp_ms: baseTime - 1000,
      session_id: "session-3",
      workbook_id: "workbook-2",
      mode: "chat",
      input: { prompt: "newest" },
      model: "unit-test-model",
      tool_calls: []
    });

    // max_entries=2 should have trimmed the oldest row.
    const all = await store.listEntries();
    expect(all.map((e) => e.id)).toEqual([keptChatId, keptInlineEditId]);

    const workbook1 = await store.listEntries({ workbook_id: "workbook-1" });
    expect(workbook1.map((e) => e.id)).toEqual([keptInlineEditId]);

    const chatOnly = await store.listEntries({ mode: "chat" });
    expect(chatOnly.map((e) => e.id)).toEqual([keptChatId]);

    const multiMode = await store.listEntries({ mode: ["chat", "inline_edit"] });
    expect(multiMode.map((e) => e.id)).toEqual([keptChatId, keptInlineEditId]);
  });

  it("supports max_age_ms retention", async () => {
    const storage = new InMemoryBinaryStorage();
    const store = await SqliteAIAuditStore.create({
      storage,
      retention: { max_age_ms: 1000 }
    });

    const oldId = randomUUID();
    await store.logEntry({
      id: oldId,
      timestamp_ms: Date.now() - 10_000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { prompt: "too old" },
      model: "unit-test-model",
      tool_calls: []
    });

    const freshId = randomUUID();
    await store.logEntry({
      id: freshId,
      timestamp_ms: Date.now(),
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { prompt: "fresh" },
      model: "unit-test-model",
      tool_calls: []
    });

    const entries = await store.listEntries();
    expect(entries.map((e) => e.id)).toEqual([freshId]);
  });
});
