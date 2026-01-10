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
    expect(entries[0]!.user_feedback).toBe("accepted");
  });
});
