import { describe, expect, it } from "vitest";

import { AIAuditRecorder } from "../src/recorder.js";
import { FailingAIAuditStore } from "../src/failing-store.js";
import { MemoryAIAuditStore } from "../src/memory-store.js";

describe("AIAuditRecorder", () => {
  it("accumulates token usage, tool call approvals, and persists via the store", async () => {
    const store = new MemoryAIAuditStore();
    const recorder = new AIAuditRecorder({
      store,
      session_id: "session-123",
      mode: "chat",
      input: { prompt: "Set A1 to 1" },
      model: "unit-test-model"
    });

    recorder.recordTokenUsage({ prompt_tokens: 10, completion_tokens: 5 });
    recorder.recordTokenUsage({ prompt_tokens: 2, completion_tokens: 3 });

    const callIndex = recorder.recordToolCall({
      id: "call-1",
      name: "write_cell",
      parameters: { cell: "Sheet1!A1", value: 1 },
      requires_approval: true
    });

    recorder.recordToolApproval("call-1", true);
    recorder.recordToolResult(callIndex, { ok: true, duration_ms: 12, result: { changed: true } });

    recorder.setUserFeedback("accepted");
    await recorder.finalize();

    const entries = await store.listEntries({ session_id: "session-123" });
    expect(entries.length).toBe(1);
    expect(entries[0]!.model).toBe("unit-test-model");
    expect(entries[0]!.token_usage).toEqual({ prompt_tokens: 12, completion_tokens: 8, total_tokens: 20 });
    expect(entries[0]!.tool_calls[0]).toMatchObject({
      name: "write_cell",
      requires_approval: true,
      approved: true,
      ok: true,
      duration_ms: 12
    });
    expect(entries[0]!.user_feedback).toBe("accepted");
  });

  it("finalize() is best-effort and records failures without throwing", async () => {
    const err = new Error("persist failed");
    const store = new FailingAIAuditStore(err);
    const recorder = new AIAuditRecorder({
      store,
      session_id: "session-fail",
      mode: "chat",
      input: { prompt: "hello" },
      model: "unit-test-model"
    });

    await expect(recorder.finalize()).resolves.toBeUndefined();
    expect(recorder.finalizeError).toBe(err);
    expect(recorder.getFinalizeError()).toBe(err);
    expect(recorder.finalize_error).toBe("persist failed");
    expect(typeof recorder.entry.latency_ms).toBe("number");
  });
});
