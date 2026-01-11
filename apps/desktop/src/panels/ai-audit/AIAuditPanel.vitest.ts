import { describe, expect, it } from "vitest";

import { LocalStorageAIAuditStore } from "../../../../../packages/ai-audit/src/local-storage-store.js";
import type { AIAuditEntry } from "../../../../../packages/ai-audit/src/types.js";

import { createAIAuditPanel } from "./AIAuditPanel";

describe("AIAuditPanel", () => {
  it("renders entries from a LocalStorageAIAuditStore (most recent first)", async () => {
    const store = new LocalStorageAIAuditStore({ key: `ai_audit_test_${Math.random().toString(16).slice(2)}` });

    const older: AIAuditEntry = {
      id: "audit-older",
      timestamp_ms: 1700000000000,
      session_id: "session-1",
      mode: "chat",
      input: { message: "older" },
      model: "model-older",
      tool_calls: [{ name: "read_range", parameters: { range: "A1:A2" }, approved: true, ok: true }],
      token_usage: { prompt_tokens: 1, completion_tokens: 2, total_tokens: 3 },
      latency_ms: 10,
    };

    const newer: AIAuditEntry = {
      id: "audit-newer",
      timestamp_ms: 1700000005000,
      session_id: "session-1",
      mode: "chat",
      input: { message: "newer" },
      model: "model-newer",
      tool_calls: [{ name: "write_cell", parameters: { cell: "A1", value: 123 }, approved: true, ok: true }],
      token_usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 },
      latency_ms: 123,
    };

    await store.logEntry(older);
    await store.logEntry(newer);

    const container = document.createElement("div");
    document.body.appendChild(container);

    const panel = createAIAuditPanel({ container, store });
    await panel.ready;

    const entries = container.querySelectorAll('[data-testid="ai-audit-entry"]');
    expect(entries).toHaveLength(2);

    // Most recent entry first.
    expect(entries[0]?.textContent).toContain("model-newer");
    expect(entries[1]?.textContent).toContain("model-older");

    // Tool call details (name + approved/ok).
    const toolCalls = container.querySelectorAll('[data-testid="ai-audit-tool-call"]');
    expect(toolCalls.length).toBeGreaterThan(0);
    expect(toolCalls[0]?.textContent).toContain("approved:");
    expect(toolCalls[0]?.textContent).toContain("ok:");

    // Token usage + latency, if present.
    expect(container.textContent).toContain("Tokens:");
    expect(container.textContent).toContain("Latency:");
  });
});
