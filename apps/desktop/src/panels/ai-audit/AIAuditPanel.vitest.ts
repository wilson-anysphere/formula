// @vitest-environment jsdom

import { describe, expect, it } from "vitest";

import type { AIAuditEntry } from "@formula/ai-audit/browser";
import { InMemoryBinaryStorage } from "@formula/ai-audit/browser";
import { SqliteAIAuditStore } from "@formula/ai-audit/sqlite";

import { createRequire } from "node:module";

import { createAIAuditPanel } from "./AIAuditPanel";

describe("AIAuditPanel", () => {
  it("renders entries from a SqliteAIAuditStore (most recent first) and filters by workbook_id", async () => {
    const require = createRequire(import.meta.url);
    const locateFile = (file: string) => require.resolve(`sql.js/dist/${file}`);

    const store = await SqliteAIAuditStore.create({ storage: new InMemoryBinaryStorage(), locateFile });

    const older: AIAuditEntry = {
      id: "audit-older",
      timestamp_ms: 1700000000000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { message: "older" },
      model: "model-older",
      tool_calls: [{ name: "read_range", parameters: { range: "A1:A2" }, approved: true, ok: true }],
      verification: {
        needs_tools: true,
        used_tools: true,
        verified: true,
        confidence: 0.9,
        warnings: [],
        claims: [{ claim: "mean(A1:A2) = 1.5", verified: true, expected: 1.5, actual: 1.5, toolEvidence: { tool: "compute_statistics" } }]
      },
      token_usage: { prompt_tokens: 1, completion_tokens: 2, total_tokens: 3 },
      latency_ms: 10,
    };

    const newer: AIAuditEntry = {
      id: "audit-newer",
      timestamp_ms: 1700000005000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { message: "newer" },
      model: "model-newer",
      tool_calls: [{ name: "write_cell", parameters: { cell: "A1", value: 123 }, approved: true, ok: true }],
      verification: {
        needs_tools: true,
        used_tools: false,
        verified: false,
        confidence: 0.2,
        warnings: ["No data tools were used; answer may be a guess."]
      },
      token_usage: { prompt_tokens: 10, completion_tokens: 5, total_tokens: 15 },
      latency_ms: 123,
    };

    const differentWorkbook: AIAuditEntry = {
      id: "audit-other-workbook",
      timestamp_ms: 1700000007000,
      session_id: "session-2",
      workbook_id: "workbook-2",
      mode: "chat",
      input: { message: "other workbook" },
      model: "model-other-workbook",
      tool_calls: [],
    };

    await store.logEntry(older);
    await store.logEntry(newer);
    await store.logEntry(differentWorkbook);

    const container = document.createElement("div");
    document.body.appendChild(container);

    const panel = createAIAuditPanel({ container, store, initialWorkbookId: "workbook-1" });
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

    // Verification details should be surfaced.
    expect(container.querySelectorAll('[data-testid="ai-audit-verification"]')).toHaveLength(2);
    expect(container.textContent).toContain("Verification:");
    expect(container.querySelectorAll('[data-testid="ai-audit-verification-claims"]')).toHaveLength(1);

    // Switching the workbook filter should update the results.
    const workbookInput = container.querySelector<HTMLInputElement>('[data-testid="ai-audit-filter-workbook"]');
    expect(workbookInput).toBeTruthy();
    if (!workbookInput) return;

    workbookInput.value = "workbook-2";
    await panel.refresh();

    const filteredEntries = container.querySelectorAll('[data-testid="ai-audit-entry"]');
    expect(filteredEntries).toHaveLength(1);
    expect(container.textContent).toContain("model-other-workbook");
  });

  it("supports basic pagination via limit + cursor (Load more)", async () => {
    const require = createRequire(import.meta.url);
    const locateFile = (file: string) => require.resolve(`sql.js/dist/${file}`);

    const store = await SqliteAIAuditStore.create({ storage: new InMemoryBinaryStorage(), locateFile });

    await store.logEntry({
      id: "entry-a",
      timestamp_ms: 1700000000000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { message: "a" },
      model: "model-a",
      tool_calls: []
    });
    await store.logEntry({
      id: "entry-b",
      timestamp_ms: 1700000001000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { message: "b" },
      model: "model-b",
      tool_calls: []
    });
    await store.logEntry({
      id: "entry-c",
      timestamp_ms: 1700000002000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { message: "c" },
      model: "model-c",
      tool_calls: []
    });

    const container = document.createElement("div");
    document.body.appendChild(container);

    const panel = createAIAuditPanel({ container, store, initialWorkbookId: "workbook-1" });
    await panel.ready;

    const pageSizeInput = container.querySelector<HTMLInputElement>('[data-testid="ai-audit-filter-page-size"]');
    expect(pageSizeInput).toBeTruthy();
    if (!pageSizeInput) return;

    pageSizeInput.value = "1";
    await panel.refresh();

    // First page: only the newest entry.
    let entries = container.querySelectorAll('[data-testid="ai-audit-entry"]');
    expect(entries).toHaveLength(1);
    expect(entries[0]?.textContent).toContain("model-c");

    // Next page should append older entries.
    await panel.loadMore();
    entries = container.querySelectorAll('[data-testid="ai-audit-entry"]');
    expect(entries).toHaveLength(2);
    expect(entries[0]?.textContent).toContain("model-c");
    expect(entries[1]?.textContent).toContain("model-b");

    await panel.loadMore();
    entries = container.querySelectorAll('[data-testid="ai-audit-entry"]');
    expect(entries).toHaveLength(3);
    expect(entries[2]?.textContent).toContain("model-a");
  });

  it("exports all matching entries (not just the currently loaded page)", async () => {
    const require = createRequire(import.meta.url);
    const locateFile = (file: string) => require.resolve(`sql.js/dist/${file}`);

    const store = await SqliteAIAuditStore.create({ storage: new InMemoryBinaryStorage(), locateFile });

    await store.logEntry({
      id: "entry-a",
      timestamp_ms: 1700000000000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { message: "a" },
      model: "model-a",
      tool_calls: []
    });
    await store.logEntry({
      id: "entry-b",
      timestamp_ms: 1700000001000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { message: "b" },
      model: "model-b",
      tool_calls: []
    });
    await store.logEntry({
      id: "entry-c",
      timestamp_ms: 1700000002000,
      session_id: "session-1",
      workbook_id: "workbook-1",
      mode: "chat",
      input: { message: "c" },
      model: "model-c",
      tool_calls: []
    });

    const container = document.createElement("div");
    document.body.appendChild(container);

    const panel = createAIAuditPanel({ container, store, initialWorkbookId: "workbook-1" });
    await panel.ready;

    const pageSizeInput = container.querySelector<HTMLInputElement>('[data-testid="ai-audit-filter-page-size"]');
    expect(pageSizeInput).toBeTruthy();
    if (!pageSizeInput) return;

    // Only load one entry into the UI state.
    pageSizeInput.value = "1";
    await panel.refresh();

    const visible = container.querySelectorAll('[data-testid="ai-audit-entry"]');
    expect(visible).toHaveLength(1);
    expect(visible[0]?.textContent).toContain("model-c");

    const exp = await panel.exportLog();
    expect(exp).toBeTruthy();
    if (!exp) return;

    const text = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result ?? ""));
      reader.onerror = () => reject(reader.error ?? new Error("FileReader failed"));
      reader.readAsText(exp.blob);
    });
    const lines = text.split("\n").filter(Boolean);
    expect(lines).toHaveLength(3);
    const parsed = lines.map((line) => JSON.parse(line) as AIAuditEntry);
    expect(parsed.map((e) => e.id)).toEqual(["entry-c", "entry-b", "entry-a"]);
  });
});
