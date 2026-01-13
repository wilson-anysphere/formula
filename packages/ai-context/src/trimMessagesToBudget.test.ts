import { describe, expect, it } from "vitest";

import { createHeuristicTokenEstimator } from "./tokenBudget.js";
import { CONTEXT_SUMMARY_MARKER, trimMessagesToBudget } from "./trimMessagesToBudget.js";

describe("trimMessagesToBudget", () => {
  it("preserves system messages and summarizes dropped history", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

    const messages = [
      { role: "system", content: "system-1" },
      { role: "system", content: "system-2" },
      ...Array.from({ length: 10 }, (_, i) => ({ role: "user", content: `user-${i}-` + "x".repeat(50) })),
      { role: "assistant", content: "latest assistant" }
    ];

    const trimmed = await trimMessagesToBudget({
      messages,
      maxTokens: 400,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 3,
      summaryMaxTokens: 100
    });

    expect(trimmed[0]).toMatchObject({ role: "system", content: "system-1" });
    expect(trimmed[1]).toMatchObject({ role: "system", content: "system-2" });

    const summary = trimmed.find((m) => m.role === "system" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER));
    expect(summary).toBeTruthy();

    // Should keep the most recent non-system messages.
    expect(trimmed.at(-1)).toMatchObject({ role: "assistant", content: "latest assistant" });
    expect(trimmed.filter((m) => m.role !== "system")).toHaveLength(3);

    expect(estimator.estimateMessagesTokens(trimmed)).toBeLessThanOrEqual(400);
  });

  it("returns messages unchanged when already under budget", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });
    const messages = [
      { role: "system", content: "sys" },
      { role: "user", content: "hello" },
      { role: "assistant", content: "hi" }
    ];

    const trimmed = await trimMessagesToBudget({ messages, maxTokens: 100, reserveForOutputTokens: 0, estimator });
    expect(trimmed).toEqual(messages);
  });

  it("replaces existing generated summary messages instead of accumulating them", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

    const messages = [
      { role: "system", content: "sys" },
      { role: "system", content: `${CONTEXT_SUMMARY_MARKER}\nOld summary` },
      ...Array.from({ length: 8 }, (_, i) => ({ role: "user", content: `user-${i}-` + "y".repeat(60) })),
      { role: "assistant", content: "final" }
    ];

    const trimmed = await trimMessagesToBudget({
      messages,
      maxTokens: 300,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 2,
      summaryMaxTokens: 80
    });

    const summaryMessages = trimmed.filter(
      (m) => m.role === "system" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER)
    );
    expect(summaryMessages).toHaveLength(1);
    expect(estimator.estimateMessagesTokens(trimmed)).toBeLessThanOrEqual(300);
  });

  it("replaces existing generated summary messages when summaryRole is assistant", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

    const messages = [
      { role: "system", content: "sys" },
      { role: "assistant", content: `${CONTEXT_SUMMARY_MARKER}\nOld summary` },
      ...Array.from({ length: 8 }, (_, i) => ({ role: "user", content: `user-${i}-` + "y".repeat(60) })),
      { role: "assistant", content: "final" }
    ];

    const trimmed = await trimMessagesToBudget({
      messages,
      maxTokens: 300,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 2,
      summaryMaxTokens: 80,
      summaryRole: "assistant"
    });

    const summaryMessages = trimmed.filter(
      (m) => m.role === "assistant" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER)
    );
    expect(summaryMessages).toHaveLength(1);
    expect(estimator.estimateMessagesTokens(trimmed)).toBeLessThanOrEqual(300);
  });

  it("respects AbortSignal (already aborted)", async () => {
    const abortController = new AbortController();
    abortController.abort();

    await expect(
      trimMessagesToBudget({
        messages: [{ role: "user", content: "hello" }],
        maxTokens: 10,
        reserveForOutputTokens: 0,
        signal: abortController.signal
      })
    ).rejects.toMatchObject({ name: "AbortError" });
  });

  it("aborts while awaiting summarize callback", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });
    const abortController = new AbortController();

    let resolveStarted: (() => void) | null = null;
    const started = new Promise<void>((resolve) => {
      resolveStarted = resolve;
    });

    const promise = trimMessagesToBudget({
      messages: [
        { role: "system", content: "sys" },
        ...Array.from({ length: 10 }, (_v, i) => ({ role: "user", content: `user-${i}-` + "x".repeat(50) })),
        { role: "assistant", content: "latest assistant" }
      ],
      maxTokens: 200,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 1,
      summaryMaxTokens: 50,
      summarize: async () => {
        resolveStarted?.();
        return new Promise(() => {});
      },
      signal: abortController.signal
    });

    await started;
    abortController.abort();
    await expect(promise).rejects.toMatchObject({ name: "AbortError" });
  });

  it("preserves tool-call coherence (no orphan tool messages) when trimming", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

    const messages = [
      { role: "system", content: "sys" },
      { role: "user", content: "older-" + "x".repeat(200) },
      {
        role: "assistant",
        content: "",
        toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "foo" } }]
      },
      { role: "tool", toolCallId: "call-1", content: "tool-result" },
      { role: "assistant", content: "used tool output" },
      { role: "user", content: "latest question" }
    ];

    const trimmed = await trimMessagesToBudget({
      messages,
      maxTokens: 220,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 3,
      summaryMaxTokens: 50
    });

    expect(trimmed.some((m) => m.role === "tool")).toBe(true);

    // Any kept tool message must be preceded by an assistant toolCalls message.
    for (let i = 0; i < trimmed.length; i += 1) {
      const msg = trimmed[i];
      if (msg.role !== "tool") continue;

      let j = i - 1;
      while (j >= 0 && trimmed[j]?.role === "tool") j -= 1;
      expect(j).toBeGreaterThanOrEqual(0);
      expect(trimmed[j]).toMatchObject({ role: "assistant" });
      expect(Array.isArray(trimmed[j].toolCalls)).toBe(true);
      expect(trimmed[j].toolCalls.map((c: any) => c.id)).toContain(msg.toolCallId);
    }

    // Any kept assistant toolCalls message should have its tool messages kept when present.
    for (let i = 0; i < trimmed.length; i += 1) {
      const msg = trimmed[i];
      if (msg.role !== "assistant" || !Array.isArray(msg.toolCalls) || msg.toolCalls.length === 0) continue;
      const expectedIds = new Set(msg.toolCalls.map((c: any) => c.id));
      const seen = new Set<string>();
      for (let j = i + 1; j < trimmed.length; j += 1) {
        if (trimmed[j].role !== "tool") break;
        seen.add(trimmed[j].toolCallId);
      }
      for (const id of expectedIds) expect(seen.has(id)).toBe(true);
    }

    expect(estimator.estimateMessagesTokens(trimmed)).toBeLessThanOrEqual(220);
  });

  it("drops orphan tool messages when trimming with preserveToolCallPairs", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

    const messages = [
      { role: "system", content: "sys" },
      { role: "user", content: "older-" + "x".repeat(200) },
      // Orphan tool message (no preceding assistant toolCalls message).
      { role: "tool", toolCallId: "call-1", content: "orphan-tool-result" },
      { role: "user", content: "latest question" }
    ];

    const trimmed = await trimMessagesToBudget({
      messages,
      maxTokens: 200,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 2,
      summaryMaxTokens: 180
    });

    expect(trimmed.some((m) => m.role === "tool")).toBe(false);

    const summary = trimmed.find((m) => typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER));
    expect(summary).toBeTruthy();
    expect(String(summary?.content)).toContain("orphan-tool-result");
  });

  it("drops tool messages whose toolCallId doesn't match the preceding assistant toolCalls", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

    const messages = [
      { role: "system", content: "sys" },
      { role: "user", content: "older-" + "x".repeat(400) },
      {
        role: "assistant",
        content: "",
        toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "foo" } }]
      },
      // Mismatched tool result (should be treated as orphan).
      { role: "tool", toolCallId: "call-2", content: "wrong-tool-result" },
      // Correct tool result for call-1.
      { role: "tool", toolCallId: "call-1", content: "right-tool-result" },
      { role: "assistant", content: "after tool" },
      { role: "user", content: "latest" }
    ];

    const trimmed = await trimMessagesToBudget({
      messages,
      maxTokens: 260,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 10,
      summaryMaxTokens: 120
    });

    expect(trimmed.some((m) => m.role === "tool" && m.toolCallId === "call-2")).toBe(false);
    expect(trimmed.some((m) => m.role === "tool" && m.toolCallId === "call-1")).toBe(true);
  });

  it("can disable tool-call coherence (preserveToolCallPairs=false)", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

    const messages = [
      { role: "system", content: "sys" },
      {
        role: "assistant",
        content: "",
        // Large tool call payload so it gets dropped before the tool message under tight budgets.
        toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "x".repeat(200) } }]
      },
      { role: "tool", toolCallId: "call-1", content: "tool-result" },
      { role: "assistant", content: "ok" },
      { role: "user", content: "latest" }
    ];

    const trimmed = await trimMessagesToBudget({
      messages,
      maxTokens: 120,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 10,
      summaryMaxTokens: 30,
      preserveToolCallPairs: false
    });

    // With tool-call pairing disabled, we may keep a tool message even if its originating assistant
    // toolCalls message was dropped into the summary.
    const toolIdx = trimmed.findIndex((m) => m.role === "tool");
    expect(toolIdx).toBeGreaterThanOrEqual(0);
    const hasAssistantToolCalls = trimmed.some((m) => m.role === "assistant" && Array.isArray(m.toolCalls) && m.toolCalls.length > 0);
    expect(hasAssistantToolCalls).toBe(false);
  });

  it("summarizes a dropped tool-call group instead of keeping a partial group under tight budgets", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

    const messages = [
      { role: "system", content: "sys" },
      {
        role: "assistant",
        content: "",
        toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "x".repeat(200) } }]
      },
      { role: "tool", toolCallId: "call-1", content: "tool-result" },
      { role: "assistant", content: "after tool" },
      { role: "user", content: "latest" }
    ];

    const trimmed = await trimMessagesToBudget({
      messages,
      maxTokens: 220,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 10,
      summaryMaxTokens: 170
    });

    // Tool call group should be summarized away, not partially kept.
    expect(trimmed.some((m) => m.role === "tool")).toBe(false);
    expect(trimmed.some((m) => m.role === "assistant" && Array.isArray(m.toolCalls) && m.toolCalls.length > 0)).toBe(false);

    const summary = trimmed.find((m) => m.role === "system" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER));
    expect(summary).toBeTruthy();
    expect(String(summary?.content)).toContain("tool_calls");
  });

  it("dropToolMessagesFirst prefers dropping completed tool-call groups to keep more chat context", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

    const messages = [
      { role: "system", content: "sys" },
      { role: "user", content: "u1" },
      { role: "assistant", content: "a1" },
      {
        role: "assistant",
        content: "",
        toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "x".repeat(250) } }]
      },
      { role: "tool", toolCallId: "call-1", content: "y".repeat(250) },
      { role: "assistant", content: "a2" },
      { role: "user", content: "u2" }
    ];

    const without = await trimMessagesToBudget({
      messages,
      maxTokens: 200,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 50,
      summaryMaxTokens: 50
    });

    const withDropFirst = await trimMessagesToBudget({
      messages,
      maxTokens: 200,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 50,
      summaryMaxTokens: 50,
      dropToolMessagesFirst: true
    });

    // Without dropToolMessagesFirst, we should lose the early chat context due to the expensive tool messages.
    expect(without.some((m) => m.role === "user" && m.content === "u1")).toBe(false);
    // With dropToolMessagesFirst, tool call group should be summarized/dropped first, preserving early chat context.
    expect(withDropFirst.some((m) => m.role === "user" && m.content === "u1")).toBe(true);
    expect(withDropFirst.some((m) => m.role === "tool")).toBe(false);
  });

  it("falls back to a summary when the most recent tool-call group cannot fit at all", async () => {
    const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

    const messages = [
      { role: "system", content: "sys" },
      {
        role: "assistant",
        content: "",
        toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "x".repeat(1000) } }]
      },
      { role: "tool", toolCallId: "call-1", content: "tool-result" }
    ];

    const trimmed = await trimMessagesToBudget({
      messages,
      maxTokens: 200,
      reserveForOutputTokens: 0,
      estimator,
      keepLastMessages: 50
    });

    // Coherent tool-call group is too large; should not keep partial remnants.
    expect(trimmed.some((m) => m.role === "assistant" && Array.isArray(m.toolCalls) && m.toolCalls.length > 0)).toBe(false);
    expect(trimmed.some((m) => m.role === "tool")).toBe(false);

    const summary = trimmed.find((m) => m.role === "system" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER));
    expect(summary).toBeTruthy();
    expect(String(summary?.content)).toContain("tool_calls");
    expect(estimator.estimateMessagesTokens(trimmed)).toBeLessThanOrEqual(200);
  });
});
