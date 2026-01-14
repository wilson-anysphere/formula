import assert from "node:assert/strict";
import test from "node:test";

import { createHeuristicTokenEstimator } from "../src/tokenBudget.js";
import { CONTEXT_SUMMARY_MARKER, trimMessagesToBudget } from "../src/trimMessagesToBudget.js";

test("trimMessagesToBudget: preserves system messages and summarizes dropped history", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "system-1" },
    { role: "system", content: "system-2" },
    ...Array.from({ length: 10 }, (_v, i) => ({ role: "user", content: `user-${i}-` + "x".repeat(50) })),
    { role: "assistant", content: "latest assistant" },
  ];

  const trimmed = await trimMessagesToBudget({
    messages,
    maxTokens: 400,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 3,
    summaryMaxTokens: 100,
  });

  assert.deepStrictEqual(trimmed[0], { role: "system", content: "system-1" });
  assert.deepStrictEqual(trimmed[1], { role: "system", content: "system-2" });

  const summary = trimmed.find(
    (m) => m.role === "system" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER),
  );
  assert.ok(summary);

  // Should keep the most recent non-system messages.
  assert.deepStrictEqual(trimmed.at(-1), { role: "assistant", content: "latest assistant" });
  assert.equal(
    trimmed.filter((m) => m.role !== "system").length,
    3,
  );

  assert.ok(estimator.estimateMessagesTokens(trimmed) <= 400);
});

test("trimMessagesToBudget: returns messages unchanged when already under budget", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });
  const messages = [
    { role: "system", content: "sys" },
    { role: "user", content: "hello" },
    { role: "assistant", content: "hi" },
  ];

  const trimmed = await trimMessagesToBudget({ messages, maxTokens: 100, reserveForOutputTokens: 0, estimator });
  assert.deepStrictEqual(trimmed, messages);
});

test("trimMessagesToBudget: replaces existing generated summary messages instead of accumulating them", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "sys" },
    { role: "system", content: `${CONTEXT_SUMMARY_MARKER}\nOld summary` },
    ...Array.from({ length: 8 }, (_v, i) => ({ role: "user", content: `user-${i}-` + "y".repeat(60) })),
    { role: "assistant", content: "final" },
  ];

  const trimmed = await trimMessagesToBudget({
    messages,
    maxTokens: 300,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 2,
    summaryMaxTokens: 80,
  });

  const summaryMessages = trimmed.filter(
    (m) => m.role === "system" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER),
  );
  assert.equal(summaryMessages.length, 1);
  assert.ok(estimator.estimateMessagesTokens(trimmed) <= 300);
});

test("trimMessagesToBudget: replaces existing generated summary messages when summaryRole is assistant", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "sys" },
    { role: "assistant", content: `${CONTEXT_SUMMARY_MARKER}\nOld summary` },
    ...Array.from({ length: 8 }, (_v, i) => ({ role: "user", content: `user-${i}-` + "y".repeat(60) })),
    { role: "assistant", content: "final" },
  ];

  const trimmed = await trimMessagesToBudget({
    messages,
    maxTokens: 300,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 2,
    summaryMaxTokens: 80,
    summaryRole: "assistant",
  });

  const summaryMessages = trimmed.filter(
    (m) => m.role === "assistant" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER),
  );
  assert.equal(summaryMessages.length, 1);
  assert.ok(estimator.estimateMessagesTokens(trimmed) <= 300);
});

test("trimMessagesToBudget: respects AbortSignal (already aborted)", async () => {
  const abortController = new AbortController();
  abortController.abort();

  await assert.rejects(
    () =>
      trimMessagesToBudget({
        messages: [{ role: "user", content: "hello" }],
        maxTokens: 10,
        reserveForOutputTokens: 0,
        signal: abortController.signal,
      }),
    (err) => {
      assert.ok(err && typeof err === "object");
      assert.equal(err.name, "AbortError");
      return true;
    },
  );
});

test("trimMessagesToBudget: aborts while awaiting summarize callback", { timeout: 2000 }, async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });
  const abortController = new AbortController();

  /** @type {null | (() => void)} */
  let resolveStarted = null;
  const started = new Promise((resolve) => {
    resolveStarted = resolve;
  });

  const promise = trimMessagesToBudget({
    messages: [
      { role: "system", content: "sys" },
      ...Array.from({ length: 10 }, (_v, i) => ({ role: "user", content: `user-${i}-` + "x".repeat(50) })),
      { role: "assistant", content: "latest assistant" },
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
    signal: abortController.signal,
  });

  await started;
  abortController.abort();
  await assert.rejects(promise, (err) => {
    assert.ok(err && typeof err === "object");
    assert.equal(err.name, "AbortError");
    return true;
  });
});

test("trimMessagesToBudget: preserves tool-call coherence (no orphan tool messages) when trimming", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "sys" },
    { role: "user", content: "older-" + "x".repeat(200) },
    {
      role: "assistant",
      content: "",
      toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "foo" } }],
    },
    { role: "tool", toolCallId: "call-1", content: "tool-result" },
    { role: "assistant", content: "used tool output" },
    { role: "user", content: "latest question" },
  ];

  const trimmed = await trimMessagesToBudget({
    messages,
    maxTokens: 220,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 3,
    summaryMaxTokens: 50,
  });

  assert.ok(trimmed.some((m) => m.role === "tool"));

  // Any kept tool message must be preceded by an assistant toolCalls message.
  for (let i = 0; i < trimmed.length; i += 1) {
    const msg = trimmed[i];
    if (msg.role !== "tool") continue;

    let j = i - 1;
    while (j >= 0 && trimmed[j]?.role === "tool") j -= 1;
    assert.ok(j >= 0);
    assert.equal(trimmed[j].role, "assistant");
    assert.ok(Array.isArray(trimmed[j].toolCalls));
    assert.ok(trimmed[j].toolCalls.map((c) => c.id).includes(msg.toolCallId));
  }

  // Any kept assistant toolCalls message should have its tool messages kept when present.
  for (let i = 0; i < trimmed.length; i += 1) {
    const msg = trimmed[i];
    if (msg.role !== "assistant" || !Array.isArray(msg.toolCalls) || msg.toolCalls.length === 0) continue;
    const expectedIds = new Set(msg.toolCalls.map((c) => c.id));
    const seen = new Set();
    for (let j = i + 1; j < trimmed.length; j += 1) {
      if (trimmed[j].role !== "tool") break;
      seen.add(trimmed[j].toolCallId);
    }
    for (const id of expectedIds) assert.ok(seen.has(id));
  }

  assert.ok(estimator.estimateMessagesTokens(trimmed) <= 220);
});

test("trimMessagesToBudget: drops orphan tool messages when trimming with preserveToolCallPairs", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "sys" },
    { role: "user", content: "older-" + "x".repeat(200) },
    // Orphan tool message (no preceding assistant toolCalls message).
    { role: "tool", toolCallId: "call-1", content: "orphan-tool-result" },
    { role: "user", content: "latest question" },
  ];

  const trimmed = await trimMessagesToBudget({
    messages,
    maxTokens: 200,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 2,
    summaryMaxTokens: 180,
  });

  assert.equal(
    trimmed.some((m) => m.role === "tool"),
    false,
  );

  const summary = trimmed.find((m) => typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER));
  assert.ok(summary);
  assert.ok(String(summary?.content).includes("orphan-tool-result"));
});

test("trimMessagesToBudget: drops tool messages whose toolCallId doesn't match the preceding assistant toolCalls", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "sys" },
    { role: "user", content: "older-" + "x".repeat(400) },
    {
      role: "assistant",
      content: "",
      toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "foo" } }],
    },
    // Mismatched tool result (should be treated as orphan).
    { role: "tool", toolCallId: "call-2", content: "wrong-tool-result" },
    // Correct tool result for call-1.
    { role: "tool", toolCallId: "call-1", content: "right-tool-result" },
    { role: "assistant", content: "after tool" },
    { role: "user", content: "latest" },
  ];

  const trimmed = await trimMessagesToBudget({
    messages,
    maxTokens: 260,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 10,
    summaryMaxTokens: 120,
  });

  assert.equal(
    trimmed.some((m) => m.role === "tool" && m.toolCallId === "call-2"),
    false,
  );
  assert.equal(
    trimmed.some((m) => m.role === "tool" && m.toolCallId === "call-1"),
    true,
  );
});

test("trimMessagesToBudget: preserves coherence for assistant messages with multiple toolCalls", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "sys" },
    { role: "user", content: "older-" + "x".repeat(300) },
    {
      role: "assistant",
      content: "",
      toolCalls: [
        { id: "call-1", name: "t1", arguments: { q: "a" } },
        { id: "call-2", name: "t2", arguments: { q: "b" } },
      ],
    },
    { role: "tool", toolCallId: "call-1", content: "r1" },
    { role: "tool", toolCallId: "call-2", content: "r2" },
    { role: "assistant", content: "after tool" },
    { role: "user", content: "latest" },
  ];

  const trimmed = await trimMessagesToBudget({
    messages,
    maxTokens: 260,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 3,
    summaryMaxTokens: 0,
  });

  const assistantIdx = trimmed.findIndex((m) => m.role === "assistant" && Array.isArray(m.toolCalls) && m.toolCalls.length > 0);
  assert.ok(assistantIdx >= 0);

  const toolIds = new Set();
  for (let i = assistantIdx + 1; i < trimmed.length; i += 1) {
    if (trimmed[i].role !== "tool") break;
    toolIds.add(trimmed[i].toolCallId);
  }
  assert.deepStrictEqual(toolIds, new Set(["call-1", "call-2"]));
});

test("trimMessagesToBudget: can disable tool-call coherence (preserveToolCallPairs=false)", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "sys" },
    {
      role: "assistant",
      content: "",
      // Large tool call payload so it gets dropped before the tool message under tight budgets.
      toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "x".repeat(200) } }],
    },
    { role: "tool", toolCallId: "call-1", content: "tool-result" },
    { role: "assistant", content: "ok" },
    { role: "user", content: "latest" },
  ];

  const trimmed = await trimMessagesToBudget({
    messages,
    maxTokens: 120,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 10,
    summaryMaxTokens: 30,
    preserveToolCallPairs: false,
  });

  const toolIdx = trimmed.findIndex((m) => m.role === "tool");
  assert.ok(toolIdx >= 0);
  const hasAssistantToolCalls = trimmed.some((m) => m.role === "assistant" && Array.isArray(m.toolCalls) && m.toolCalls.length > 0);
  assert.equal(hasAssistantToolCalls, false);
});

test("trimMessagesToBudget: summarizes a dropped tool-call group instead of keeping a partial group under tight budgets", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "sys" },
    {
      role: "assistant",
      content: "",
      toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "x".repeat(200) } }],
    },
    { role: "tool", toolCallId: "call-1", content: "tool-result" },
    { role: "assistant", content: "after tool" },
    { role: "user", content: "latest" },
  ];

  const trimmed = await trimMessagesToBudget({
    messages,
    maxTokens: 220,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 10,
    summaryMaxTokens: 170,
  });

  // Tool call group should be summarized away, not partially kept.
  assert.equal(
    trimmed.some((m) => m.role === "tool"),
    false,
  );
  assert.equal(
    trimmed.some((m) => m.role === "assistant" && Array.isArray(m.toolCalls) && m.toolCalls.length > 0),
    false,
  );

  const summary = trimmed.find((m) => m.role === "system" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER));
  assert.ok(summary);
  assert.ok(String(summary?.content).includes("tool_calls"));
});

test("trimMessagesToBudget: dropToolMessagesFirst prefers dropping completed tool-call groups to keep more chat context", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "sys" },
    { role: "user", content: "u1" },
    { role: "assistant", content: "a1" },
    {
      role: "assistant",
      content: "",
      toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "x".repeat(250) } }],
    },
    { role: "tool", toolCallId: "call-1", content: "y".repeat(250) },
    { role: "assistant", content: "a2" },
    { role: "user", content: "u2" },
  ];

  const without = await trimMessagesToBudget({
    messages,
    maxTokens: 200,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 50,
    summaryMaxTokens: 50,
  });

  const withDropFirst = await trimMessagesToBudget({
    messages,
    maxTokens: 200,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 50,
    summaryMaxTokens: 50,
    dropToolMessagesFirst: true,
  });

  // Without dropToolMessagesFirst, we should lose the early chat context due to the expensive tool messages.
  assert.equal(without.some((m) => m.role === "user" && m.content === "u1"), false);
  // With dropToolMessagesFirst, tool call group should be summarized/dropped first, preserving early chat context.
  assert.equal(withDropFirst.some((m) => m.role === "user" && m.content === "u1"), true);
  assert.equal(withDropFirst.some((m) => m.role === "tool"), false);
});

test("trimMessagesToBudget: falls back to a summary when the most recent tool-call group cannot fit at all", async () => {
  const estimator = createHeuristicTokenEstimator({ charsPerToken: 1, tokensPerMessageOverhead: 0 });

  const messages = [
    { role: "system", content: "sys" },
    {
      role: "assistant",
      content: "",
      toolCalls: [{ id: "call-1", name: "lookup", arguments: { q: "x".repeat(1000) } }],
    },
    { role: "tool", toolCallId: "call-1", content: "tool-result" },
  ];

  const trimmed = await trimMessagesToBudget({
    messages,
    maxTokens: 200,
    reserveForOutputTokens: 0,
    estimator,
    keepLastMessages: 50,
  });

  // Coherent tool-call group is too large; should not keep partial remnants.
  assert.equal(
    trimmed.some((m) => m.role === "assistant" && Array.isArray(m.toolCalls) && m.toolCalls.length > 0),
    false,
  );
  assert.equal(
    trimmed.some((m) => m.role === "tool"),
    false,
  );

  const summary = trimmed.find((m) => m.role === "system" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER));
  assert.ok(summary);
  assert.ok(String(summary?.content).includes("tool_calls"));
  assert.ok(estimator.estimateMessagesTokens(trimmed) <= 200);
});

