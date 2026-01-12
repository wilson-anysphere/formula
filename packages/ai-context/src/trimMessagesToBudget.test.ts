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
});
