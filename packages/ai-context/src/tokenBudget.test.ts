import { describe, expect, it } from "vitest";

import { packSectionsToTokenBudget } from "./tokenBudget.js";

describe("tokenBudget", () => {
  it("packSectionsToTokenBudget respects a custom TokenEstimator", () => {
    const charEstimator = {
      estimateTextTokens: (text: string) => String(text ?? "").length,
      estimateMessageTokens: (message: any) => JSON.stringify(message ?? "").length,
      estimateMessagesTokens: (messages: any[]) => JSON.stringify(messages ?? []).length
    };

    const sections = [
      { key: "first", priority: 2, text: "x".repeat(100) },
      { key: "second", priority: 1, text: "y".repeat(10) }
    ];

    const maxTokens = 40;

    // Default estimator (~4 chars/token) will not trim the 100-char section.
    const packedDefault = packSectionsToTokenBudget(sections, maxTokens);
    expect(packedDefault).toHaveLength(2);
    expect(packedDefault[0]).toEqual({ key: "first", text: "x".repeat(100) });

    // Under the strict 1 char/token estimator, the first section must be trimmed and
    // should consume the entire budget, dropping lower priority sections.
    const packedStrict = packSectionsToTokenBudget(sections, maxTokens, charEstimator as any);
    expect(packedStrict).toHaveLength(1);
    expect(packedStrict[0]).toEqual({
      key: "first",
      // suffix is deterministic; with 1 char/token, the prefix fits exactly.
      text: "x".repeat(8) + "\n…(trimmed to fit token budget)…"
    });

    expect(charEstimator.estimateTextTokens(packedStrict[0].text)).toBeLessThanOrEqual(maxTokens);
    // The default-packed context would exceed the strict budget.
    expect(charEstimator.estimateTextTokens(packedDefault[0].text)).toBeGreaterThan(maxTokens);
  });

  it("packSectionsToTokenBudget respects AbortSignal", () => {
    const abortController = new AbortController();
    abortController.abort();

    let error: unknown = null;
    try {
      packSectionsToTokenBudget([{ key: "a", priority: 1, text: "hello" }], 10, undefined, {
        signal: abortController.signal,
      });
    } catch (err) {
      error = err;
    }

    expect(error).toMatchObject({ name: "AbortError" });
  });
});
