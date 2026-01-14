import { describe, expect, it } from "vitest";

import { packSectionsToTokenBudget, packSectionsToTokenBudgetWithReport, stableJsonStringify } from "./tokenBudget.js";

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

  it("packSectionsToTokenBudget is deterministic when priorities tie", () => {
    const sections = [
      { key: "b", priority: 1, text: "B" },
      { key: "z", priority: 2, text: "Z" },
      { key: "a", priority: 1, text: "A" },
      { key: "y", priority: 2, text: "Y" }
    ];

    const packed = packSectionsToTokenBudget(sections, 1_000);
    expect(packed.map((s) => s.key)).toEqual(["z", "y", "b", "a"]);
  });

  it("stableJsonStringify is stable for Map insertion order", () => {
    const map1 = new Map<string, number>();
    map1.set("b", 1);
    map1.set("a", 2);

    const map2 = new Map<string, number>();
    map2.set("a", 2);
    map2.set("b", 1);

    expect(stableJsonStringify(map1)).toBe(stableJsonStringify(map2));
    expect(stableJsonStringify(map1)).toBe('[["a",2],["b",1]]');
  });

  it("stableJsonStringify is stable for Set insertion order", () => {
    const set1 = new Set<string>();
    set1.add("b");
    set1.add("a");

    const set2 = new Set<string>();
    set2.add("a");
    set2.add("b");

    expect(stableJsonStringify(set1)).toBe(stableJsonStringify(set2));
    expect(stableJsonStringify(set1)).toBe('["a","b"]');
  });

  it("stableJsonStringify handles cyclic objects without crashing", () => {
    const obj: any = { a: 1 };
    obj.self = obj;
    expect(stableJsonStringify(obj)).toBe('{"a":1,"self":"[Circular]"}');
  });

  it("stableJsonStringify handles cyclic Map/Set structures without crashing", () => {
    const map: any = new Map();
    map.set("self", map);
    expect(stableJsonStringify(map)).toBe('[["self","[Circular]"]]');

    const set: any = new Set();
    set.add(set);
    expect(stableJsonStringify(set)).toBe('["[Circular]"]');
  });

  it("packSectionsToTokenBudgetWithReport reports token usage, trims, and drops", () => {
    const charEstimator = {
      estimateTextTokens: (text: string) => String(text ?? "").length,
      estimateMessageTokens: (message: any) => JSON.stringify(message ?? "").length,
      estimateMessagesTokens: (messages: any[]) => JSON.stringify(messages ?? []).length
    };

    const suffix = "\n…(trimmed to fit token budget)…";
    const maxTokens = 40;

    const { packed, report } = packSectionsToTokenBudgetWithReport(
      [
        { key: "b", priority: 1, text: "y".repeat(10) },
        { key: "first", priority: 2, text: "x".repeat(100) },
        { key: "a", priority: 1, text: "z".repeat(10) }
      ],
      maxTokens,
      charEstimator as any
    );

    expect(packed).toEqual([{ key: "first", text: "x".repeat(8) + suffix }]);
    expect(report.maxTokens).toBe(maxTokens);
    expect(report.remainingTokens).toBe(0);
    expect(report.sections).toEqual([
      {
        key: "first",
        priority: 2,
        tokensPreTrim: 100,
        tokensPostTrim: 40,
        trimmed: true,
        dropped: false
      },
      {
        key: "b",
        priority: 1,
        tokensPreTrim: 10,
        tokensPostTrim: 0,
        trimmed: false,
        dropped: true
      },
      {
        key: "a",
        priority: 1,
        tokensPreTrim: 10,
        tokensPostTrim: 0,
        trimmed: false,
        dropped: true
      }
    ]);
  });
});
