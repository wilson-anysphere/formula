import assert from "node:assert/strict";
import test from "node:test";

import { packSectionsToTokenBudget, packSectionsToTokenBudgetWithReport, stableJsonStringify } from "../src/tokenBudget.js";

test("tokenBudget: packSectionsToTokenBudget respects a custom TokenEstimator", () => {
  const charEstimator = {
    estimateTextTokens: (text) => String(text ?? "").length,
    estimateMessageTokens: (message) => JSON.stringify(message ?? "").length,
    estimateMessagesTokens: (messages) => JSON.stringify(messages ?? []).length,
  };

  const sections = [
    { key: "first", priority: 2, text: "x".repeat(100) },
    { key: "second", priority: 1, text: "y".repeat(10) },
  ];

  const maxTokens = 40;

  // Default estimator (~4 chars/token) will not trim the 100-char section.
  const packedDefault = packSectionsToTokenBudget(sections, maxTokens);
  assert.equal(packedDefault.length, 2);
  assert.deepStrictEqual(packedDefault[0], { key: "first", text: "x".repeat(100) });

  // Under the strict 1 char/token estimator, the first section must be trimmed and
  // should consume the entire budget, dropping lower priority sections.
  const packedStrict = packSectionsToTokenBudget(sections, maxTokens, charEstimator);
  assert.equal(packedStrict.length, 1);
  assert.deepStrictEqual(packedStrict[0], {
    key: "first",
    text: "x".repeat(8) + "\n…(trimmed to fit token budget)…",
  });

  assert.ok(charEstimator.estimateTextTokens(packedStrict[0].text) <= maxTokens);
  // The default-packed context would exceed the strict budget.
  assert.ok(charEstimator.estimateTextTokens(packedDefault[0].text) > maxTokens);
});

test("tokenBudget: packSectionsToTokenBudget respects AbortSignal", () => {
  const abortController = new AbortController();
  abortController.abort();

  let error = null;
  try {
    packSectionsToTokenBudget([{ key: "a", priority: 1, text: "hello" }], 10, undefined, {
      signal: abortController.signal,
    });
  } catch (err) {
    error = err;
  }

  assert.ok(error && typeof error === "object");
  assert.equal(error.name, "AbortError");
});

test("tokenBudget: packSectionsToTokenBudget is deterministic when priorities tie", () => {
  const sections = [
    { key: "b", priority: 1, text: "B" },
    { key: "z", priority: 2, text: "Z" },
    { key: "a", priority: 1, text: "A" },
    { key: "y", priority: 2, text: "Y" },
  ];

  const packed = packSectionsToTokenBudget(sections, 1_000);
  assert.deepStrictEqual(
    packed.map((s) => s.key),
    ["z", "y", "b", "a"],
  );
});

test("tokenBudget: stableJsonStringify is stable for Map insertion order", () => {
  const map1 = new Map();
  map1.set("b", 1);
  map1.set("a", 2);

  const map2 = new Map();
  map2.set("a", 2);
  map2.set("b", 1);

  assert.equal(stableJsonStringify(map1), stableJsonStringify(map2));
  assert.equal(stableJsonStringify(map1), '[["a",2],["b",1]]');
});

test("tokenBudget: stableJsonStringify is stable for Set insertion order", () => {
  const set1 = new Set();
  set1.add("b");
  set1.add("a");

  const set2 = new Set();
  set2.add("a");
  set2.add("b");

  assert.equal(stableJsonStringify(set1), stableJsonStringify(set2));
  assert.equal(stableJsonStringify(set1), '["a","b"]');
});

test("tokenBudget: stableJsonStringify handles cyclic objects without crashing", () => {
  const obj = { a: 1 };
  obj.self = obj;
  assert.equal(stableJsonStringify(obj), '{"a":1,"self":"[Circular]"}');
});

test("tokenBudget: stableJsonStringify handles cyclic Map/Set structures without crashing", () => {
  const map = new Map();
  map.set("self", map);
  assert.equal(stableJsonStringify(map), '[["self","[Circular]"]]');

  const set = new Set();
  set.add(set);
  assert.equal(stableJsonStringify(set), '["[Circular]"]');
});

test("tokenBudget: stableJsonStringify does not call custom toString() on objects when serialization fails", () => {
  let called = false;
  const obj = {
    get boom() {
      throw new Error("boom");
    },
    toString() {
      called = true;
      return "TopSecret";
    },
  };

  assert.equal(stableJsonStringify(obj), '"[Unserializable]"');
  assert.equal(called, false);
});

test("tokenBudget: packSectionsToTokenBudgetWithReport reports token usage, trims, and drops", () => {
  const charEstimator = {
    estimateTextTokens: (text) => String(text ?? "").length,
    estimateMessageTokens: (message) => JSON.stringify(message ?? "").length,
    estimateMessagesTokens: (messages) => JSON.stringify(messages ?? []).length,
  };

  const suffix = "\n…(trimmed to fit token budget)…";
  const maxTokens = 40;

  const { packed, report } = packSectionsToTokenBudgetWithReport(
    [
      { key: "b", priority: 1, text: "y".repeat(10) },
      { key: "first", priority: 2, text: "x".repeat(100) },
      { key: "a", priority: 1, text: "z".repeat(10) },
    ],
    maxTokens,
    charEstimator,
  );

  assert.deepStrictEqual(packed, [{ key: "first", text: "x".repeat(8) + suffix }]);
  assert.equal(report.maxTokens, maxTokens);
  assert.equal(report.remainingTokens, 0);
  assert.deepStrictEqual(report.sections, [
    {
      key: "first",
      priority: 2,
      tokensPreTrim: 100,
      tokensPostTrim: 40,
      trimmed: true,
      dropped: false,
    },
    {
      key: "b",
      priority: 1,
      tokensPreTrim: 10,
      tokensPostTrim: 0,
      trimmed: false,
      dropped: true,
    },
    {
      key: "a",
      priority: 1,
      tokensPreTrim: 10,
      tokensPostTrim: 0,
      trimmed: false,
      dropped: true,
    },
  ]);
});
