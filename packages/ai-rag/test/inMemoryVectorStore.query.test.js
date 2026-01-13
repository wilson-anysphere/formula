import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { cosineSimilarity, normalizeL2 } from "../src/store/vectorMath.js";

/**
 * Baseline implementation that mirrors `InMemoryVectorStore.query` behavior:
 * score everything, full sort descending (tie-break by id asc), then slice.
 *
 * @param {InMemoryVectorStore} store
 * @param {ArrayLike<number>} vector
 * @param {number} topK
 * @param {{ filter?: (metadata: any, id: string) => boolean, workbookId?: string }} [opts]
 */
async function baselineQuery(store, vector, topK, opts) {
  const q = normalizeL2(vector);
  const records = await store.list({
    filter: opts?.filter,
    workbookId: opts?.workbookId,
    includeVector: true,
  });

  const scored = [];
  for (const rec of records) {
    const score = cosineSimilarity(q, rec.vector);
    scored.push({ id: rec.id, score, metadata: rec.metadata });
  }
  scored.sort((a, b) => {
    const scoreCmp = b.score - a.score;
    if (scoreCmp !== 0) return scoreCmp;
    if (a.id < b.id) return -1;
    if (a.id > b.id) return 1;
    return 0;
  });
  return scored.slice(0, topK);
}

test("InMemoryVectorStore.query matches baseline full-sort behavior (deterministic ties, filters)", async () => {
  const store = new InMemoryVectorStore({ dimension: 3 });
  await store.upsert([
    { id: "b", vector: [1, 0, 0], metadata: { workbookId: "wb1", tag: "keep" } },
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb1", tag: "keep" } },
    { id: "d", vector: [0, 1, 0], metadata: { workbookId: "wb1", tag: "keep" } },
    { id: "c", vector: [0, 1, 0], metadata: { workbookId: "wb1", tag: "keep" } },
    { id: "e", vector: [1, 1, 0], metadata: { workbookId: "wb1", tag: "keep" } },
    { id: "f", vector: [1, 0, 0], metadata: { workbookId: "wb2", tag: "other" } },
    { id: "g", vector: [1, 0, 0], metadata: { workbookId: "wb1", tag: "filtered" } },
  ]);

  {
    const opts = { workbookId: "wb1" };
    const expected = await baselineQuery(store, [1, 0, 0], 1, opts);
    const actual = await store.query([1, 0, 0], 1, opts);
    assert.deepStrictEqual(actual, expected);
    // Explicitly exercise deterministic tie handling at the cutoff: "a" should win over "b"/"g".
    assert.equal(actual[0]?.id, "a");
  }

  {
    const opts = { workbookId: "wb1", filter: (md) => md.tag === "keep" };
    const expected = await baselineQuery(store, [1, 0, 0], 10, opts);
    const actual = await store.query([1, 0, 0], 10, opts);
    assert.deepStrictEqual(actual, expected);
    assert.deepStrictEqual(
      actual.map((r) => r.id),
      ["a", "b", "e", "c", "d"],
    );
  }
});

test("InMemoryVectorStore.query keeps only topK candidates while scanning (no full scored array)", async () => {
  const store = new InMemoryVectorStore({ dimension: 3 });
  const records = [];
  for (let i = 0; i < 1000; i += 1) {
    // Deterministic, non-zero vectors. Many ties by design.
    const x = i % 3 === 0 ? 1 : 0;
    const y = i % 3 === 1 ? 1 : 0;
    const z = i % 3 === 2 ? 1 : 0;
    records.push({ id: `id${i}`, vector: [x, y, z], metadata: { workbookId: "wb" } });
  }
  await store.upsert(records);

  const topK = 5;

  const originalSort = Array.prototype.sort;
  const originalPush = Array.prototype.push;
  let maxSortLen = 0;
  let pushCalls = 0;

  Array.prototype.sort = function (...args) {
    // If query reverted to the old behavior, we'd see a `.sort()` over 1000 items here.
    maxSortLen = Math.max(maxSortLen, this.length);
    return originalSort.apply(this, args);
  };
  Array.prototype.push = function (...args) {
    pushCalls += 1;
    return originalPush.apply(this, args);
  };

  try {
    const results = await store.query([1, 0, 0], topK, { workbookId: "wb" });
    assert.equal(results.length, topK);
  } finally {
    Array.prototype.sort = originalSort;
    Array.prototype.push = originalPush;
  }

  assert.ok(maxSortLen <= topK, `expected query to sort <= ${topK} items, sorted ${maxSortLen}`);
  assert.ok(pushCalls <= topK, `expected query to push <= ${topK} items, pushed ${pushCalls}`);
});
