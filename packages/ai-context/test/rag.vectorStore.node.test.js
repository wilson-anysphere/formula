import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryVectorStore } from "../src/rag.js";

test("InMemoryVectorStore.search: returns deterministic id ordering when similarity scores tie (full sort path)", async () => {
  const store = new InMemoryVectorStore();
  const embedding = [1, 0, 0];

  // Insert out of order to ensure the sort tie-breaker (id) drives final ordering
  // when similarity scores are identical.
  await store.add([
    { id: "chunk-b", embedding, metadata: null, text: "b" },
    { id: "chunk-a", embedding, metadata: null, text: "a" },
    { id: "chunk-c", embedding, metadata: null, text: "c" },
  ]);

  // Use a `topK` larger than the store size so `search()` takes its full-sort path.
  const results = await store.search([1, 0, 0], 10);
  assert.deepStrictEqual(
    results.map((r) => r.item.id),
    ["chunk-a", "chunk-b", "chunk-c"],
  );
});

test("InMemoryVectorStore.search: accounts for extra embedding dimensions when computing cosine similarity", async () => {
  const store = new InMemoryVectorStore();

  // `chunk-long` shares the same prefix direction as the query but has a large extra
  // component; it should not tie with the shorter, exact-match embedding.
  await store.add([
    { id: "chunk-long", embedding: [1, 0, 100], metadata: null, text: "long" },
    { id: "chunk-short", embedding: [1, 0], metadata: null, text: "short" },
  ]);

  const results = await store.search([1, 0], 2);
  assert.deepStrictEqual(
    results.map((r) => r.item.id),
    ["chunk-short", "chunk-long"],
  );
});

test("InMemoryVectorStore.search: respects AbortSignal", async () => {
  const store = new InMemoryVectorStore();

  const abortController = new AbortController();
  abortController.abort();

  await assert.rejects(store.search([1, 0], 1, { signal: abortController.signal }), { name: "AbortError" });
});

test("InMemoryVectorStore.search: checks AbortSignal while scanning items", async () => {
  const store = new InMemoryVectorStore();

  const abortController = new AbortController();
  let didAbort = false;

  const embedding = new Proxy([1, 0, 0], {
    get(target, prop, receiver) {
      // Trigger cancellation on the first read of the vector to ensure the search loop
      // observes aborts between item iterations.
      if (!didAbort && prop === "0") {
        didAbort = true;
        abortController.abort();
      }
      return Reflect.get(target, prop, receiver);
    },
  });

  await store.add([
    { id: "chunk-a", embedding, metadata: null, text: "a" },
    { id: "chunk-b", embedding: [1, 0, 0], metadata: null, text: "b" },
  ]);

  await assert.rejects(store.search([1, 0, 0], 1, { signal: abortController.signal }), { name: "AbortError" });
  assert.equal(didAbort, true);
});

test("InMemoryVectorStore.search: treats non-finite topK (NaN) as 'all results'", async () => {
  const store = new InMemoryVectorStore();
  const embedding = [1, 0, 0];

  await store.add([
    { id: "chunk-b", embedding, metadata: null, text: "b" },
    { id: "chunk-a", embedding, metadata: null, text: "a" },
  ]);

  const results = await store.search([1, 0, 0], Number.NaN);
  assert.deepStrictEqual(
    results.map((r) => r.item.id),
    ["chunk-a", "chunk-b"],
  );
});

test("InMemoryVectorStore.search: returns an empty result set when topK <= 0", async () => {
  const store = new InMemoryVectorStore();
  const embedding = [1, 0, 0];

  await store.add([{ id: "chunk-a", embedding, metadata: null, text: "a" }]);

  assert.deepStrictEqual(await store.search([1, 0, 0], 0), []);
  assert.deepStrictEqual(await store.search([1, 0, 0], -1), []);
});
