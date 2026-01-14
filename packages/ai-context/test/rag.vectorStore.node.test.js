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

