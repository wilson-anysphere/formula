import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryVectorStore } from "../src/rag.js";

test("InMemoryVectorStore.search: selects deterministic ids when scores tie under topK", async () => {
  const store = new InMemoryVectorStore();
  const embedding = [1, 0, 0];

  // Insert out of order to ensure deterministic tie-breaking (id) is what drives
  // final ordering when similarity scores are identical.
  await store.add([
    { id: "chunk-b", embedding, metadata: null, text: "b" },
    { id: "chunk-a", embedding, metadata: null, text: "a" },
    { id: "chunk-c", embedding, metadata: null, text: "c" },
  ]);

  const results = await store.search([1, 0, 0], 2);
  assert.deepEqual(
    results.map((r) => r.item.id),
    ["chunk-a", "chunk-b"],
  );
});

