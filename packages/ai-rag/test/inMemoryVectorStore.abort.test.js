import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";

test("InMemoryVectorStore.list respects AbortSignal", async () => {
  const store = new InMemoryVectorStore({ dimension: 3 });
  await store.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);

  const abortController = new AbortController();
  abortController.abort();

  await assert.rejects(store.list({ workbookId: "wb", signal: abortController.signal }), { name: "AbortError" });
});

test("InMemoryVectorStore.query respects AbortSignal", async () => {
  const store = new InMemoryVectorStore({ dimension: 3 });
  await store.upsert([
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } },
    { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb" } },
  ]);

  const abortController = new AbortController();
  abortController.abort();

  await assert.rejects(store.query([1, 0, 0], 1, { workbookId: "wb", signal: abortController.signal }), {
    name: "AbortError",
  });
});

