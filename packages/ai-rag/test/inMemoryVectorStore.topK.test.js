import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";

test("InMemoryVectorStore.query validates topK", async () => {
  const store = new InMemoryVectorStore({ dimension: 3 });
  await store.upsert([
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } },
    { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb" } },
  ]);

  assert.deepEqual(await store.query([1, 0, 0], 0, { workbookId: "wb" }), []);
  assert.deepEqual(await store.query([1, 0, 0], -1, { workbookId: "wb" }), []);
  await assert.rejects(store.query([1, 0, 0], Number.NaN, { workbookId: "wb" }), /Invalid topK/);

  const floored = await store.query([1, 0, 0], 1.9, { workbookId: "wb" });
  assert.equal(floored.length, 1);
  assert.equal(floored[0]?.id, "a");
});

