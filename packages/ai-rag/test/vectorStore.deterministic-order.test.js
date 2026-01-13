import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { SqliteVectorStore } from "../src/store/sqliteVectorStore.js";

let sqlJsAvailable = true;
try {
  await import("sql.js");
} catch {
  sqlJsAvailable = false;
}

test("InMemoryVectorStore.query breaks score ties by id (ascending)", async () => {
  const store = new InMemoryVectorStore({ dimension: 3 });

  // Insert out-of-order ids so the test fails if the implementation preserves insertion
  // order for equal scores instead of applying a deterministic tie-break.
  await store.upsert([
    { id: "b", vector: [1, 0, 0], metadata: { workbookId: "wb" } },
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } },
    { id: "c", vector: [0, 1, 0], metadata: { workbookId: "wb" } },
  ]);

  const hits = await store.query([1, 0, 0], 2, { workbookId: "wb" });
  assert.deepEqual(
    hits.map((h) => h.id),
    ["a", "b"],
  );
});

test("SqliteVectorStore.query breaks score ties by id (ascending)", { skip: !sqlJsAvailable }, async () => {
  const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
  try {
    await store.upsert([
      { id: "b", vector: [1, 0, 0], metadata: { workbookId: "wb" } },
      { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } },
      { id: "c", vector: [0, 1, 0], metadata: { workbookId: "wb" } },
    ]);

    const hits = await store.query([1, 0, 0], 2, { workbookId: "wb" });
    assert.deepEqual(
      hits.map((h) => h.id),
      ["a", "b"],
    );
  } finally {
    await store.close();
  }
});

