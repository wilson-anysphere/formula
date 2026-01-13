import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";

let sqlJsAvailable = true;
try {
  // Keep this as a computed dynamic import (no literal bare specifier) so
  // `scripts/run-node-tests.mjs` can still execute this file when `node_modules/`
  // is missing.
  const sqlJsModuleName = "sql" + ".js";
  await import(sqlJsModuleName);
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

  // Ensure tie-breaking influences selection at the cutoff (topK=1), not just output ordering.
  const top1 = await store.query([1, 0, 0], 1, { workbookId: "wb" });
  assert.deepEqual(
    top1.map((h) => h.id),
    ["a"],
  );

  const hits = await store.query([1, 0, 0], 2, { workbookId: "wb" });
  assert.deepEqual(
    hits.map((h) => h.id),
    ["a", "b"],
  );
});

test("SqliteVectorStore.query breaks score ties by id (ascending)", { skip: !sqlJsAvailable }, async () => {
  // Same reasoning as above: avoid literal dynamic import specifiers so
  // node:test can run this file in dependency-free environments.
  const modulePath = "../src/store/" + "sqliteVectorStore.js";
  const { SqliteVectorStore } = await import(modulePath);
  const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
  try {
    await store.upsert([
      { id: "b", vector: [1, 0, 0], metadata: { workbookId: "wb" } },
      { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } },
      { id: "c", vector: [0, 1, 0], metadata: { workbookId: "wb" } },
    ]);

    const top1 = await store.query([1, 0, 0], 1, { workbookId: "wb" });
    assert.deepEqual(
      top1.map((h) => h.id),
      ["a"],
    );

    const hits = await store.query([1, 0, 0], 2, { workbookId: "wb" });
    assert.deepEqual(
      hits.map((h) => h.id),
      ["a", "b"],
    );
  } finally {
    await store.close();
  }
});
