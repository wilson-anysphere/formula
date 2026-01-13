import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryBinaryStorage } from "../src/store/binaryStorage.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { JsonVectorStore } from "../src/store/jsonVectorStore.js";

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

/**
 * @param {any} store
 */
async function seed(store) {
  await store.upsert([
    {
      id: "a",
      vector: [1, 0, 0],
      metadata: { workbookId: "wb", contentHash: "ch-a", metadataHash: "mh-a" },
    },
    {
      id: "b",
      vector: [0, 1, 0],
      metadata: { workbookId: "wb", contentHash: "ch-b" },
    },
    {
      id: "c",
      vector: [0, 0, 1],
      metadata: { workbookId: "wb2", metadataHash: "mh-c" },
    },
  ]);
}

test("VectorStore.listContentHashes returns workbook-scoped hashes", async () => {
  const store = new InMemoryVectorStore({ dimension: 3 });
  await seed(store);
  const rows = await store.listContentHashes({ workbookId: "wb" });
  assert.deepEqual(rows, [
    { id: "a", contentHash: "ch-a", metadataHash: "mh-a" },
    { id: "b", contentHash: "ch-b", metadataHash: null },
  ]);
});

test("JsonVectorStore.listContentHashes loads from storage first", async () => {
  const storage = new InMemoryBinaryStorage();
  const store1 = new JsonVectorStore({ dimension: 3, autoSave: true, storage });
  await seed(store1);
  await store1.close();

  const store2 = new JsonVectorStore({ dimension: 3, autoSave: false, storage });
  const rows = await store2.listContentHashes({ workbookId: "wb" });
  assert.deepEqual(rows, [
    { id: "a", contentHash: "ch-a", metadataHash: "mh-a" },
    { id: "b", contentHash: "ch-b", metadataHash: null },
  ]);
  await store2.close();
});

test(
  "SqliteVectorStore.listContentHashes avoids metadata JSON parsing and returns hashes",
  { skip: !sqlJsAvailable },
  async () => {
    // Same reasoning as above: avoid literal dynamic import specifiers so node:test can run this
    // file in dependency-free environments.
    const modulePath = "../src/store/" + "sqliteVectorStore.js";
    const { SqliteVectorStore } = await import(modulePath);
    const store = await SqliteVectorStore.create({
      dimension: 3,
      autoSave: false,
      storage: new InMemoryBinaryStorage(),
    });
    try {
      await seed(store);
      const rows = await store.listContentHashes({ workbookId: "wb" });
      assert.deepEqual(rows, [
        { id: "a", contentHash: "ch-a", metadataHash: "mh-a" },
        { id: "b", contentHash: "ch-b", metadataHash: null },
      ]);
    } finally {
      await store.close();
    }
  }
);

test("VectorStore.listContentHashes respects AbortSignal", async () => {
  const store = new InMemoryVectorStore({ dimension: 3 });
  await seed(store);
  const ac = new AbortController();
  ac.abort();
  await assert.rejects(store.listContentHashes({ signal: ac.signal }), { name: "AbortError" });
});

