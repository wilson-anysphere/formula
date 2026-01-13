import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryBinaryStorage } from "../src/store/binaryStorage.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { JsonVectorStore } from "../src/store/jsonVectorStore.js";

function sortIds(records) {
  return records.map((r) => r.id).sort();
}

test("InMemoryVectorStore.deleteWorkbook + clear", async () => {
  const store = new InMemoryVectorStore({ dimension: 3 });
  await store.upsert([
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb1" } },
    { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb1" } },
    { id: "c", vector: [0, 0, 1], metadata: { workbookId: "wb2" } },
    { id: "d", vector: [1, 1, 0], metadata: { workbookId: "wb2" } },
  ]);

  const deleted = await store.deleteWorkbook("wb1");
  assert.equal(deleted, 2);

  const remaining = await store.list({ includeVector: false });
  assert.deepEqual(sortIds(remaining), ["c", "d"]);

  await store.clear();
  const afterClear = await store.list({ includeVector: false });
  assert.deepEqual(afterClear, []);
});

test("JsonVectorStore.deleteWorkbook + clear (autoSave persists)", async () => {
  const storage = new InMemoryBinaryStorage();

  const store1 = new JsonVectorStore({ storage, dimension: 3, autoSave: true });
  await store1.upsert([
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb1" } },
    { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb1" } },
    { id: "c", vector: [0, 0, 1], metadata: { workbookId: "wb2" } },
    { id: "d", vector: [1, 1, 0], metadata: { workbookId: "wb2" } },
  ]);

  const deleted = await store1.deleteWorkbook("wb1");
  assert.equal(deleted, 2);

  // Ensure deleteWorkbook persisted without requiring close().
  const store2 = new JsonVectorStore({ storage, dimension: 3, autoSave: true });
  const remaining = await store2.list({ includeVector: false });
  assert.deepEqual(sortIds(remaining), ["c", "d"]);

  await store2.clear();

  // Ensure clear persisted without requiring close().
  const store3 = new JsonVectorStore({ storage, dimension: 3, autoSave: false });
  const afterClear = await store3.list({ includeVector: false });
  assert.deepEqual(afterClear, []);

  await store1.close();
  await store2.close();
  await store3.close();
});

test("JsonVectorStore.clear persists even when existing payload can't be loaded (dimension mismatch)", async () => {
  const storage = new InMemoryBinaryStorage();

  const store1 = new JsonVectorStore({ storage, dimension: 3, autoSave: true });
  await store1.upsert([{ id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb" } }]);

  // Create a store with a different dimension but opt out of automatic reset so the
  // old payload remains present. `clear()` should still overwrite the persisted bytes.
  const store2 = new JsonVectorStore({ storage, dimension: 4, autoSave: true, resetOnCorrupt: false });
  await store2.clear();

  const bytes = await storage.load();
  assert.ok(bytes);
  const parsed = JSON.parse(new TextDecoder().decode(bytes));
  assert.equal(parsed.dimension, 4);
  assert.deepEqual(parsed.records, []);

  await store1.close();
  await store2.close();
});
