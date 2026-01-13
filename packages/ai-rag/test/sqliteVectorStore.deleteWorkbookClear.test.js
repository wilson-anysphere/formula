import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryBinaryStorage } from "../src/store/binaryStorage.js";

function sortIds(records) {
  return records.map((r) => r.id).sort();
}

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

test("SqliteVectorStore.deleteWorkbook + clear (autoSave persists)", { skip: !sqlJsAvailable }, async () => {
  const storage = new InMemoryBinaryStorage();

  const modulePath = "../src/store/" + "sqliteVectorStore.js";
  const { SqliteVectorStore } = await import(modulePath);

  const store1 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  await store1.upsert([
    { id: "a", vector: [1, 0, 0], metadata: { workbookId: "wb1" } },
    { id: "b", vector: [0, 1, 0], metadata: { workbookId: "wb1" } },
    { id: "c", vector: [0, 0, 1], metadata: { workbookId: "wb2" } },
    { id: "d", vector: [1, 1, 0], metadata: { workbookId: "wb2" } },
  ]);

  const deleted = await store1.deleteWorkbook("wb1");
  assert.equal(deleted, 2);

  // Ensure deleteWorkbook persisted without requiring close().
  const store2 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: true });
  const remaining = await store2.list({ includeVector: false });
  assert.deepEqual(sortIds(remaining), ["c", "d"]);

  await store2.clear();

  // Ensure clear persisted without requiring close().
  const store3 = await SqliteVectorStore.create({ storage, dimension: 3, autoSave: false });
  const afterClear = await store3.list({ includeVector: false });
  assert.deepEqual(afterClear, []);

  await store1.close();
  await store2.close();
  await store3.close();
});
