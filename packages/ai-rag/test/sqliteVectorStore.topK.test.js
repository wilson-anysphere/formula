import assert from "node:assert/strict";
import test from "node:test";

import { SqliteVectorStore } from "../src/store/sqliteVectorStore.js";

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

test("SqliteVectorStore.query validates topK", { skip: !sqlJsAvailable }, async () => {
  const store = await SqliteVectorStore.create({ dimension: 3, autoSave: false });
  try {
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
  } finally {
    await store.close();
  }
});
