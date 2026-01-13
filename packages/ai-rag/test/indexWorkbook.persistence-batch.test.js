import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";
import { indexWorkbook } from "../src/pipeline/indexWorkbook.js";
import { InMemoryBinaryStorage } from "../src/store/binaryStorage.js";
import { JsonVectorStore } from "../src/store/jsonVectorStore.js";
import { SqliteVectorStore } from "../src/store/sqliteVectorStore.js";

class CountingBinaryStorage extends InMemoryBinaryStorage {
  constructor() {
    super();
    this.saveCalls = 0;
  }

  async save(data) {
    this.saveCalls += 1;
    await super.save(data);
  }
}

function makeWorkbookTwoTables() {
  return {
    id: "wb-indexWorkbook-persist-batch",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          [{ v: "A" }, { v: "B" }, null, { v: "C" }, { v: "D" }],
          [{ v: 1 }, { v: 2 }, null, { v: 3 }, { v: 4 }],
        ],
      },
    ],
    tables: [
      { name: "T1", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } },
      { name: "T2", sheetName: "Sheet1", rect: { r0: 0, c0: 3, r1: 1, c1: 4 } },
    ],
  };
}

function makeWorkbookUpsertAndDelete() {
  const workbook = makeWorkbookTwoTables();

  // Change T1 so we trigger an upsert.
  workbook.sheets[0].cells[1][0] = { v: 999 };

  // Remove T2 and its cells so we trigger a delete without generating a replacement chunk.
  workbook.tables = workbook.tables.filter((t) => t.name !== "T2");
  workbook.sheets[0].cells[0][3] = null;
  workbook.sheets[0].cells[0][4] = null;
  workbook.sheets[0].cells[1][3] = null;
  workbook.sheets[0].cells[1][4] = null;

  return workbook;
}

test("indexWorkbook persists once when batching JsonVectorStore mutations", async () => {
  const embedder = new HashEmbedder({ dimension: 128 });
  const storage = new CountingBinaryStorage();
  const store = new JsonVectorStore({ storage, dimension: 128, autoSave: true });

  await indexWorkbook({ workbook: makeWorkbookTwoTables(), vectorStore: store, embedder });
  storage.saveCalls = 0;

  const res = await indexWorkbook({ workbook: makeWorkbookUpsertAndDelete(), vectorStore: store, embedder });
  assert.equal(res.upserted, 1);
  assert.equal(res.deleted, 1);
  assert.equal(storage.saveCalls, 1);
});

let sqlJsAvailable = true;
try {
  await import("sql.js");
} catch {
  sqlJsAvailable = false;
}

test(
  "indexWorkbook persists once when batching SqliteVectorStore mutations",
  { skip: !sqlJsAvailable },
  async () => {
    const embedder = new HashEmbedder({ dimension: 128 });
    const storage = new CountingBinaryStorage();
    const store = await SqliteVectorStore.create({ storage, dimension: 128, autoSave: true });

    await indexWorkbook({ workbook: makeWorkbookTwoTables(), vectorStore: store, embedder });
    storage.saveCalls = 0;

    const res = await indexWorkbook({ workbook: makeWorkbookUpsertAndDelete(), vectorStore: store, embedder });
    assert.equal(res.upserted, 1);
    assert.equal(res.deleted, 1);
    assert.equal(storage.saveCalls, 1);

    await store.close();
  }
);

