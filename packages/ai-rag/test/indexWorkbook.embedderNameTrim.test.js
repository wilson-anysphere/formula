import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { indexWorkbook } from "../src/pipeline/indexWorkbook.js";

test("indexWorkbook trims embedder.name before storing it in metadata", async () => {
  const workbook = {
    id: "wb-embedder-name-trim",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "hello" }]],
      },
    ],
    tables: [{ name: "T1", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } }],
  };

  const base = new HashEmbedder({ dimension: 32 });
  const embedder = {
    name: `  ${base.name}  `,
    embedTexts: (texts, options) => base.embedTexts(texts, options),
  };

  const store = new InMemoryVectorStore({ dimension: 32 });
  await indexWorkbook({ workbook, vectorStore: store, embedder });

  const records = await store.list({ workbookId: workbook.id, includeVector: false });
  assert.ok(records.length > 0);
  for (const r of records) {
    assert.equal(r.metadata.embedder, base.name);
  }
});

