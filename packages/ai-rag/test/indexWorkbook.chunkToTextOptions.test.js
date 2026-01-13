import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";
import { indexWorkbook } from "../src/pipeline/indexWorkbook.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";

test("indexWorkbook forwards maxColumnsForSchema/maxColumnsForRows to chunkToText", async () => {
  const colCount = 10;
  const headers = Array.from({ length: colCount }, (_, i) => ({ v: `Header${i + 1}` }));
  const row = Array.from({ length: colCount }, (_, i) => ({ v: `Value${i + 1}` }));

  const workbook = {
    id: "wb-chunk-to-text-opts",
    sheets: [{ name: "Sheet1", cells: [headers, row] }],
    tables: [{ name: "WideTable", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: colCount - 1 } }],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const store = new InMemoryVectorStore({ dimension: 64 });

  await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    sampleRows: 1,
    maxColumnsForSchema: 3,
    maxColumnsForRows: 3,
  });

  const id = `${workbook.id}::Sheet1::table::WideTable`;
  const rec = await store.get(id);
  assert.ok(rec);

  const text = rec.metadata.text;
  assert.match(text, /… \(\+7 more columns\)/);
  assert.doesNotMatch(text, /Header10/);
  assert.doesNotMatch(text, /Value10/);
});

