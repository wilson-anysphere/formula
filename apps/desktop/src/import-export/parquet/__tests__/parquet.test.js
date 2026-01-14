import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { arrowTableFromColumns, arrowTableToParquet, parquetToArrowTable } from "@formula/data-io";

import { DocumentController } from "../../../document/documentController.js";
import { MockEngine } from "../../../document/engine.js";
import {
  exportDocumentRangeToParquet,
  importParquetIntoDocument,
} from "../index.js";

const FIXTURE_URL = new URL(
  "../../../../../../packages/data-io/test/fixtures/simple.parquet",
  import.meta.url
);

let parquetAvailable = true;
try {
  // Validate Parquet support is actually usable via the data-io helpers (pnpm workspaces
  // don't necessarily hoist `apache-arrow`/`parquet-wasm` to the repo root).
  await arrowTableToParquet(arrowTableFromColumns({ __probe: new Int32Array([1]) }));
} catch {
  parquetAvailable = false;
}

test("Parquet import writes batches into DocumentController (with header)", { skip: !parquetAvailable }, async () => {
  const parquetBytes = new Uint8Array(await readFile(FIXTURE_URL));

  const doc = new DocumentController({ engine: new MockEngine() });
  await importParquetIntoDocument(doc, "Sheet1", "A1", parquetBytes, {
    batchSize: 2,
    includeHeader: true,
  });

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, "id");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 1 }).value, "Alice");
  assert.equal(doc.getCell("Sheet1", { row: 3, col: 3 }).value, 3.75);
});

test("Parquet import can omit header row", { skip: !parquetAvailable }, async () => {
  const parquetBytes = new Uint8Array(await readFile(FIXTURE_URL));

  const doc = new DocumentController({ engine: new MockEngine() });
  await importParquetIntoDocument(doc, "Sheet1", "B2", parquetBytes, {
    batchSize: 2,
    includeHeader: false,
  });

  assert.equal(doc.getCell("Sheet1", { row: 1, col: 1 }).value, 1);
  assert.equal(doc.getCell("Sheet1", { row: 0, col: 1 }).value, null);
});

test("Parquet export from document range produces a readable parquet file", { skip: !parquetAvailable }, async () => {
  const parquetBytes = new Uint8Array(await readFile(FIXTURE_URL));

  const doc = new DocumentController({ engine: new MockEngine() });
  await importParquetIntoDocument(doc, "Sheet1", "A1", parquetBytes, {
    batchSize: 2,
    includeHeader: true,
  });

  const exported = await exportDocumentRangeToParquet(doc, "Sheet1", "A1:D4", {
    headerRow: true,
    compression: "UNCOMPRESSED",
  });

  const roundTrip = await parquetToArrowTable(exported, { batchSize: 2 });
  assert.deepEqual(
    roundTrip.schema.fields.map((f) => f.name),
    ["id", "name", "active", "score"]
  );
  assert.equal(roundTrip.numRows, 3);
  assert.equal(roundTrip.getChildAt(1).get(2), "Carla");
  assert.equal(roundTrip.getChildAt(3).get(1), 2.25);
});

test("Parquet export handles rich text + in-cell image values", { skip: !parquetAvailable }, async () => {
  const doc = new DocumentController({ engine: new MockEngine() });
  doc.setRangeValues("Sheet1", "A1:B2", [
    [
      { text: "Name", runs: [{ start: 0, end: 4, style: { bold: true } }] },
      { type: "image", value: { imageId: "img-header", altText: "Logo" } },
    ],
    ["Alice", { type: "image", value: { imageId: "img-row" } }],
  ]);

  const exported = await exportDocumentRangeToParquet(doc, "Sheet1", "A1:B2", {
    headerRow: true,
    compression: "UNCOMPRESSED",
  });

  const roundTrip = await parquetToArrowTable(exported, { batchSize: 2 });
  assert.deepEqual(roundTrip.schema.fields.map((f) => f.name), ["Name", "Logo"]);
  assert.equal(roundTrip.numRows, 1);
  assert.equal(roundTrip.getChildAt(0).get(0), "Alice");
  assert.equal(roundTrip.getChildAt(1).get(0), "[Image]");
});

test("Parquet export rejects huge ranges before scanning cells", async () => {
  let scanned = 0;
  const doc = {
    getCell() {
      scanned += 1;
      throw new Error("Should not scan");
    },
  };

  await assert.rejects(
    () => exportDocumentRangeToParquet(doc, "Sheet1", "A1:Z8000"),
    /Range too large to export to Parquet/i,
  );
  assert.equal(scanned, 0);
});

test("Parquet export does not resurrect deleted sheets when called with a stale sheet id (no phantom creation)", async () => {
  const doc = new DocumentController({ engine: new MockEngine() });

  // Ensure Sheet1 exists so deleting Sheet2 doesn't trip the last-sheet guard.
  doc.getCell("Sheet1", { row: 0, col: 0 });
  doc.setCellValue("Sheet2", { row: 0, col: 0 }, "two");
  assert.deepEqual(doc.getSheetIds(), ["Sheet1", "Sheet2"]);

  doc.deleteSheet("Sheet2");
  assert.deepEqual(doc.getSheetIds(), ["Sheet1"]);

  await assert.rejects(() => exportDocumentRangeToParquet(doc, "Sheet2", "A1"), /Unknown sheet/i);
  assert.deepEqual(doc.getSheetIds(), ["Sheet1"]);
});
