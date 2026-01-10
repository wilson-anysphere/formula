import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { parquetToArrowTable } from "@formula/data-io";

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

test("Parquet import writes batches into DocumentController (with header)", async () => {
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

test("Parquet import can omit header row", async () => {
  const parquetBytes = new Uint8Array(await readFile(FIXTURE_URL));

  const doc = new DocumentController({ engine: new MockEngine() });
  await importParquetIntoDocument(doc, "Sheet1", "B2", parquetBytes, {
    batchSize: 2,
    includeHeader: false,
  });

  assert.equal(doc.getCell("Sheet1", { row: 1, col: 1 }).value, 1);
  assert.equal(doc.getCell("Sheet1", { row: 0, col: 1 }).value, null);
});

test("Parquet export from document range produces a readable parquet file", async () => {
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

