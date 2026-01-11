import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../../../../../packages/power-query/src/engine.js";
import { HttpConnector } from "../../../../../packages/power-query/src/connectors/http.js";

import { DocumentController } from "../../document/documentController.js";
import { MockEngine } from "../../document/engine.js";

import { applyQueryToDocument } from "../applyToDocument.ts";
import { dateToExcelSerial } from "../../shared/valueParsing.js";
import { MS_PER_DAY, PqDateTimeZone, PqDecimal, PqDuration, PqTime } from "../../../../../packages/power-query/src/values.js";

test("applyQueryToDocument requests non-materializing streaming execution", async () => {
  const engine = {
    executeQueryStreaming: async (_query, _context, opts) => {
      assert.equal(opts.materialize, false);
      await opts.onBatch({ rowOffset: 0, values: [["A", "B"]] });
      await opts.onBatch({ rowOffset: 1, values: [[1, 2], [3, 4]] });
      return { schema: { columns: [{ name: "A", type: "any" }, { name: "B", type: "any" }], inferred: true }, rowCount: 2, columnCount: 2 };
    },
  };

  const doc = new DocumentController({ engine: new MockEngine() });

  const query = {
    id: "q_stream_stub",
    name: "Stream stub",
    source: { type: "range", range: { values: [["A", "B"]], hasHeaders: true } },
    steps: [],
  };

  const destination = { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true };

  const result = await applyQueryToDocument(doc, query, destination, { engine, batchSize: 5 });

  assert.deepEqual(result, { rows: 3, cols: 2 });
  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, "A");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 0 }).value, 1);
  assert.equal(doc.getCell("Sheet1", { row: 2, col: 1 }).value, 4);
});

test("applyQueryToDocument writes query output into the destination range (file CSV source)", async () => {
  const engine = new QueryEngine({
    fileAdapter: {
      readText: async () => ["Region,Sales", "East,100", "West,200"].join("\n"),
    },
  });

  const doc = new DocumentController({ engine: new MockEngine() });

  const query = {
    id: "q_sales",
    name: "Sales",
    source: { type: "csv", path: "/tmp/sales.csv", options: { hasHeaders: true } },
    steps: [],
  };

  const destination = { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true };

  const result = await applyQueryToDocument(doc, query, destination, { engine, batchSize: 1 });

  assert.deepEqual(result, { rows: 3, cols: 2 });
  assert.deepEqual(destination.lastOutputSize, { rows: 3, cols: 2 });

  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, "Region");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 0 }).value, "East");
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 1 }).value, 100);
  assert.equal(doc.getCell("Sheet1", { row: 2, col: 1 }).value, 200);
});

test("applyQueryToDocument loads formula-like strings as values (not formulas)", async () => {
  const engine = new QueryEngine();
  const doc = new DocumentController({ engine: new MockEngine() });

  const query = {
    id: "q_text",
    name: "Text",
    source: {
      type: "range",
      range: {
        values: [["=Header"], ["=1+1"], ["'literal"]],
        hasHeaders: true,
      },
    },
    steps: [],
  };

  const destination = { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true };
  await applyQueryToDocument(doc, query, destination, { engine, batchSize: 2 });

  const header = doc.getCell("Sheet1", { row: 0, col: 0 });
  assert.equal(header.value, "=Header");
  assert.equal(header.formula, null);

  const formulaLike = doc.getCell("Sheet1", { row: 1, col: 0 });
  assert.equal(formulaLike.value, "=1+1");
  assert.equal(formulaLike.formula, null);

  const apostrophe = doc.getCell("Sheet1", { row: 2, col: 0 });
  assert.equal(apostrophe.value, "'literal");
  assert.equal(apostrophe.formula, null);
});

test("applyQueryToDocument converts Date objects into Excel serial numbers", async () => {
  const engine = new QueryEngine();
  const doc = new DocumentController({ engine: new MockEngine() });

  const when = new Date("2020-01-01T00:00:00.000Z");

  const query = {
    id: "q_date",
    name: "Date",
    source: {
      type: "range",
      range: {
        values: [["When"], [when]],
        hasHeaders: true,
      },
    },
    steps: [],
  };

  const destination = { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true };
  await applyQueryToDocument(doc, query, destination, { engine, batchSize: 2 });

  const cell = doc.getCell("Sheet1", { row: 1, col: 0 });
  assert.equal(cell.value, dateToExcelSerial(when));
  assert.equal(cell.formula, null);
});

test("applyQueryToDocument converts datetime/time/duration values into Excel serial numbers", async () => {
  const engine = new QueryEngine();
  const doc = new DocumentController({ engine: new MockEngine() });

  const when = new Date("2020-01-01T12:00:00.000Z");
  const time = new PqTime(6 * 60 * 60 * 1000 + 30 * 60 * 1000);
  const duration = new PqDuration(1.5 * MS_PER_DAY);

  const query = {
    id: "q_datetime_time_duration",
    name: "Datetime/Time/Duration",
    source: {
      type: "range",
      range: {
        values: [
          ["When", "Time", "Duration"],
          [when, time, duration],
        ],
        hasHeaders: true,
      },
    },
    steps: [],
  };

  const destination = { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true };
  await applyQueryToDocument(doc, query, destination, { engine, batchSize: 2 });

  assert.equal(doc.getCell("Sheet1", { row: 1, col: 0 }).value, dateToExcelSerial(when));
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 1 }).value, time.milliseconds / MS_PER_DAY);
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 2 }).value, duration.milliseconds / MS_PER_DAY);
});

test("applyQueryToDocument converts decimal wrapper values into numbers", async () => {
  const engine = new QueryEngine();
  const doc = new DocumentController({ engine: new MockEngine() });

  const query = {
    id: "q_decimal",
    name: "Decimal",
    source: {
      type: "range",
      range: {
        values: [["Dec"], [new PqDecimal("123.450")]],
        hasHeaders: true,
      },
    },
    steps: [],
  };

  const destination = { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true };
  await applyQueryToDocument(doc, query, destination, { engine, batchSize: 2 });

  assert.equal(doc.getCell("Sheet1", { row: 1, col: 0 }).value, 123.45);
});

test("applyQueryToDocument converts datetimezone and binary values into document-safe scalars", async () => {
  const engine = new QueryEngine();
  const doc = new DocumentController({ engine: new MockEngine() });

  const dtz = PqDateTimeZone.from("2020-01-01T12:00:00.000+02:00");
  assert.ok(dtz, "expected datetimezone literal to parse");
  const bytes = new Uint8Array([1, 2, 3]);

  const query = {
    id: "q_dtz_binary",
    name: "Datetimezone/Binary",
    source: {
      type: "range",
      range: {
        values: [["Zone", "Bin"], [dtz, bytes]],
        hasHeaders: true,
      },
    },
    steps: [],
  };

  const destination = { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true };
  await applyQueryToDocument(doc, query, destination, { engine, batchSize: 2 });

  assert.equal(doc.getCell("Sheet1", { row: 1, col: 0 }).value, dateToExcelSerial(new Date("2020-01-01T10:00:00.000Z")));
  assert.equal(doc.getCell("Sheet1", { row: 1, col: 1 }).value, "AQID");
});

test("applyQueryToDocument supports HTTP sources via injected HttpConnector", async () => {
  const fetchStub = async () =>
    new Response(["Name,Score", "Alice,10", "Bob,20"].join("\n"), {
      status: 200,
      headers: { "content-type": "text/csv" },
    });

  const engine = new QueryEngine({
    connectors: {
      http: new HttpConnector({ fetch: fetchStub }),
    },
  });

  const doc = new DocumentController({ engine: new MockEngine() });

  const query = {
    id: "q_http",
    name: "HTTP CSV",
    source: { type: "api", url: "https://example.test/data.csv", method: "GET" },
    steps: [],
  };

  const destination = { sheetId: "Sheet1", start: { row: 1, col: 1 }, includeHeader: true };

  await applyQueryToDocument(doc, query, destination, { engine, batchSize: 2 });

  assert.equal(doc.getCell("Sheet1", { row: 1, col: 1 }).value, "Name");
  assert.equal(doc.getCell("Sheet1", { row: 2, col: 2 }).value, 10);
  assert.equal(doc.getCell("Sheet1", { row: 3, col: 1 }).value, "Bob");
});

test("applyQueryToDocument stops and reverts writes when cancelled", async () => {
  const engine = new QueryEngine();
  const doc = new DocumentController({ engine: new MockEngine() });

  const query = {
    id: "q_cancel",
    name: "Cancel",
    source: {
      type: "range",
      range: {
        values: [
          ["A", "B"],
          [1, 2],
          [3, 4],
          [5, 6],
          [7, 8],
        ],
        hasHeaders: true,
      },
    },
    steps: [],
  };

  const destination = { sheetId: "Sheet1", start: { row: 0, col: 0 }, includeHeader: true };

  const controller = new AbortController();
  const promise = applyQueryToDocument(doc, query, destination, {
    engine,
    batchSize: 1,
    signal: controller.signal,
    onProgress: (evt) => {
      if (evt.type === "batch" && evt.totalRowsWritten >= 1) {
        controller.abort();
      }
    },
  });

  await assert.rejects(promise, (err) => err?.name === "AbortError");

  // The batch should have been cancelled, leaving the sheet untouched.
  assert.equal(doc.getCell("Sheet1", { row: 0, col: 0 }).value, null);
  assert.equal((doc).batchDepth, 0);
});
