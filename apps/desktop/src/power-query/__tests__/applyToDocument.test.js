import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../../../../../packages/power-query/src/engine.js";
import { HttpConnector } from "../../../../../packages/power-query/src/connectors/http.js";

import { DocumentController } from "../../document/documentController.js";
import { MockEngine } from "../../document/engine.js";

import { applyQueryToDocument } from "../applyToDocument.ts";
import { dateToExcelSerial } from "../../shared/valueParsing.js";

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
