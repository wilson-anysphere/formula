import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "../../src/table.js";
import { FileConnector } from "../../src/connectors/file.js";
import { HttpConnector } from "../../src/connectors/http.js";
import { SqlConnector } from "../../src/connectors/sql.js";

test("FileConnector: CSV returns table + metadata", async () => {
  const connector = new FileConnector({
    readText: async () => ["Region,Sales", "East,100", "West,200"].join("\n"),
  });

  const { table, meta } = await connector.execute({ format: "csv", path: "/tmp/sales.csv", csv: { hasHeaders: true } });
  assert.equal(table.rows.length, 2);
  assert.deepEqual(table.columns.map((c) => c.name), ["Region", "Sales"]);
  assert.equal(meta.rowCount, 2);
  assert.equal(meta.provenance.kind, "file");
  assert.equal(meta.provenance.path, "/tmp/sales.csv");
  assert.equal(meta.provenance.format, "csv");
  assert.ok(meta.refreshedAt instanceof Date);
});

test("FileConnector: JSON supports jsonPath selection", async () => {
  const connector = new FileConnector({
    readText: async () => JSON.stringify({ data: [{ a: 1 }, { a: 2 }] }),
  });

  const { table, meta } = await connector.execute({
    format: "json",
    path: "/tmp/data.json",
    json: { jsonPath: "data" },
  });

  assert.deepEqual(table.toGrid(), [["a"], [1], [2]]);
  assert.equal(meta.rowCount, 2);
  assert.equal(meta.provenance.format, "json");
  assert.equal(meta.provenance.jsonPath, "data");
});

test("HttpConnector: merges credential headers when using fetchTable adapter", async () => {
  /** @type {Record<string, string>} */
  let observedHeaders = {};

  const connector = new HttpConnector({
    fetchTable: async (_url, options) => {
      observedHeaders = options.headers ?? {};
      return DataTable.fromGrid(
        [
          ["id", "value"],
          [1, "a"],
        ],
        { hasHeaders: true, inferTypes: true },
      );
    },
  });

  const { table, meta } = await connector.execute(
    { url: "https://example.com/api", method: "GET", headers: { "X-Test": "1" } },
    { credentials: { headers: { Authorization: "Bearer token" } } },
  );

  assert.equal(table.rows.length, 1);
  assert.equal(observedHeaders["X-Test"], "1");
  assert.equal(observedHeaders.Authorization, "Bearer token");
  assert.equal(meta.provenance.kind, "http");
  assert.equal(meta.provenance.url, "https://example.com/api");
});

test("SqlConnector: forwards credentials + signal to adapter", async () => {
  /** @type {any} */
  let observedOptions = null;

  const connector = new SqlConnector({
    querySql: async (_connection, _sql, options) => {
      observedOptions = options;
      return DataTable.fromGrid(
        [
          ["n"],
          [1],
        ],
        { hasHeaders: true, inferTypes: true },
      );
    },
  });

  const controller = new AbortController();
  const credentials = { user: "alice" };
  const { table, meta } = await connector.execute(
    { connection: { id: "db1" }, sql: "SELECT 1 AS n", params: [1, "x"] },
    { signal: controller.signal, credentials },
  );

  assert.equal(table.rows.length, 1);
  assert.deepEqual(observedOptions.params, [1, "x"]);
  assert.equal(observedOptions.credentials, credentials);
  assert.equal(observedOptions.signal, controller.signal);
  assert.equal(meta.provenance.kind, "sql");
  assert.equal(meta.provenance.sql, "SELECT 1 AS n");
});
