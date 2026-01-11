import assert from "node:assert/strict";
import test from "node:test";

import { ODataConnector } from "../../src/connectors/odata.js";
import { QueryEngine } from "../../src/engine.js";

/**
 * @param {unknown} json
 */
function jsonResponse(json) {
  /** @type {Map<string, string>} */
  const headers = new Map([["content-type", "application/json"]]);
  return {
    ok: true,
    status: 200,
    headers: {
      get(name) {
        return headers.get(String(name).toLowerCase()) ?? null;
      },
    },
    async json() {
      return json;
    },
  };
}

test("QueryEngine: executes fully folded OData plan (select/filter/orderby/top)", async () => {
  /** @type {string[]} */
  const urls = [];
  const connector = new ODataConnector({
    fetch: async (url) => {
      urls.push(String(url));
      return jsonResponse({ value: [{ Id: 2, Name: "B" }, { Id: 1, Name: "A" }] });
    },
  });

  const engine = new QueryEngine({
    connectors: { odata: connector },
  });

  const query = {
    id: "q_odata_folded",
    name: "OData folded",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Name"] } },
      {
        id: "s2",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Price", operator: "greaterThan", value: 20 } },
      },
      { id: "s3", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "Price", direction: "descending" }] } },
      { id: "s4", name: "Take", operation: { type: "take", count: 2 } },
    ],
  };

  const { table, meta } = await engine.executeQueryWithMeta(/** @type {any} */ (query), {}, {});
  assert.equal(urls.length, 1);
  assert.equal(
    urls[0],
    "https://example.com/odata/Products?$select=Id,Name&$filter=Price%20gt%2020&$orderby=Price%20desc&$top=2",
  );

  assert.ok(meta.folding, "expected folding metadata");
  assert.equal(meta.folding.kind, "odata");
  assert.equal(meta.folding.planType, "odata");
  assert.equal(meta.folding.url, urls[0]);
  assert.deepEqual(
    meta.folding.steps.map((s) => s.status),
    ["folded", "folded", "folded", "folded"],
  );

  // The engine should return the server-provided rows directly (no local steps).
  assert.deepEqual(table.toGrid(), [
    ["Id", "Name"],
    [2, "B"],
    [1, "A"],
  ]);
});

test("QueryEngine: executes hybrid OData plan and runs remaining steps locally", async () => {
  /** @type {string[]} */
  const urls = [];
  const connector = new ODataConnector({
    fetch: async (url) => {
      urls.push(String(url));
      return jsonResponse({ value: [{ Region: "East", Sales: 100 }, { Region: null, Sales: 150 }] });
    },
  });

  const engine = new QueryEngine({
    connectors: { odata: connector },
  });

  const query = {
    id: "q_odata_hybrid",
    name: "OData hybrid",
    source: { type: "odata", url: "https://example.com/odata/Sales" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Region", "Sales"] } },
      { id: "s2", name: "Fill Down", operation: { type: "fillDown", columns: ["Region"] } },
    ],
  };

  const { table, meta } = await engine.executeQueryWithMeta(/** @type {any} */ (query), {}, {});
  assert.equal(urls.length, 1);
  assert.equal(urls[0], "https://example.com/odata/Sales?$select=Region,Sales");

  assert.ok(meta.folding, "expected folding metadata");
  assert.equal(meta.folding.kind, "odata");
  assert.equal(meta.folding.planType, "hybrid");
  assert.equal(meta.folding.url, urls[0]);
  assert.deepEqual(
    meta.folding.steps.map((s) => s.status),
    ["folded", "local"],
  );
  assert.equal(meta.folding.steps[1].reason, "unsupported_op");
  assert.equal(meta.folding.localStepOffset, 1);

  // fillDown should run locally after the OData request.
  assert.deepEqual(table.toGrid(), [
    ["Region", "Sales"],
    ["East", 100],
    ["East", 150],
  ]);
});

