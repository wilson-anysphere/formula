import assert from "node:assert/strict";
import test from "node:test";

import { ODataFoldingEngine } from "../../src/folding/odata.js";

test("OData folding: pushes select/filter/orderby/top into query options", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_fold",
    name: "OData fold",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Price", operator: "greaterThan", value: 20 } },
      },
      { id: "s2", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "Price", direction: "descending" }] } },
      { id: "s3", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Name"] } },
      { id: "s4", name: "Take", operation: { type: "take", count: 10 } },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "odata");
  assert.equal(
    explained.plan.url,
    "https://example.com/odata/Products?$select=Id,Name&$filter=Price%20gt%2020&$orderby=Price%20desc&$top=10",
  );
  assert.deepEqual(
    explained.steps.map((s) => s.status),
    ["folded", "folded", "folded", "folded"],
  );
});

test("OData folding: falls back to local when an operation is unsupported", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_hybrid",
    name: "OData hybrid",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Name"] } },
      {
        id: "s2",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Name", operator: "equals", value: { bad: true } } },
      },
      { id: "s3", name: "Take", operation: { type: "take", count: 5 } },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "hybrid");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$select=Id,Name");
  assert.deepEqual(
    explained.steps.map((s) => s.status),
    ["folded", "local", "local"],
  );
  assert.equal(explained.steps[1].reason, "unsupported_predicate");
  assert.equal(explained.steps[2].reason, "folding_stopped");
});

test("OData folding: contains defaults to case-insensitive and casts values to text", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_contains",
    name: "OData contains",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Name", operator: "contains", value: "ABC" } },
      },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "odata");
  assert.equal(
    explained.plan.url,
    "https://example.com/odata/Products?$filter=contains(tolower(cast(Name,Edm.String)),%20tolower(%27ABC%27))",
  );
});

test("OData folding: equals ignores caseSensitive and stays case-sensitive", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_equals_case",
    name: "OData equals case",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: {
          type: "filterRows",
          predicate: { type: "comparison", column: "Name", operator: "equals", value: "ABC", caseSensitive: false },
        },
      },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "odata");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$filter=Name%20eq%20%27ABC%27");
});

test("OData folding: empty contains needle breaks folding (preserves local semantics)", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_contains_empty",
    name: "OData contains empty",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      { id: "s1", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "Name", operator: "contains", value: "" } } },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "local");
  assert.equal(explained.plan.url, "https://example.com/odata/Products");
  assert.equal(explained.steps[0].status, "local");
  assert.equal(explained.steps[0].reason, "unsupported_predicate");
});

test("OData folding: predicate columns must exist after selectColumns (preserves local errors)", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_missing_col",
    name: "OData missing col",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id"] } },
      { id: "s2", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "Price", operator: "greaterThan", value: 0 } } },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "hybrid");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$select=Id");
  assert.deepEqual(
    explained.steps.map((s) => s.status),
    ["folded", "local"],
  );
});
