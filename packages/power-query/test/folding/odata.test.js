import assert from "node:assert/strict";
import test from "node:test";

import { ODataFoldingEngine, buildODataUrl } from "../../src/folding/odata.js";

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

test("OData folding: folds removeColumns when a projection is known", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_remove_cols",
    name: "OData remove cols",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      { id: "s1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Name", "Price"] } },
      { id: "s2", name: "Remove", operation: { type: "removeColumns", columns: ["Price"] } },
      { id: "s3", name: "Take", operation: { type: "take", count: 5 } },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "odata");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$select=Id,Name&$top=5");
  assert.deepEqual(
    explained.steps.map((s) => s.status),
    ["folded", "folded", "folded"],
  );
});

test("OData folding: removeColumns can fold against source URL $select", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_remove_cols_base",
    name: "OData remove cols base",
    source: { type: "odata", url: "https://example.com/odata/Products?$select=Id,Name,Price" },
    steps: [{ id: "s1", name: "Remove", operation: { type: "removeColumns", columns: ["Price"] } }],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "odata");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$select=Id,Name");
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

test("buildODataUrl: preserves existing query options when not overridden", () => {
  assert.equal(
    buildODataUrl("https://example.com/odata/Products?$top=5&foo=bar", {}),
    "https://example.com/odata/Products?$top=5&foo=bar",
  );
  assert.equal(
    buildODataUrl("https://example.com/odata/Products?$top=5&foo=bar", { top: 2 }),
    "https://example.com/odata/Products?foo=bar&$top=2",
  );
  assert.equal(
    buildODataUrl("https://example.com/odata/Products?$TOP=5&foo=bar", { top: 2 }),
    "https://example.com/odata/Products?foo=bar&$top=2",
  );
});

test("OData folding: respects $top embedded in the source URL", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_base_top",
    name: "OData base $top",
    source: { type: "odata", url: "https://example.com/odata/Products?$top=5" },
    steps: [{ id: "s1", name: "Take", operation: { type: "take", count: 10 } }],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "odata");
  // take(10) should not expand past the base URL's $top=5.
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$top=5");
});

test("OData folding: combines source URL $filter with folded filterRows", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_base_filter",
    name: "OData base $filter",
    source: { type: "odata", url: "https://example.com/odata/Products?$filter=Price%20gt%2020" },
    steps: [
      { id: "s1", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "Price", operator: "greaterThan", value: 30 } } },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "odata");
  assert.ok(explained.plan.url.includes("$filter="));
  assert.ok(explained.plan.url.includes("Price%20gt%2020"));
  assert.ok(explained.plan.url.includes("Price%20gt%2030"));
});

test("OData folding: pushes skip into query options", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_skip",
    name: "OData skip",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      { id: "s1", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "Price", operator: "greaterThan", value: 20 } } },
      { id: "s2", name: "Skip", operation: { type: "skip", count: 10 } },
      { id: "s3", name: "Take", operation: { type: "take", count: 5 } },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "odata");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$filter=Price%20gt%2020&$skip=10&$top=5");
});

test("OData folding: rewrites take then skip into $skip+$top", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_take_skip",
    name: "OData take then skip",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      { id: "s1", name: "Take", operation: { type: "take", count: 10 } },
      { id: "s2", name: "Skip", operation: { type: "skip", count: 3 } },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "odata");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$skip=3&$top=7");
});

test("OData folding: does not fold filterRows after skip (preserves local semantics)", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_skip_then_filter",
    name: "OData skip then filter",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      { id: "s1", name: "Skip", operation: { type: "skip", count: 5 } },
      {
        id: "s2",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Price", operator: "greaterThan", value: 20 } },
      },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "hybrid");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$skip=5");
  assert.deepEqual(explained.steps.map((s) => s.status), ["folded", "local"]);
});

test("OData folding: does not fold filterRows when source URL includes $skip", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_base_skip_filter",
    name: "OData base skip then filter",
    source: { type: "odata", url: "https://example.com/odata/Products?$skip=5" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Price", operator: "greaterThan", value: 20 } },
      },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "local");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$skip=5");
  assert.deepEqual(explained.steps.map((s) => s.status), ["local"]);
});

test("OData folding: does not fold filterRows after take (preserves local semantics)", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_take_then_filter",
    name: "OData take then filter",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [
      { id: "s1", name: "Take", operation: { type: "take", count: 5 } },
      {
        id: "s2",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Price", operator: "greaterThan", value: 20 } },
      },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "hybrid");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$top=5");
  assert.deepEqual(
    explained.steps.map((s) => s.status),
    ["folded", "local"],
  );
});

test("OData folding: does not fold filterRows when source URL includes $top", () => {
  const folding = new ODataFoldingEngine();
  const query = {
    id: "q_odata_base_top_filter",
    name: "OData base top then filter",
    source: { type: "odata", url: "https://example.com/odata/Products?$top=5" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Price", operator: "greaterThan", value: 20 } },
      },
    ],
  };

  const explained = folding.explain(/** @type {any} */ (query));
  assert.equal(explained.plan.type, "local");
  assert.equal(explained.plan.url, "https://example.com/odata/Products?$top=5");
  assert.deepEqual(explained.steps.map((s) => s.status), ["local"]);
});
