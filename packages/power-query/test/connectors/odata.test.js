import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "../../src/table.js";
import { QueryEngine } from "../../src/engine.js";
import { ODataConnector } from "../../src/connectors/odata.js";
import { getFileSourceId, getHttpSourceId, getSourceIdForProvenance, getSourceIdForQuerySource } from "../../src/privacy/sourceId.js";

/**
 * @param {unknown} json
 * @param {{ status?: number; headers?: Record<string, string> }} [options]
 */
function makeJsonResponse(json, options = {}) {
  const status = options.status ?? 200;
  const headerMap = new Map(Object.entries(options.headers ?? {}));
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: {
      get(name) {
        return headerMap.get(String(name).toLowerCase()) ?? headerMap.get(String(name)) ?? null;
      },
    },
    async json() {
      return json;
    },
  };
}

test("ODataConnector: follows @odata.nextLink pagination", async () => {
  /** @type {string[]} */
  const urls = [];

  const connector = new ODataConnector({
    fetch: async (url) => {
      urls.push(String(url));
      if (urls.length === 1) {
        return makeJsonResponse({
          value: [
            { Id: 1, Name: "A" },
            { Id: 2, Name: "B" },
          ],
          "@odata.nextLink": "Products?$skiptoken=abc",
        });
      }
      return makeJsonResponse({
        value: [{ Id: 3, Name: "C" }],
      });
    },
  });

  const result = await connector.execute({ url: "https://example.com/odata/Products" });
  assert.equal(urls.length, 2);
  assert.deepEqual(result.table.toGrid(), [
    ["Id", "Name"],
    [1, "A"],
    [2, "B"],
    [3, "C"],
  ]);
});

test("ODataConnector: respects $top for pagination short-circuiting", async () => {
  /** @type {string[]} */
  const urls = [];

  const connector = new ODataConnector({
    fetch: async (url) => {
      urls.push(String(url));
      return makeJsonResponse({
        value: [
          { Id: 1, Name: "A" },
          { Id: 2, Name: "B" },
        ],
        "@odata.nextLink": "Products?$skiptoken=abc",
      });
    },
  });

  const result = await connector.execute({ url: "https://example.com/odata/Products", query: { top: 2 } });
  assert.equal(urls.length, 1, "expected no nextLink follow-up once $top rows are collected");
  assert.equal(urls[0], "https://example.com/odata/Products?$top=2");
  assert.deepEqual(result.table.toGrid(), [
    ["Id", "Name"],
    [1, "A"],
    [2, "B"],
  ]);
});

test("ODataConnector: respects $top embedded in the URL for pagination short-circuiting", async () => {
  /** @type {string[]} */
  const urls = [];

  const connector = new ODataConnector({
    fetch: async (url) => {
      urls.push(String(url));
      if (urls.length === 1) {
        return makeJsonResponse({
          value: [{ Id: 1 }, { Id: 2 }],
          "@odata.nextLink": "Products?$skiptoken=a",
        });
      }
      if (urls.length === 2) {
        return makeJsonResponse({
          value: [{ Id: 3 }, { Id: 4 }],
          "@odata.nextLink": "Products?$skiptoken=b",
        });
      }
      if (urls.length === 3) {
        return makeJsonResponse({
          value: [{ Id: 5 }, { Id: 6 }],
          "@odata.nextLink": "Products?$skiptoken=c",
        });
      }
      return makeJsonResponse({
        value: [{ Id: 7 }, { Id: 8 }],
      });
    },
  });

  const result = await connector.execute({ url: "https://example.com/odata/Products?$top=5" });
  assert.equal(urls.length, 3, "expected pagination to stop once the URL's $top rows are collected");
  assert.equal(urls[0], "https://example.com/odata/Products?$top=5");
  assert.equal(urls[1], "https://example.com/odata/Products?$skiptoken=a");
  assert.deepEqual(result.table.toGrid(), [["Id"], [1], [2], [3], [4], [5]]);
});

test("ODataConnector: preserves $select column order and schema (even when empty)", async () => {
  /** @type {string[]} */
  const urls = [];

  const connector = new ODataConnector({
    fetch: async (url) => {
      urls.push(String(url));
      return makeJsonResponse({ value: [] });
    },
  });

  const result = await connector.execute({
    url: "https://example.com/odata/Products",
    query: { select: ["Name", "Id"] },
  });

  assert.equal(urls[0], "https://example.com/odata/Products?$select=Name,Id");
  assert.deepEqual(result.table.toGrid(), [["Name", "Id"]]);
});

test("ODataConnector: uses $select column order for row materialization", async () => {
  const connector = new ODataConnector({
    fetch: async () =>
      makeJsonResponse({
        value: [{ Id: 1, Name: "A" }],
      }),
  });

  const result = await connector.execute({
    url: "https://example.com/odata/Products",
    query: { select: ["Name", "Id"] },
  });

  assert.deepEqual(result.table.toGrid(), [
    ["Name", "Id"],
    ["A", 1],
  ]);
});

test("ODataConnector: preserves $select embedded in the URL for schema inference", async () => {
  const connector = new ODataConnector({
    fetch: async () => makeJsonResponse({ value: [] }),
  });

  const result = await connector.execute({
    url: "https://example.com/odata/Products?$select=Name,Id",
  });

  assert.deepEqual(result.table.toGrid(), [["Name", "Id"]]);
});

test("ODataConnector: tolerates payloads that are a single object (no value wrapper)", async () => {
  const connector = new ODataConnector({
    fetch: async () => makeJsonResponse({ Id: 1, Name: "A" }),
  });

  const result = await connector.execute({ url: "https://example.com/odata/Products(1)" });
  assert.deepEqual(result.table.toGrid(), [
    ["Id", "Name"],
    [1, "A"],
  ]);
});

test("ODataConnector: supports rowsPath/jsonPath overrides for custom payload envelopes", async () => {
  const connector = new ODataConnector({
    fetch: async () =>
      makeJsonResponse({
        data: {
          items: [{ Id: 1 }, { Id: 2 }],
        },
      }),
  });

  const result = await connector.execute({ url: "https://example.com/odata/Products", rowsPath: "data.items" });
  assert.deepEqual(result.table.toGrid(), [["Id"], [1], [2]]);

  const resultViaAlias = await connector.execute({ url: "https://example.com/odata/Products", jsonPath: "data.items" });
  assert.deepEqual(resultViaAlias.table.toGrid(), [["Id"], [1], [2]]);
});

test("ODataConnector: supports legacy d.results payload shape", async () => {
  const connector = new ODataConnector({
    fetch: async () =>
      makeJsonResponse({
        d: {
          results: [{ Id: 1 }, { Id: 2 }],
        },
      }),
  });

  const result = await connector.execute({ url: "https://example.com/odata/Products" });
  assert.deepEqual(result.table.toGrid(), [["Id"], [1], [2]]);
});

test("ODataConnector: getSourceState returns known source state on 304", async () => {
  const knownEtag = '"v1"';
  const knownSourceTimestamp = new Date("2024-01-01T00:00:00.000Z");
  let headerChecked = false;

  /** @type {typeof fetch} */
  const fetchMock = async (_url, init) => {
    const method = String(init?.method ?? "GET").toUpperCase();
    if (method !== "HEAD") throw new Error("expected HEAD");
    const ifNoneMatch = Object.entries(/** @type {any} */ (init?.headers ?? {})).find(([name]) => name.toLowerCase() === "if-none-match")?.[1];
    assert.equal(ifNoneMatch, knownEtag);
    headerChecked = true;
    return {
      ok: false,
      status: 304,
      headers: { get: () => null },
      async json() {
        return {};
      },
    };
  };

  const connector = new ODataConnector({ fetch: fetchMock });
  const state = await connector.getSourceState({ url: "https://example.com/odata/Products" }, { knownEtag, knownSourceTimestamp });
  assert.equal(headerChecked, true);
  assert.deepEqual(state, { etag: knownEtag, sourceTimestamp: knownSourceTimestamp });
});

test("privacy ids: OData sources map to stable http source ids", () => {
  const source = { type: "odata", url: "https://example.com/odata/Products" };
  const sourceId = getSourceIdForQuerySource(/** @type {any} */ (source));
  assert.equal(sourceId, getHttpSourceId(source.url));

  const provId = getSourceIdForProvenance({ kind: "odata", url: "https://example.com/odata/Products?$top=1", method: "GET" });
  assert.equal(provId, sourceId);
});

test("QueryEngine: privacy firewall blocks combining OData + CSV across privacy levels", async () => {
  const privateCsvPath = "/tmp/private.csv";
  const publicODataUrl = "https://public.example.com/odata/Products";

  const engine = new QueryEngine({
    privacyMode: "enforce",
    fileAdapter: {
      readText: async () => ["Id,Region", "1,East"].join("\n"),
    },
    connectors: {
      odata: {
        id: "odata",
        permissionKind: "http:request",
        getCacheKey: (req) => req,
        execute: async () => {
          const table = DataTable.fromGrid(
            [
              ["Id", "Target"],
              [1, 10],
            ],
            { hasHeaders: true, inferTypes: true },
          );
          return {
            table,
            meta: {
              refreshedAt: new Date(0),
              schema: { columns: table.columns, inferred: true },
              rowCount: table.rowCount,
              rowCountEstimate: table.rowCount,
              provenance: { kind: "odata", url: publicODataUrl, method: "GET" },
            },
          };
        },
      },
    },
  });

  const publicQuery = {
    id: "q_public_odata",
    name: "Public OData",
    source: { type: "odata", url: publicODataUrl },
    steps: [],
  };

  const privateQuery = {
    id: "q_private_csv",
    name: "Private CSV",
    source: { type: "csv", path: privateCsvPath },
    steps: [
      {
        id: "s_merge",
        name: "Merge",
        operation: { type: "merge", rightQuery: "q_public_odata", joinType: "left", leftKey: "Id", rightKey: "Id" },
      },
    ],
  };

  const csvSourceId = getFileSourceId(privateCsvPath);
  const odataSourceId = getHttpSourceId(publicODataUrl);

  await assert.rejects(
    () =>
      engine.executeQuery(privateQuery, {
        queries: { q_public_odata: publicQuery },
        privacy: { levelsBySourceId: { [csvSourceId]: "private", [odataSourceId]: "public" } },
      }),
    /Formula\.Firewall/,
  );
});
