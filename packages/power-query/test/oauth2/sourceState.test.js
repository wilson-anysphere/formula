import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { HttpConnector } from "../../src/connectors/http.js";
import { QueryEngine } from "../../src/engine.js";
import { DataTable } from "../../src/table.js";

test("HttpConnector.getSourceState injects OAuth2 bearer token (and retries once on 401)", async () => {
  /** @type {any[]} */
  const tokenCalls = [];

  const oauth2Manager = {
    getAccessToken: async (opts) => {
      tokenCalls.push(opts);
      return { accessToken: opts.forceRefresh ? "token-2" : "token-1", expiresAtMs: null, refreshToken: null };
    },
  };

  let headCalls = 0;
  /** @type {typeof fetch} */
  const fetchFn = async (_url, init) => {
    assert.equal(init?.method, "HEAD");
    const auth = /** @type {any} */ (init?.headers)?.Authorization;
    headCalls++;
    if (auth === "Bearer token-1") {
      return new Response("unauthorized", { status: 401 });
    }
    assert.equal(auth, "Bearer token-2");
    return new Response("", {
      status: 200,
      headers: {
        etag: "W/\"123\"",
        "last-modified": "Mon, 01 Jan 2024 00:00:00 GMT",
      },
    });
  };

  const connector = new HttpConnector({ fetch: fetchFn, oauth2Manager });
  const state = await connector.getSourceState(
    { url: "https://api.example/data", auth: { type: "oauth2", providerId: "example" } },
    { now: () => 0 },
  );

  assert.equal(headCalls, 2);
  assert.equal(tokenCalls.length, 2);
  assert.equal(tokenCalls[0].providerId, "example");
  assert.equal(tokenCalls[0].forceRefresh, false);
  assert.equal(tokenCalls[1].forceRefresh, true);
  assert.equal(state.etag, "W/\"123\"");
  assert.ok(state.sourceTimestamp instanceof Date);
});

test("QueryEngine cache validation includes API auth in source-state targets", async () => {
  const cache = new CacheManager({ store: new MemoryCacheStore() });
  let executeCalls = 0;

  const http = {
    id: "http",
    permissionKind: "http:request",
    getCacheKey: (request) => {
      const base = { connector: "http", url: request.url, method: request.method ?? "GET" };
      if (request.auth?.type === "oauth2") {
        // @ts-ignore - stable JSON shape
        base.auth = { type: "oauth2", providerId: request.auth.providerId };
      }
      return base;
    },
    getSourceState: async (_request) => ({}),
    execute: async (_request, options = {}) => {
      executeCalls++;
      const now = options.now ?? (() => Date.now());
      const table = DataTable.fromGrid([["id"], [1]], { hasHeaders: true, inferTypes: true });
      return {
        table,
        meta: {
          refreshedAt: new Date(now()),
          schema: { columns: table.columns, inferred: true },
          rowCount: table.rows.length,
          rowCountEstimate: table.rows.length,
          provenance: { kind: "http", url: "https://api.example/data", method: "GET" },
        },
      };
    },
  };

  const engine = new QueryEngine({ cache, connectors: { http } });
  const query = {
    id: "q1",
    name: "Query 1",
    source: { type: "api", url: "https://api.example/data", method: "GET", auth: { type: "oauth2", providerId: "example" } },
    steps: [],
  };

  const first = await engine.executeQuery(query, {}, {});
  const second = await engine.executeQuery(query, {}, {});

  assert.deepEqual(first.toGrid(), [["id"], [1]]);
  assert.deepEqual(second.toGrid(), [["id"], [1]]);
  assert.equal(executeCalls, 1, "second execution should be a cache hit");
});

