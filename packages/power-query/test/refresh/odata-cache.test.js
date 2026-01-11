import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { ODataConnector } from "../../src/connectors/odata.js";
import { QueryEngine } from "../../src/engine.js";

test("QueryEngine: folded OData queries reuse cache entries with source-state validation", async () => {
  let now = 0;
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => now });

  let etag = '"v1"';
  const initialEtag = etag;
  let lastModified = "Mon, 01 Jan 2024 00:00:00 GMT";
  let getCount = 0;
  /** @type {number[]} */
  const headStatuses = [];
  /** @type {Array<Record<string, string> | undefined>} */
  const headHeaders = [];

  /**
   * @param {{
   *   status?: number;
   *   headers?: Record<string, string>;
   *   body?: string;
   * }} init
   */
  function makeResponse(init = {}) {
    const status = init.status ?? 200;
    const headers = new Map(Object.entries(init.headers ?? {}).map(([k, v]) => [k.toLowerCase(), v]));
    const body = init.body ?? "";
    return {
      ok: status >= 200 && status < 300,
      status,
      headers: {
        get(name) {
          return headers.get(String(name).toLowerCase()) ?? null;
        },
      },
      async json() {
        return JSON.parse(body);
      },
    };
  }

  /** @type {typeof fetch} */
  const fetchMock = async (_url, init) => {
    const method = String(init?.method ?? "GET").toUpperCase();
    if (method === "HEAD") {
      // @ts-ignore - fetch init headers are passed as a plain object in tests.
      headHeaders.push(init?.headers);
      const ifNoneMatch = Object.entries(/** @type {any} */ (init?.headers ?? {})).find(([name]) => name.toLowerCase() === "if-none-match")?.[1];
      if (ifNoneMatch === etag) {
        headStatuses.push(304);
        return makeResponse({ status: 304 });
      }
      headStatuses.push(200);
      // Simulate a service that supports conditional requests but omits ETag from
      // HEAD responses. The engine should still cache-validate via If-None-Match
      // using the ETag captured from GET responses.
      return makeResponse({ headers: { "last-modified": lastModified } });
    }
    getCount += 1;
    return makeResponse({
      headers: { etag },
      body: JSON.stringify({ value: [{ Id: getCount }] }),
    });
  };

  const engine = new QueryEngine({
    cache,
    defaultCacheTtlMs: 10_000,
    connectors: { odata: new ODataConnector({ fetch: fetchMock }) },
  });

  const query = {
    id: "q_odata_cache",
    name: "OData cache",
    source: { type: "odata", url: "https://example.com/odata/Products" },
    steps: [{ id: "s1", name: "Take", operation: { type: "take", count: 1 } }],
    refreshPolicy: { type: "manual" },
  };

  const first = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(first.meta.cache?.hit, false);
  assert.equal(getCount, 1);
  assert.deepEqual(headStatuses, [200]);
  assert.equal(first.meta.sources[0]?.etag, initialEtag);

  const second = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(second.meta.cache?.hit, true);
  assert.equal(getCount, 1, "cache hit should not refetch the resource");
  assert.deepEqual(headStatuses, [200, 304]);
  assert.equal(
    Object.entries(headHeaders[1] ?? {}).find(([name]) => name.toLowerCase() === "if-none-match")?.[1],
    initialEtag,
  );

  now = 1;
  etag = '"v2"';
  lastModified = "Tue, 02 Jan 2024 00:00:00 GMT";

  const third = await engine.executeQueryWithMeta(query, {}, {});
  assert.equal(third.meta.cache?.hit, false);
  assert.equal(getCount, 2, "changed source state should invalidate the cache entry");
  assert.deepEqual(headStatuses, [200, 304, 200, 200]);
  assert.equal(
    Object.entries(headHeaders[2] ?? {}).find(([name]) => name.toLowerCase() === "if-none-match")?.[1],
    initialEtag,
  );
});
