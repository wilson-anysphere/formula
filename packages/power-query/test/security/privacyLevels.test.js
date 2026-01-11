import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { SqlConnector } from "../../src/connectors/sql.js";
import { QueryEngine } from "../../src/engine.js";
import { QueryFoldingEngine } from "../../src/folding/sql.js";
import { getFileSourceId, getHttpSourceId, getSqlSourceId } from "../../src/privacy/sourceId.js";
import { DataTable } from "../../src/table.js";

test("privacy levels: merge folding allowed when both sources are same sourceId + privacy level", async () => {
  /** @type {{ sql: string }[]} */
  const calls = [];

  const engine = new QueryEngine({
    privacyMode: "enforce",
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql) => {
          calls.push({ sql });
          return DataTable.fromGrid(
            [
              ["Id", "Region", "Target"],
              [1, "East", 10],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const connectionA = { id: "db1" };
  const connectionB = { id: "db1" }; // distinct object; should still be treated as the same sourceId.
  const sourceId = getSqlSourceId(connectionA);

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection: connectionB, query: "SELECT * FROM targets" },
    steps: [{ id: "r1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    source: { type: "database", connection: connectionA, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      { id: "l1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Region"] } },
      { id: "l2", name: "Merge", operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKey: "Id", rightKey: "Id" } },
    ],
  };

  const result = await engine.executeQuery(
    left,
    { queries: { q_right: right }, privacy: { levelsBySourceId: { [sourceId]: "organizational" } } },
    { onProgress: () => {} },
  );

  assert.equal(calls.length, 1, "expected merge to fold into a single SQL query");
  assert.match(calls[0].sql, /\bJOIN\b/);
  assert.deepEqual(result.toGrid(), [
    ["Id", "Region", "Target"],
    [1, "East", 10],
  ]);
});

test("privacy levels: append folding allowed when both sources are same sourceId + privacy level", async () => {
  /** @type {{ sql: string }[]} */
  const calls = [];

  const engine = new QueryEngine({
    privacyMode: "enforce",
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql) => {
          calls.push({ sql });
          return DataTable.fromGrid(
            [
              ["Id", "Value"],
              [1, "a"],
              [2, "b"],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const baseConn = { id: "db1" };
  const otherConn = { id: "db1" };
  const sourceId = getSqlSourceId(baseConn);

  const other = {
    id: "q_other",
    name: "Other",
    source: { type: "database", connection: otherConn, query: "SELECT * FROM b" },
    steps: [{ id: "o1", name: "Select", operation: { type: "selectColumns", columns: ["Value", "Id"] } }],
  };

  const base = {
    id: "q_base",
    name: "Base",
    source: { type: "database", connection: baseConn, query: "SELECT * FROM a", dialect: "postgres" },
    steps: [
      { id: "b1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Value"] } },
      { id: "b2", name: "Append", operation: { type: "append", queries: ["q_other"] } },
    ],
  };

  await engine.executeQuery(base, { queries: { q_other: other }, privacy: { levelsBySourceId: { [sourceId]: "organizational" } } }, {});

  assert.equal(calls.length, 1, "expected append to fold into a single SQL query");
  assert.match(calls[0].sql, /\bUNION ALL\b/);
});

test("privacy levels: merge folding prevented when privacy levels differ (hybrid + diagnostic)", async () => {
  const folding = new QueryFoldingEngine();

  const leftConn = { id: "db1" };
  const rightConn = { id: "db2" };
  const leftSourceId = getSqlSourceId(leftConn);
  const rightSourceId = getSqlSourceId(rightConn);

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection: rightConn, query: "SELECT * FROM targets" },
    steps: [{ id: "r1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    source: { type: "database", connection: leftConn, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      { id: "l1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Region"] } },
      { id: "l2", name: "Merge", operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKey: "Id", rightKey: "Id" } },
    ],
  };

  const plan = folding.compile(left, {
    queries: { q_right: right },
    privacyMode: "enforce",
    privacyLevelsBySourceId: { [leftSourceId]: "organizational", [rightSourceId]: "public" },
  });
  assert.equal(plan.type, "hybrid");
  assert.ok(Array.isArray(plan.diagnostics) && plan.diagnostics.length > 0);

  /** @type {{ sql: string }[]} */
  const calls = [];
  /** @type {any[]} */
  const events = [];

  const engine = new QueryEngine({
    privacyMode: "enforce",
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql) => {
          calls.push({ sql });
          return DataTable.fromGrid(
            [
              ["Id", "Region", "Target"],
              [1, "East", 10],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  await engine.executeQuery(
    left,
    { queries: { q_right: right }, privacy: { levelsBySourceId: { [leftSourceId]: "organizational", [rightSourceId]: "public" } } },
    { onProgress: (e) => events.push(e) },
  );

  assert.ok(calls.length >= 2, "expected separate SQL round-trips when merge folding is prevented");
  assert.ok(!calls[0].sql.includes("JOIN"), "expected the folded SQL prefix to not include a join");

  const firewallEvents = events.filter((e) => e.type === "privacy:firewall" && e.phase === "folding" && e.operation === "merge");
  assert.ok(firewallEvents.length > 0, "expected a folding firewall diagnostic event");
  assert.deepEqual(
    firewallEvents[0].sources.map((s) => ({ sourceId: s.sourceId, level: s.level })),
    [
      { sourceId: leftSourceId, level: "organizational" },
      { sourceId: rightSourceId, level: "public" },
    ],
  );
});

test("privacy levels: strict mode blocks execution when combining Private + Public", async () => {
  /** @type {any[]} */
  const events = [];

  const csvPath = "/tmp/private.csv";
  const apiUrl = "https://public.example.com/data";

  const csvSourceId = getFileSourceId(csvPath);
  const apiSourceId = getHttpSourceId(apiUrl);

  const engine = new QueryEngine({
    privacyMode: "enforce",
    fileAdapter: {
      readText: async () => ["Id,Region", "1,East"].join("\n"),
    },
    apiAdapter: {
      fetchTable: async () =>
        DataTable.fromGrid(
          [
            ["Id", "Target"],
            [1, 10],
          ],
          { hasHeaders: true, inferTypes: true },
        ),
    },
  });

  const publicQuery = {
    id: "q_public",
    name: "Public API",
    source: { type: "api", url: apiUrl, method: "GET" },
    steps: [],
  };

  const privateQuery = {
    id: "q_private",
    name: "Private CSV",
    source: { type: "csv", path: csvPath },
    steps: [
      {
        id: "s_merge",
        name: "Merge",
        operation: { type: "merge", rightQuery: "q_public", joinType: "left", leftKey: "Id", rightKey: "Id" },
      },
    ],
  };

  await assert.rejects(
    () =>
      engine.executeQuery(
        privateQuery,
        {
          queries: { q_public: publicQuery },
          privacy: { levelsBySourceId: { [csvSourceId]: "private", [apiSourceId]: "public" } },
        },
        { onProgress: (e) => events.push(e) },
      ),
    /Formula\.Firewall/,
  );

  const blocks = events.filter((e) => e.type === "privacy:firewall" && e.action === "block");
  assert.ok(blocks.length > 0, "expected a blocking firewall diagnostic event");
});

test("privacy levels: strict mode blocks when combining Private + Public via precomputed queryResults", async () => {
  /** @type {any[]} */
  const events = [];

  const csvPath = "/tmp/private.csv";
  const apiUrl = "https://public.example.com/data";

  const csvSourceId = getFileSourceId(csvPath);
  const apiSourceId = getHttpSourceId(apiUrl);

  const engine = new QueryEngine({
    privacyMode: "enforce",
    fileAdapter: {
      readText: async () => ["Id,Region", "1,East"].join("\n"),
    },
  });

  const publicTable = DataTable.fromGrid(
    [
      ["Id", "Target"],
      [1, 10],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const now = new Date(0);
  const publicMeta = {
    queryId: "q_public",
    startedAt: now,
    completedAt: now,
    refreshedAt: now,
    sources: [
      {
        refreshedAt: now,
        schema: { columns: publicTable.columns, inferred: true },
        rowCount: publicTable.rowCount,
        rowCountEstimate: publicTable.rowCount,
        provenance: { kind: "http", url: apiUrl, method: "GET" },
      },
    ],
    outputSchema: { columns: publicTable.columns, inferred: true },
    outputRowCount: publicTable.rowCount,
  };

  const privateQuery = {
    id: "q_private",
    name: "Private CSV",
    source: { type: "csv", path: csvPath },
    steps: [
      {
        id: "s_merge",
        name: "Merge",
        operation: { type: "merge", rightQuery: "q_public", joinType: "left", leftKey: "Id", rightKey: "Id" },
      },
    ],
  };

  await assert.rejects(
    () =>
      engine.executeQuery(
        privateQuery,
        {
          queries: {},
          queryResults: { q_public: { table: publicTable, meta: /** @type {any} */ (publicMeta) } },
          privacy: { levelsBySourceId: { [csvSourceId]: "private", [apiSourceId]: "public" } },
        },
        { onProgress: (e) => events.push(e) },
      ),
    /Formula\.Firewall/,
  );

  const blocks = events.filter((e) => e.type === "privacy:firewall" && e.action === "block");
  assert.ok(blocks.length > 0, "expected a blocking firewall diagnostic event");
});

test("privacy levels: sql sourceId respects SqlConnector.getConnectionIdentity", async () => {
  /** @type {any[]} */
  const events = [];

  const sqlConnector = new SqlConnector({
    querySql: async (_connection, _sql) =>
      DataTable.fromGrid(
        [
          ["Id", "Region", "Target"],
          [1, "East", 10],
        ],
        { hasHeaders: true, inferTypes: true },
      ),
    getConnectionIdentity: (connection) => {
      const c = /** @type {any} */ (connection);
      return { server: c.server, database: c.database };
    },
  });

  const engine = new QueryEngine({
    privacyMode: "enforce",
    connectors: { sql: sqlConnector },
  });

  const leftConn = { id: "db1", server: "s1", database: "a" };
  const rightConn = { id: "db2", server: "s2", database: "b" };

  const leftConnectionId = /** @type {any} */ (sqlConnector.getCacheKey({ connection: leftConn, sql: "SELECT 1" })).connectionId;
  const rightConnectionId = /** @type {any} */ (sqlConnector.getCacheKey({ connection: rightConn, sql: "SELECT 1" })).connectionId;
  assert.ok(typeof leftConnectionId === "string" && leftConnectionId.length > 0);
  assert.ok(typeof rightConnectionId === "string" && rightConnectionId.length > 0);

  const leftSourceId = getSqlSourceId(leftConnectionId);
  const rightSourceId = getSqlSourceId(rightConnectionId);

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection: rightConn, query: "SELECT * FROM targets" },
    steps: [{ id: "r1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    source: { type: "database", connection: leftConn, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      { id: "l1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Region"] } },
      { id: "l2", name: "Merge", operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKey: "Id", rightKey: "Id" } },
    ],
  };

  await assert.rejects(
    () =>
      engine.executeQuery(
        left,
        { queries: { q_right: right }, privacy: { levelsBySourceId: { [leftSourceId]: "private", [rightSourceId]: "public" } } },
        { onProgress: (e) => events.push(e) },
      ),
    /Formula\.Firewall/,
  );

  const blocks = events.filter((e) => e.type === "privacy:firewall" && e.phase === "combine" && e.action === "block");
  assert.ok(blocks.length > 0, "expected a blocking firewall diagnostic event");
  assert.deepEqual(
    blocks[0].sources.map((s) => s.sourceId),
    [leftSourceId, rightSourceId].sort(),
  );
});

test("privacy levels: warn mode allows execution but emits diagnostic", async () => {
  /** @type {any[]} */
  const events = [];

  const csvPath = "/tmp/private.csv";
  const apiUrl = "https://public.example.com/data";

  const csvSourceId = getFileSourceId(csvPath);
  const apiSourceId = getHttpSourceId(apiUrl);

  const engine = new QueryEngine({
    privacyMode: "warn",
    fileAdapter: {
      readText: async () => ["Id,Region", "1,East", "2,West"].join("\n"),
    },
    apiAdapter: {
      fetchTable: async () =>
        DataTable.fromGrid(
          [
            ["Id", "Target"],
            [1, 10],
            [3, 30],
          ],
          { hasHeaders: true, inferTypes: true },
        ),
    },
  });

  const publicQuery = {
    id: "q_public",
    name: "Public API",
    source: { type: "api", url: apiUrl, method: "GET" },
    steps: [],
  };

  const privateQuery = {
    id: "q_private",
    name: "Private CSV",
    source: { type: "csv", path: csvPath },
    steps: [
      {
        id: "s_merge",
        name: "Merge",
        operation: { type: "merge", rightQuery: "q_public", joinType: "left", leftKey: "Id", rightKey: "Id" },
      },
    ],
  };

  const result = await engine.executeQuery(
    privateQuery,
    {
      queries: { q_public: publicQuery },
      privacy: { levelsBySourceId: { [csvSourceId]: "private", [apiSourceId]: "public" } },
    },
    { onProgress: (e) => events.push(e) },
  );

  assert.deepEqual(result.toGrid(), [
    ["Id", "Region", "Target"],
    [1, "East", 10],
    [2, "West", null],
  ]);

  const warnings = events.filter((e) => e.type === "privacy:firewall" && e.action === "warn");
  assert.ok(warnings.length > 0, "expected a warning firewall diagnostic event");
});

test("privacy levels: cache keys incorporate privacy levels + mode", async () => {
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => 0 });

  const csvPath = "/tmp/sales.csv";
  const csvSourceId = getFileSourceId(csvPath);

  const query = {
    id: "q_sales",
    name: "Sales",
    source: { type: "csv", path: csvPath },
    steps: [],
  };

  const engineIgnore = new QueryEngine({ privacyMode: "ignore", cache, fileAdapter: { readText: async () => "" } });
  const engineEnforce = new QueryEngine({ privacyMode: "enforce", cache, fileAdapter: { readText: async () => "" } });

  const keyPublic = await engineIgnore.getCacheKey(query, { privacy: { levelsBySourceId: { [csvSourceId]: "public" } } }, {});
  const keyPrivate = await engineIgnore.getCacheKey(query, { privacy: { levelsBySourceId: { [csvSourceId]: "private" } } }, {});
  const keyEnforce = await engineEnforce.getCacheKey(query, { privacy: { levelsBySourceId: { [csvSourceId]: "public" } } }, {});

  assert.ok(keyPublic && keyPrivate && keyEnforce);
  assert.notEqual(keyPublic, keyPrivate, "privacy level changes should change the cache key");
  assert.notEqual(keyPublic, keyEnforce, "privacy mode changes should change the cache key");
});

test("privacy source ids: http origins normalize IPv6 hosts with brackets", () => {
  assert.equal(getHttpSourceId("http://[::1]/data"), "http://[::1]:80");
  assert.equal(getHttpSourceId("https://[::1]/data"), "https://[::1]:443");
});

test("privacy source ids: sql source ids are not double-prefixed", () => {
  assert.equal(getSqlSourceId("sql:db1"), "sql:db1");
  assert.equal(getSqlSourceId({ id: "sql:db1" }), "sql:db1");
});

test("privacy source ids: file source ids preserve UNC prefix", () => {
  assert.equal(getFileSourceId("\\\\server\\share\\dir\\..\\file.csv"), "//server/share/file.csv");
});
