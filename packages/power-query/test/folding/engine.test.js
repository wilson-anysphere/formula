import assert from "node:assert/strict";
import test from "node:test";

import { CacheManager } from "../../src/cache/cache.js";
import { MemoryCacheStore } from "../../src/cache/memory.js";
import { QueryEngine } from "../../src/engine.js";
import { DataTable } from "../../src/table.js";
import { SqlConnector } from "../../src/connectors/sql.js";

test("QueryEngine: executes folded SQL when database dialect is provided", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined } | null} */
  let observed = null;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          observed = { sql, params: options?.params };
          // Simulate database applying the filter by returning only matching rows.
          return DataTable.fromGrid(
            [
              ["Region", "Sales"],
              ["East", 100],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const query = {
    id: "q_fold",
    name: "Fold",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
    ],
  };

  const result = await engine.executeQuery(query, { queries: {} }, {});
  assert.ok(observed, "expected SQL connector to be invoked");
  assert.match(observed.sql, /WHERE/);
  assert.deepEqual(observed.params, ["East"]);
  assert.deepEqual(result.toGrid(), [
    ["Region", "Sales"],
    ["East", 100],
  ]);
});

test("QueryEngine: pushes ExecuteOptions.limit down when the plan fully folds to SQL", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined } | null} */
  let observed = null;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          observed = { sql, params: options?.params };
          // Pretend the database applied both the filter and limit.
          return DataTable.fromGrid(
            [
              ["Region", "Sales"],
              ["East", 100],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const query = {
    id: "q_fold_limit",
    name: "Fold + Limit",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
    ],
  };

  await engine.executeQuery(query, { queries: {} }, { limit: 10 });
  assert.ok(observed, "expected SQL connector to be invoked");
  assert.match(observed.sql, /\bLIMIT\b/);
  assert.deepEqual(observed.params, ["East", 10]);
});

test("QueryEngine: without a dialect, executes steps locally (no folding)", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined } | null} */
  let observed = null;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          observed = { sql, params: options?.params };
          // Return unfiltered rows; local engine should filter them.
          return DataTable.fromGrid(
            [
              ["Region", "Sales"],
              ["East", 100],
              ["West", 200],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const query = {
    id: "q_local",
    name: "Local",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM sales" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
    ],
  };

  const result = await engine.executeQuery(query, { queries: {} }, {});
  assert.ok(observed, "expected SQL connector to be invoked");
  assert.equal(observed.sql, "SELECT * FROM sales");
  assert.equal(observed.params, undefined);
  assert.deepEqual(result.toGrid(), [
    ["Region", "Sales"],
    ["East", 100],
  ]);
});

test("QueryEngine: executes hybrid folded SQL then runs remaining steps locally", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined } | null} */
  let observed = null;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          observed = { sql, params: options?.params };
          return DataTable.fromGrid(
            [
              ["Region", "Sales"],
              ["East", 100],
              [null, 150],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const query = {
    id: "q_hybrid",
    name: "Hybrid",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Sales", operator: "greaterThan", value: 0 } },
      },
      {
        id: "s2",
        name: "Fill Down",
        operation: { type: "fillDown", columns: ["Region"] },
      },
    ],
  };

  const { table: result, meta } = await engine.executeQueryWithMeta(query, { queries: {} }, {});
  assert.ok(observed, "expected SQL connector to be invoked");
  assert.match(observed.sql, /WHERE/);
  assert.deepEqual(observed.params, [0]);
  assert.ok(meta.folding, "expected folding metadata");
  assert.equal(meta.folding.planType, "hybrid");
  assert.equal(meta.folding.dialect, "postgres");
  assert.deepEqual(
    meta.folding.steps.map((s) => s.status),
    ["folded", "local"],
  );
  assert.equal(meta.folding.steps[1].reason, "unsupported_op");
  assert.equal(meta.folding.localStepOffset, 1);

  // fillDown should run locally after the SQL query.
  assert.deepEqual(result.toGrid(), [
    ["Region", "Sales"],
    ["East", 100],
    ["East", 150],
  ]);
});

test("QueryEngine: folds merge into a single SQL query when both sides are foldable", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined }[]} */
  const calls = [];
  const connection = {};

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          calls.push({ sql, params: options?.params });
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

  const right = {
    id: "q_right",
    name: "Targets",
    source: { type: "database", connection, query: "SELECT * FROM targets" },
    steps: [{ id: "r1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Target"] } }],
  };

  const left = {
    id: "q_left",
    name: "Sales",
    source: { type: "database", connection, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      { id: "l1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Region"] } },
      { id: "l2", name: "Merge", operation: { type: "merge", rightQuery: "q_right", joinType: "left", leftKey: "Id", rightKey: "Id" } },
    ],
  };

  const result = await engine.executeQuery(left, { queries: { q_right: right } }, {});
  assert.equal(calls.length, 1, "expected a single SQL roundtrip when merge folds");
  assert.match(calls[0].sql, /\bJOIN\b/);
  assert.deepEqual(result.toGrid(), [
    ["Id", "Region", "Target"],
    [1, "East", 10],
  ]);
});

test("QueryEngine: folds append into a single SQL query when schemas are compatible", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined }[]} */
  const calls = [];
  const connection = {};

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async (_connection, sql, options) => {
          calls.push({ sql, params: options?.params });
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

  const other = {
    id: "q_other",
    name: "Other",
    source: { type: "database", connection, query: "SELECT * FROM b" },
    steps: [{ id: "o1", name: "Select", operation: { type: "selectColumns", columns: ["Value", "Id"] } }],
  };

  const base = {
    id: "q_base",
    name: "Base",
    source: { type: "database", connection, query: "SELECT * FROM a", dialect: "postgres" },
    steps: [
      { id: "b1", name: "Select", operation: { type: "selectColumns", columns: ["Id", "Value"] } },
      { id: "b2", name: "Append", operation: { type: "append", queries: ["q_other"] } },
    ],
  };

  const result = await engine.executeQuery(base, { queries: { q_other: other } }, {});
  assert.equal(calls.length, 1, "expected a single SQL roundtrip when append folds");
  assert.match(calls[0].sql, /\bUNION ALL\b/);
  assert.deepEqual(result.toGrid(), [
    ["Id", "Value"],
    [1, "a"],
    [2, "b"],
  ]);
});

test("QueryEngine: schema discovery enables folding renameColumn/changeType without an explicit projection", async () => {
  /** @type {{ sql: string, params: unknown[] | undefined }[]} */
  const calls = [];
  let schemaCalls = 0;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        getSchema: async () => {
          schemaCalls += 1;
          return { columns: ["Region", "Sales"], types: { Region: "string", Sales: "number" } };
        },
        querySql: async (_connection, sql, options) => {
          calls.push({ sql, params: options?.params });
          // Return data as though the database applied the rename + cast.
          return DataTable.fromGrid(
            [
              ["Region", "Amount"],
              ["East", 100],
            ],
            { hasHeaders: true, inferTypes: true },
          );
        },
      }),
    },
  });

  const query = {
    id: "q_schema_folding",
    name: "Schema Folding",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM raw", dialect: "postgres" },
    steps: [
      { id: "s1", name: "Rename", operation: { type: "renameColumn", oldName: "Sales", newName: "Amount" } },
      { id: "s2", name: "Type", operation: { type: "changeType", column: "Amount", newType: "number" } },
    ],
  };

  const first = await engine.executeQuery(query, { queries: {} }, {});
  assert.equal(schemaCalls, 1);
  assert.equal(query.source.columns, undefined, "schema discovery should not mutate the caller query");
  assert.equal(calls.length, 1);
  assert.match(calls[0].sql, /\bCAST\b/);
  assert.deepEqual(first.toGrid(), [
    ["Region", "Amount"],
    ["East", 100],
  ]);

  // Schema discovery should be cached within the engine instance.
  await engine.executeQuery(query, { queries: {} }, {});
  assert.equal(schemaCalls, 1);
});

test("QueryEngine: schema discovery cache varies by credentialId", async () => {
  let schemaCalls = 0;
  let currentCredentialId = "cred-a";

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        getSchema: async () => {
          schemaCalls += 1;
          return { columns: ["Region", "Sales"], types: { Region: "string", Sales: "number" } };
        },
        querySql: async () =>
          DataTable.fromGrid(
            [
              ["Region", "Amount"],
              ["East", 100],
            ],
            { hasHeaders: true, inferTypes: true },
          ),
      }),
    },
    onPermissionRequest: async () => true,
    onCredentialRequest: async () => ({ credentialId: currentCredentialId }),
  });

  const query = {
    id: "q_schema_cache_by_cred",
    name: "Schema Cache by Credential",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM raw", dialect: "postgres" },
    steps: [
      { id: "s1", name: "Rename", operation: { type: "renameColumn", oldName: "Sales", newName: "Amount" } },
      { id: "s2", name: "Type", operation: { type: "changeType", column: "Amount", newType: "number" } },
    ],
  };

  await engine.executeQuery(query, { queries: {} }, {});
  assert.equal(schemaCalls, 1);

  // Same credentials -> should reuse schema cache entry.
  await engine.executeQuery(query, { queries: {} }, {});
  assert.equal(schemaCalls, 1);

  // Different credentials -> should not reuse cached schema.
  currentCredentialId = "cred-b";
  await engine.executeQuery(query, { queries: {} }, {});
  assert.equal(schemaCalls, 2);
});

test("QueryEngine: database credentials are requested once per execution when cache key generation runs", async () => {
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => 0 });

  let permissionCount = 0;
  let credentialCount = 0;

  const engine = new QueryEngine({
    cache,
    connectors: {
      sql: new SqlConnector({
        querySql: async () =>
          DataTable.fromGrid(
            [
              ["Value"],
              [1],
            ],
            { hasHeaders: true, inferTypes: true },
          ),
      }),
    },
    onPermissionRequest: async () => {
      permissionCount += 1;
      return true;
    },
    onCredentialRequest: async () => {
      credentialCount += 1;
      return { credentialId: "cred-1" };
    },
  });

  const query = {
    id: "q_db_credential_cache",
    name: "DB Credential Cache",
    // No dialect: no folding. We want to validate that cache key generation (which requests credentials)
    // shares the same credential result with the subsequent source execution in the same run.
    source: { type: "database", connection: { id: "db1" }, query: "SELECT 1" },
    steps: [],
  };

  await engine.executeQueryWithMeta(query, { queries: {} }, {});
  assert.equal(permissionCount, 1);
  assert.equal(credentialCount, 1);
});

test("QueryEngine: SqlConnector.getConnectionIdentity can override connection.id for cache keys", async () => {
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => 0 });

  const engine = new QueryEngine({
    cache,
    connectors: {
      sql: new SqlConnector({
        getConnectionIdentity: (connection) => {
          // Deliberately incorporate more than `connection.id` so different hosts
          // with the same logical id produce distinct cache keys.
          if (!connection || typeof connection !== "object") return null;
          // @ts-ignore - runtime indexing
          return { id: connection.id, host: connection.host };
        },
        querySql: async () =>
          DataTable.fromGrid(
            [
              ["Value"],
              [1],
            ],
            { hasHeaders: true, inferTypes: true },
          ),
      }),
    },
  });

  const base = {
    id: "q_db_identity_override",
    name: "DB Identity Override",
    source: { type: "database", connection: { id: "db1", host: "a" }, query: "SELECT 1" },
    steps: [],
  };
  const otherHostSameId = {
    ...base,
    source: { ...base.source, connection: { id: "db1", host: "b" } },
  };

  const keyA = await engine.getCacheKey(base, { queries: {} }, {});
  const keyB = await engine.getCacheKey(otherHostSameId, { queries: {} }, {});
  assert.ok(keyA);
  assert.ok(keyB);
  assert.notEqual(keyA, keyB);
});

test("QueryEngine: database permission/credentials are requested once per execution when folding runs", async () => {
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => 0 });

  let permissionCount = 0;
  let credentialCount = 0;

  const engine = new QueryEngine({
    cache,
    connectors: {
      sql: new SqlConnector({
        querySql: async () =>
          DataTable.fromGrid(
            [
              ["Region", "Sales"],
              ["East", 100],
            ],
            { hasHeaders: true, inferTypes: true },
          ),
      }),
    },
    onPermissionRequest: async () => {
      permissionCount += 1;
      return true;
    },
    onCredentialRequest: async () => {
      credentialCount += 1;
      return { credentialId: "cred-1" };
    },
  });

  const query = {
    id: "q_db_fold_credential_cache",
    name: "DB Fold Credential Cache",
    source: { type: "database", connection: { id: "db1" }, query: "SELECT * FROM sales", dialect: "postgres" },
    steps: [
      {
        id: "s1",
        name: "Filter",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
    ],
  };

  await engine.executeQueryWithMeta(query, { queries: {} }, {});
  assert.equal(permissionCount, 1);
  assert.equal(credentialCount, 1);
});

test("QueryEngine: permission/credential caches do not collide for DB sources without connection identity", async () => {
  let permissionCount = 0;
  let credentialCount = 0;

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async () =>
          DataTable.fromGrid(
            [
              ["Value"],
              [1],
            ],
            { hasHeaders: true, inferTypes: true },
          ),
      }),
    },
    onPermissionRequest: async () => {
      permissionCount += 1;
      return true;
    },
    onCredentialRequest: async () => {
      credentialCount += 1;
      return { credentialId: `cred-${credentialCount}` };
    },
  });

  const session = engine.createSession({ now: () => 0 });

  const q1 = { id: "q_db_no_id_1", name: "DB no identity 1", source: { type: "database", connection: {}, query: "SELECT 1" }, steps: [] };
  const q2 = { id: "q_db_no_id_2", name: "DB no identity 2", source: { type: "database", connection: {}, query: "SELECT 1" }, steps: [] };

  await engine.executeQueryWithMetaInSession(q1, { queries: {} }, {}, session);
  await engine.executeQueryWithMetaInSession(q2, { queries: {} }, {}, session);

  assert.equal(permissionCount, 2);
  assert.equal(credentialCount, 2);
});

test("QueryEngine: permission/credential caches reuse prompts for the same DB connection handle without identity", async () => {
  let permissionCount = 0;
  let credentialCount = 0;
  const connection = {};

  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async () =>
          DataTable.fromGrid(
            [
              ["Value"],
              [1],
            ],
            { hasHeaders: true, inferTypes: true },
          ),
      }),
    },
    onPermissionRequest: async () => {
      permissionCount += 1;
      return true;
    },
    onCredentialRequest: async () => {
      credentialCount += 1;
      return { credentialId: `cred-${credentialCount}` };
    },
  });

  const session = engine.createSession({ now: () => 0 });

  const q1 = {
    id: "q_db_no_id_same_1",
    name: "DB no identity same 1",
    source: { type: "database", connection, query: "SELECT 1" },
    steps: [],
  };
  const q2 = {
    id: "q_db_no_id_same_2",
    name: "DB no identity same 2",
    source: { type: "database", connection, query: "SELECT 1" },
    steps: [],
  };

  await engine.executeQueryWithMetaInSession(q1, { queries: {} }, {}, session);
  await engine.executeQueryWithMetaInSession(q2, { queries: {} }, {}, session);

  assert.equal(permissionCount, 1);
  assert.equal(credentialCount, 1);
});

test("QueryEngine: table source ids do not collide for DB sources without identity", async () => {
  const engine = new QueryEngine({
    connectors: {
      sql: new SqlConnector({
        querySql: async () =>
          DataTable.fromGrid(
            [
              ["Value"],
              [1],
            ],
            { hasHeaders: true, inferTypes: true },
          ),
      }),
    },
  });

  const q1 = { id: "q_db_source_ids_1", name: "DB source ids 1", source: { type: "database", connection: {}, query: "SELECT 1" }, steps: [] };
  const q2 = { id: "q_db_source_ids_2", name: "DB source ids 2", source: { type: "database", connection: {}, query: "SELECT 1" }, steps: [] };

  const r1 = await engine.executeQueryWithMeta(q1, { queries: {} }, { cache: { mode: "bypass" } });
  const r2 = await engine.executeQueryWithMeta(q2, { queries: {} }, { cache: { mode: "bypass" } });

  // @ts-ignore - testing internal bookkeeping
  const ids1 = Array.from(engine.getTableSourceIds(r1.table));
  // @ts-ignore - testing internal bookkeeping
  const ids2 = Array.from(engine.getTableSourceIds(r2.table));

  assert.equal(ids1.length, 1);
  assert.equal(ids2.length, 1);
  assert.notEqual(ids1[0], ids2[0]);
});

test("QueryEngine: source-state cache validation respects explicit database connectionId", async () => {
  const store = new MemoryCacheStore();
  const cache = new CacheManager({ store, now: () => 0 });

  let queryCalls = 0;
  const sql = new SqlConnector({
    querySql: async () => {
      queryCalls += 1;
      return DataTable.fromGrid(
        [
          ["Value"],
          [1],
        ],
        { hasHeaders: true, inferTypes: true },
      );
    },
  });

  // Pretend the host can validate DB state (e.g. via a schema version query).
  // This forces the engine to use `collectSourceStateTargetsFromSource`.
  // @ts-ignore - runtime augmentation
  sql.getSourceState = async () => ({ etag: "v1" });

  const engine = new QueryEngine({ cache, connectors: { sql } });

  const query = {
    id: "q_db_state_validation",
    name: "DB state validation",
    source: {
      type: "database",
      // Explicit connectionId differs from what the connector could derive from `connection`.
      connectionId: "explicit-connection",
      connection: { id: "db1" },
      query: "SELECT 1",
    },
    steps: [],
  };

  const first = await engine.executeQueryWithMeta(query, { queries: {} }, {});
  assert.equal(first.meta.cache?.hit, false);
  assert.equal(queryCalls, 1);

  const second = await engine.executeQueryWithMeta(query, { queries: {} }, {});
  assert.equal(second.meta.cache?.hit, true);
  assert.equal(queryCalls, 1, "cache hit should not re-execute the DB query");
});
