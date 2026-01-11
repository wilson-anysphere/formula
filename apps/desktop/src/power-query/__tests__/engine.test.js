import assert from "node:assert/strict";
import test from "node:test";

import { createDesktopQueryEngine } from "../engine.ts";
import { createDefaultOrgPolicy } from "../../../../../packages/security/dlp/src/policy.js";
import { DLP_ACTION } from "../../../../../packages/security/dlp/src/actions.js";
import { getHttpSourceId } from "../../../../../packages/power-query/src/privacy/sourceId.js";
import { DataTable } from "../../../../../packages/power-query/src/table.js";

test("createDesktopQueryEngine uses Tauri invoke file commands when FS plugin is unavailable", async () => {
  const originalTauri = globalThis.__TAURI__;

  /** @type {{ cmd: string, args: any }[]} */
  const calls = [];

  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd, args) => {
        calls.push({ cmd, args });
        if (cmd === "stat_file") {
          return { mtimeMs: 123 };
        }
        if (cmd === "read_text_file") {
          return ["A,B", "1,2"].join("\n");
        }
        if (cmd === "read_binary_file") {
          // Base64 for bytes [1, 2, 3].
          return "AQID";
        }
        throw new Error(`Unexpected invoke: ${cmd}`);
      },
    },
  };

  try {
    const engine = createDesktopQueryEngine();

    const query = {
      id: "q_csv",
      name: "CSV",
      source: { type: "csv", path: "/tmp/test.csv", options: { hasHeaders: true } },
      steps: [],
    };

    const table = await engine.executeQuery(query, {}, {});
    assert.deepEqual(table.toGrid(), [["A", "B"], [1, 2]]);

    const bytes = await engine.fileAdapter.readBinary("/tmp/test.bin");
    assert.deepEqual(Array.from(bytes), [1, 2, 3]);

    assert.ok(calls.some((c) => c.cmd === "read_text_file"));
    assert.ok(calls.some((c) => c.cmd === "read_binary_file"));
    assert.ok(calls.some((c) => c.cmd === "stat_file"));
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});

test("createDesktopQueryEngine exposes chunked file reads for streaming (readBinaryStream/openFile)", async () => {
  const originalTauri = globalThis.__TAURI__;

  /** @type {{ cmd: string, args: any }[]} */
  const calls = [];

  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd, args) => {
        calls.push({ cmd, args });
        if (cmd === "stat_file") {
          return { mtimeMs: 123, sizeBytes: 3 };
        }
        if (cmd === "read_binary_file_range") {
          // Return bytes [1, 2, 3] as base64, then EOF.
          const offset = Number(args?.offset ?? 0);
          if (offset === 0) return "AQID";
          if (offset === 1) return "AgM=";
          return "";
        }
        throw new Error(`Unexpected invoke: ${cmd}`);
      },
    },
  };

  try {
    const engine = createDesktopQueryEngine();

    // readBinaryStream
    const streamed = [];
    for await (const chunk of engine.fileAdapter.readBinaryStream("/tmp/test.bin")) {
      streamed.push(...chunk);
    }
    assert.deepEqual(streamed, [1, 2, 3]);
    assert.ok(calls.some((c) => c.cmd === "read_binary_file_range"));

    // openFile -> slice -> arrayBuffer
    const blob = await engine.fileAdapter.openFile("/tmp/test.bin");
    assert.equal(blob.size, 3);
    const buf = await blob.slice(1, 3).arrayBuffer();
    assert.deepEqual(Array.from(new Uint8Array(buf)), [2, 3]);
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});

test("createDesktopQueryEngine uses file mtimes to validate cache entries", async () => {
  const originalTauri = globalThis.__TAURI__;

  let mtimeMs = 1_000;
  let readCount = 0;

  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd) => {
        if (cmd === "stat_file") {
          return { mtimeMs };
        }
        if (cmd === "read_text_file") {
          readCount += 1;
          return ["A,B", "1,2"].join("\n");
        }
        throw new Error(`Unexpected invoke: ${cmd}`);
      },
    },
  };

  try {
    const engine = createDesktopQueryEngine();

    const query = {
      id: "q_csv_cache",
      name: "CSV Cache",
      source: { type: "csv", path: "/tmp/test.csv", options: { hasHeaders: true } },
      steps: [],
    };

    await engine.executeQuery(query, {}, {});
    await engine.executeQuery(query, {}, {});
    assert.equal(readCount, 1, "expected second execution to reuse cached result");

    mtimeMs = 2_000;
    await engine.executeQuery(query, {}, {});
    assert.equal(readCount, 2, "expected cache to invalidate when mtime changes");
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});

test("createDesktopQueryEngine resolves table sources via Tauri list_tables/get_range", async () => {
  const originalTauri = globalThis.__TAURI__;

  /** @type {{ cmd: string, args: any }[]} */
  const calls = [];

  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd, args) => {
        calls.push({ cmd, args });
        if (cmd === "list_tables") {
          return [
            {
              name: "Sales",
              sheet_id: "sheet1",
              start_row: 5,
              start_col: 2,
              end_row: 7,
              end_col: 3,
              columns: ["Region", "Sales"],
            },
          ];
        }
        if (cmd === "get_range") {
          assert.deepEqual(args, { sheet_id: "sheet1", start_row: 5, start_col: 2, end_row: 7, end_col: 3 });
          return {
            // Intentionally return incorrect header cells; the adapter should prefer TableInfo.columns.
            values: [
              [
                { value: "WrongRegion", formula: null, display_value: "WrongRegion" },
                { value: "WrongSales", formula: null, display_value: "WrongSales" },
              ],
              [
                { value: "East", formula: null, display_value: "East" },
                { value: 100, formula: null, display_value: "100" },
              ],
              [
                { value: "West", formula: null, display_value: "West" },
                { value: 200, formula: null, display_value: "200" },
              ],
            ],
            start_row: 5,
            start_col: 2,
          };
        }
        throw new Error(`Unexpected invoke: ${cmd}`);
      },
    },
  };

  try {
    const engine = createDesktopQueryEngine();

    const query = {
      id: "q_table",
      name: "Sales Table",
      source: { type: "table", table: "Sales" },
      steps: [],
    };

    const table = await engine.executeQuery(query, {}, {});
    assert.deepEqual(table.toGrid(), [
      ["Region", "Sales"],
      ["East", 100],
      ["West", 200],
    ]);

    assert.ok(calls.some((c) => c.cmd === "list_tables"));
    assert.ok(calls.some((c) => c.cmd === "get_range"));
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});

test("createDesktopQueryEngine enforces DLP policy on external connector permissions", async () => {
  const policy = createDefaultOrgPolicy();

  const engine = createDesktopQueryEngine({
    dlp: {
      documentId: "doc1",
      classificationStore: {
        list: () => [
          {
            selector: { scope: "document", documentId: "doc1" },
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
      policy,
    },
    fileAdapter: {
      readText: async () => ["A,B", "1,2"].join("\n"),
      readBinary: async () => new Uint8Array([1, 2, 3]),
    },
    onPermissionPrompt: () => {
      throw new Error("Permission prompt should not run when DLP blocks first");
    },
  });

  const query = {
    id: "q_csv",
    name: "CSV",
    source: { type: "csv", path: "/tmp/test.csv", options: { hasHeaders: true } },
    steps: [],
  };

  await assert.rejects(engine.executeQuery(query, {}, {}), (err) => err?.name === "DlpViolationError");
});

test("createDesktopQueryEngine supports dynamic DLP policy resolvers", async () => {
  const policy = createDefaultOrgPolicy();
  let policyCalls = 0;

  const engine = createDesktopQueryEngine({
    dlp: {
      documentId: "doc1",
      classificationStore: {
        list: () => [
          {
            selector: { scope: "document", documentId: "doc1" },
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
      policy: async () => {
        policyCalls += 1;
        return policy;
      },
    },
    fileAdapter: {
      readText: async () => ["A,B", "1,2"].join("\n"),
      readBinary: async () => new Uint8Array([1, 2, 3]),
    },
    onPermissionPrompt: () => {
      throw new Error("Permission prompt should not run when DLP blocks first");
    },
  });

  const query = {
    id: "q_csv",
    name: "CSV",
    source: { type: "csv", path: "/tmp/test.csv", options: { hasHeaders: true } },
    steps: [],
  };

  await assert.rejects(engine.executeQuery(query, {}, {}), (err) => err?.name === "DlpViolationError");
  assert.equal(policyCalls, 1);
});

test("createDesktopQueryEngine caches permission prompts across executions", async () => {
  let promptCount = 0;
  const engine = createDesktopQueryEngine({
    fileAdapter: {
      readText: async () => ["A,B", "1,2"].join("\n"),
      readBinary: async () => new Uint8Array([1, 2, 3]),
    },
    onPermissionPrompt: async () => {
      promptCount += 1;
      return true;
    },
  });

  const query = {
    id: "q_csv_perm",
    name: "CSV",
    source: { type: "csv", path: "/tmp/test.csv", options: { hasHeaders: true } },
    steps: [],
  };

  // Bypass query result caching so we can specifically validate the permission prompt cache.
  await engine.executeQuery(query, {}, { cache: { mode: "bypass" } });
  await engine.executeQuery(query, {}, { cache: { mode: "bypass" } });

  assert.equal(promptCount, 1);
});

test("createDesktopQueryEngine permission prompt cache distinguishes opaque database connections", async () => {
  let promptCount = 0;
  const engine = createDesktopQueryEngine({
    fileAdapter: {
      readText: async () => "",
      readBinary: async () => new Uint8Array(),
    },
    onPermissionPrompt: async () => {
      promptCount += 1;
      return true;
    },
  });

  const connA = {};
  const connB = {};

  /** @type {any} */
  const detailsA = {
    source: { type: "database", connection: connA, query: "select 1" },
    request: { connection: connA, sql: "select 1" },
  };
  /** @type {any} */
  const detailsB = {
    source: { type: "database", connection: connB, query: "select 1" },
    request: { connection: connB, sql: "select 1" },
  };

  await engine.onPermissionRequest("database:query", detailsA);
  await engine.onPermissionRequest("database:query", detailsA);
  await engine.onPermissionRequest("database:query", detailsB);

  assert.equal(promptCount, 2);
});

test("createDesktopQueryEngine wires oauth2Manager into HttpConnector", async () => {
  /** @type {any[]} */
  const oauthCalls = [];
  const oauth2Manager = {
    getAccessToken: async (opts) => {
      oauthCalls.push(opts);
      return { accessToken: "token", expiresAtMs: null, refreshToken: null };
    },
  };

  /** @type {typeof fetch} */
  const fetchFn = async (_url, init) => {
    const auth = /** @type {any} */ (init?.headers)?.Authorization;
    assert.equal(auth, "Bearer token");
    return new Response(JSON.stringify([{ id: 1 }]), { status: 200, headers: { "content-type": "application/json" } });
  };

  const engine = createDesktopQueryEngine({
    fetch: fetchFn,
    oauth2Manager,
    fileAdapter: { readText: async () => "", readBinary: async () => new Uint8Array() },
  });

  const query = {
    id: "q_api",
    name: "API",
    source: {
      type: "api",
      url: "https://api.example/data",
      method: "GET",
      auth: { type: "oauth2", providerId: "example" },
    },
    steps: [],
  };

  const table = await engine.executeQuery(query, {}, { cache: { validation: "none" } });
  assert.deepEqual(table.toGrid(), [["id"], [1]]);
  assert.equal(oauthCalls.length, 1);
  assert.equal(oauthCalls[0].providerId, "example");
});

test("createDesktopQueryEngine maps DLP classification into workbook privacy levels", async () => {
  const policy = createDefaultOrgPolicy();
  // Allow external connectors for Restricted documents so this test can observe the
  // privacy firewall behavior instead of a DLP block.
  policy.rules[DLP_ACTION.EXTERNAL_CONNECTOR].maxAllowed = "Restricted";

  const apiUrl = "https://public.example.com/data";
  const apiSourceId = getHttpSourceId(apiUrl);

  const engine = createDesktopQueryEngine({
    privacyMode: "enforce",
    dlp: {
      documentId: "doc1",
      classificationStore: {
        list: () => [
          {
            selector: { scope: "document", documentId: "doc1" },
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
      policy,
    },
    fileAdapter: {
      readText: async () => "",
      readBinary: async () => new Uint8Array(),
    },
    fetch: async () =>
      new Response(JSON.stringify([{ Id: 1, Target: 10 }]), { status: 200, headers: { "content-type": "application/json" } }),
  });

  const publicQuery = {
    id: "q_public",
    name: "Public API",
    source: { type: "api", url: apiUrl, method: "GET" },
    steps: [],
  };

  const privateQuery = {
    id: "q_private",
    name: "Private Range",
    source: { type: "range", range: { values: [["Id", "Region"], [1, "East"]], hasHeaders: true } },
    steps: [
      {
        id: "s_merge",
        name: "Merge",
        operation: { type: "merge", rightQuery: "q_public", joinType: "left", leftKey: "Id", rightKey: "Id" },
      },
    ],
  };

  // Only provide a privacy level for the API source; the desktop engine should
  // infer the workbook range privacy from DLP classification.
  await assert.rejects(
    engine.executeQuery(privateQuery, {
      queries: { q_public: publicQuery },
      privacy: { levelsBySourceId: { [apiSourceId]: "public" } },
    }),
    /Formula\.Firewall/,
  );
});

test("createDesktopQueryEngine does not apply workbook privacy levels to non-workbook provenance sourceIds", async () => {
  const policy = createDefaultOrgPolicy();
  // Allow external connectors for Internal documents so we can observe formula firewall behavior.
  policy.rules[DLP_ACTION.EXTERNAL_CONNECTOR].maxAllowed = "Restricted";

  const apiUrl = "https://public.example.com/data";
  const apiSourceId = getHttpSourceId(apiUrl);

  const engine = createDesktopQueryEngine({
    privacyMode: "enforce",
    dlp: {
      documentId: "doc1",
      classificationStore: {
        list: () => [
          {
            selector: { scope: "document", documentId: "doc1" },
            classification: { level: "Internal", labels: [] },
          },
        ],
      },
      policy,
    },
    fileAdapter: {
      readText: async () => "",
      readBinary: async () => new Uint8Array(),
    },
    fetch: async () =>
      new Response(JSON.stringify([{ Id: 1, Region: "East" }]), { status: 200, headers: { "content-type": "application/json" } }),
  });

  const sqlTable = DataTable.fromGrid([["Id", "Target"], [1, 10]], { hasHeaders: true, inferTypes: true });
  const sqlSourceId = "sql:db1";
  const now = new Date(0);
  const sqlMeta = {
    queryId: "q_sql",
    startedAt: now,
    completedAt: now,
    refreshedAt: now,
    sources: [
      {
        refreshedAt: now,
        schema: { columns: sqlTable.columns, inferred: true },
        rowCount: sqlTable.rowCount,
        rowCountEstimate: sqlTable.rowCount,
        // Include a sourceId here to ensure DesktopQueryEngine does not incorrectly
        // treat it as a workbook source just because it has a sourceId field.
        provenance: { kind: "sql", sourceId: sqlSourceId, connectionId: "db1", sql: "SELECT 1" },
      },
    ],
    outputSchema: { columns: sqlTable.columns, inferred: true },
    outputRowCount: sqlTable.rowCount,
  };

  const query = {
    id: "q_merge",
    name: "Merge Public API + SQL",
    source: { type: "api", url: apiUrl, method: "GET" },
    steps: [
      {
        id: "s_merge",
        name: "Merge",
        operation: { type: "merge", rightQuery: "q_sql", joinType: "left", leftKey: "Id", rightKey: "Id" },
      },
    ],
  };

  // Provide only the API privacy level; SQL should default to unknown (conservative),
  // which should trigger the formula firewall when combined with Public.
  await assert.rejects(
    engine.executeQuery(query, {
      queryResults: { q_sql: { table: sqlTable, meta: sqlMeta } },
      privacy: { levelsBySourceId: { [apiSourceId]: "public" } },
    }),
    /Formula\.Firewall/,
  );
});

test("createDesktopQueryEngine applies workbook privacy levels to host queryResults", async () => {
  const policy = createDefaultOrgPolicy();
  policy.rules[DLP_ACTION.EXTERNAL_CONNECTOR].maxAllowed = "Restricted";

  const apiUrl = "https://public.example.com/data";
  const apiSourceId = getHttpSourceId(apiUrl);

  const engine = createDesktopQueryEngine({
    privacyMode: "enforce",
    dlp: {
      documentId: "doc1",
      classificationStore: {
        list: () => [
          {
            selector: { scope: "document", documentId: "doc1" },
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
      policy,
    },
    fileAdapter: {
      readText: async () => "",
      readBinary: async () => new Uint8Array(),
    },
    fetch: async () =>
      new Response(JSON.stringify([{ Id: 1, Region: "East" }]), { status: 200, headers: { "content-type": "application/json" } }),
  });

  const privateTable = await engine.executeQuery(
    {
      id: "q_private_source",
      name: "Private Table Source",
      source: { type: "range", range: { values: [["Id", "Target"], [1, 10]], hasHeaders: true } },
      steps: [],
    },
    {},
    {},
  );

  const now = new Date(0);
  const privateMeta = {
    queryId: "q_private",
    startedAt: now,
    completedAt: now,
    refreshedAt: now,
    sources: [
      {
        refreshedAt: now,
        schema: { columns: privateTable.columns, inferred: true },
        rowCount: privateTable.rowCount,
        rowCountEstimate: privateTable.rowCount,
        provenance: { kind: "table", table: "Sales" },
      },
    ],
    outputSchema: { columns: privateTable.columns, inferred: true },
    outputRowCount: privateTable.rowCount,
  };

  const query = {
    id: "q_api",
    name: "API",
    source: { type: "api", url: apiUrl, method: "GET" },
    steps: [
      {
        id: "s_merge",
        name: "Merge",
        operation: { type: "merge", rightQuery: "q_private", joinType: "left", leftKey: "Id", rightKey: "Id" },
      },
    ],
  };

  await assert.rejects(
    engine.executeQuery(
      query,
      {
        // Supply queryResults without the corresponding `queries` definition.
        queryResults: { q_private: { table: privateTable, meta: privateMeta } },
        privacy: { levelsBySourceId: { [apiSourceId]: "public" } },
      },
      {},
    ),
    /Formula\.Firewall/,
  );
});

test("createDesktopQueryEngine executes database sources via sql_query", async () => {
  const originalTauri = globalThis.__TAURI__;

  /** @type {{ cmd: string, args: any }[]} */
  const calls = [];

  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd, args) => {
        calls.push({ cmd, args });
        if (cmd === "sql_query") {
          return { columns: ["A"], types: { A: "number" }, rows: [[1]] };
        }
        throw new Error(`Unexpected invoke: ${cmd}`);
      },
    },
  };

  try {
    const engine = createDesktopQueryEngine({
      fileAdapter: { readText: async () => "", readBinary: async () => new Uint8Array() },
    });

    const query = {
      id: "q_db",
      name: "DB",
      source: { type: "database", connection: { kind: "sqlite", path: "/tmp/test.db" }, query: "SELECT 1 AS A" },
      steps: [],
    };

    const table = await engine.executeQuery(query, {}, {});
    assert.deepEqual(table.toGrid(), [["A"], [1]]);

    assert.equal(calls.length, 1);
    assert.equal(calls[0].cmd, "sql_query");
    assert.deepEqual(calls[0].args.connection, { kind: "sqlite", path: "/tmp/test.db" });
    assert.equal(calls[0].args.sql, "SELECT 1 AS A");
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});

test("database cache keys vary by connection identity and credentialId", async () => {
  let credentialId = "cred-a";
  const engine = createDesktopQueryEngine({
    fileAdapter: { readText: async () => "", readBinary: async () => new Uint8Array() },
    onCredentialRequest: async () => ({ credentialId }),
  });

  const baseQuery = {
    id: "q_db_cache",
    name: "DB Cache",
    source: { type: "database", connection: { kind: "sqlite", path: "/tmp/a.db" }, query: "SELECT 1" },
    steps: [],
  };

  const key1 = await engine.getCacheKey(baseQuery, {}, {});
  assert.ok(key1);

  const key2 = await engine.getCacheKey(
    { ...baseQuery, source: { ...baseQuery.source, connection: { kind: "sqlite", path: "/tmp/b.db" } } },
    {},
    {},
  );
  assert.notEqual(key1, key2);

  credentialId = "cred-b";
  const key3 = await engine.getCacheKey(baseQuery, {}, {});
  assert.notEqual(key1, key3);
});

test("sqlite connections with relative paths are treated as non-cacheable", async () => {
  const engine = createDesktopQueryEngine({
    fileAdapter: { readText: async () => "", readBinary: async () => new Uint8Array() },
  });

  const query = {
    id: "q_db_rel",
    name: "DB Rel",
    source: { type: "database", connection: { kind: "sqlite", path: "relative.db" }, query: "SELECT 1" },
    steps: [],
  };

  const key = await engine.getCacheKey(query, {}, {});
  assert.equal(key, null);
});

test("sqlite database cache entries are invalidated when the db file mtime changes", async () => {
  let mtimeMs = 1_000;
  let queryCount = 0;

  const engine = createDesktopQueryEngine({
    fileAdapter: {
      readText: async () => "",
      readBinary: async () => new Uint8Array(),
      stat: async () => ({ mtimeMs }),
    },
  });

  const originalTauri = globalThis.__TAURI__;
  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd) => {
        if (cmd === "sql_query") {
          queryCount += 1;
          return { columns: ["A"], types: { A: "number" }, rows: [[1]] };
        }
        throw new Error(`Unexpected invoke: ${cmd}`);
      },
    },
  };

  try {
    const query = {
      id: "q_db_cache_mtime",
      name: "DB Cache mtime",
      source: { type: "database", connection: { kind: "sqlite", path: "/tmp/test.db" }, query: "SELECT 1 AS A" },
      steps: [],
    };

    await engine.executeQuery(query, {}, {});
    await engine.executeQuery(query, {}, {});
    assert.equal(queryCount, 1, "expected second execution to reuse cached result");

    mtimeMs = 2_000;
    await engine.executeQuery(query, {}, {});
    assert.equal(queryCount, 2, "expected cache to invalidate when sqlite db mtime changes");
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});

test("sql_get_schema resolves credential handles before invoking Tauri", async () => {
  const originalTauri = globalThis.__TAURI__;

  /** @type {{ cmd: string, args: any }[]} */
  const calls = [];
  let secretCalls = 0;

  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd, args) => {
        calls.push({ cmd, args });
        if (cmd === "sql_get_schema") {
          return { columns: ["A"], types: { A: "number" } };
        }
        if (cmd === "sql_query") {
          return { columns: ["A"], types: { A: "number" }, rows: [[1]] };
        }
        throw new Error(`Unexpected invoke: ${cmd}`);
      },
    },
  };

  try {
    const engine = createDesktopQueryEngine({
      fileAdapter: { readText: async () => "", readBinary: async () => new Uint8Array() },
      onCredentialRequest: async () => ({
        credentialId: "cred-1",
        getSecret: async () => {
          secretCalls += 1;
          return { password: "pw" };
        },
      }),
    });

    const query = {
      id: "q_db_schema_creds",
      name: "DB Schema creds",
      source: {
        type: "database",
        connection: { kind: "postgres", host: "localhost", port: 5432, database: "db", user: "u" },
        query: "SELECT 1 AS A",
        dialect: "postgres",
      },
      steps: [],
    };

    await engine.executeQuery(query, {}, { cache: { validation: "none" } });

    assert.ok(
      calls.some((c) => c.cmd === "sql_get_schema" && c.args.credentials && c.args.credentials.password === "pw"),
      "expected sql_get_schema to receive resolved credential secret",
    );
    assert.ok(
      calls.some((c) => c.cmd === "sql_query" && c.args.credentials && c.args.credentials.password === "pw"),
      "expected sql_query to receive resolved credential secret",
    );
    assert.equal(secretCalls, 2, "expected credential handle getSecret() to be called for schema + query execution");
  } finally {
    globalThis.__TAURI__ = originalTauri;
  }
});
