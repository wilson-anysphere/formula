import assert from "node:assert/strict";
import test from "node:test";

import { createDesktopQueryEngine } from "../engine.ts";
import { createDefaultOrgPolicy } from "../../../../../packages/security/dlp/src/policy.js";
import { DLP_ACTION } from "../../../../../packages/security/dlp/src/actions.js";
import { getHttpSourceId } from "../../../../../packages/power-query/src/privacy/sourceId.js";

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
