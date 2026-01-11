import assert from "node:assert/strict";
import test from "node:test";

import { createDesktopQueryEngine } from "../engine.ts";
import { createDefaultOrgPolicy } from "../../../../../packages/security/dlp/src/policy.js";

test("createDesktopQueryEngine uses Tauri invoke file commands when FS plugin is unavailable", async () => {
  const originalTauri = globalThis.__TAURI__;

  /** @type {{ cmd: string, args: any }[]} */
  const calls = [];

  globalThis.__TAURI__ = {
    core: {
      invoke: async (cmd, args) => {
        calls.push({ cmd, args });
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

  await engine.executeQuery(query, {}, {});
  await engine.executeQuery(query, {}, {});

  assert.equal(promptCount, 1);
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
