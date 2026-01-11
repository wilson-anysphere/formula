const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const cjsApi = require("..");

test("dual entrypoint: CJS and ESM exports stay in sync", async () => {
  const esmPath = pathToFileURL(path.join(__dirname, "..", "index.mjs")).href;
  const esmApi = await import(esmPath);

  const cjsKeys = Object.keys(cjsApi).sort();
  const esmKeys = Object.keys(esmApi).filter((k) => k !== "default").sort();

  assert.deepEqual(
    esmKeys,
    cjsKeys,
    `ESM export surface drifted.\nCJS: ${cjsKeys.join(", ")}\nESM: ${esmKeys.join(", ")}`
  );

  const namespaces = [
    "workbook",
    "sheets",
    "cells",
    "commands",
    "functions",
    "network",
    "clipboard",
    "ui",
    "storage",
    "config",
    "events"
  ];

  for (const ns of namespaces) {
    assert.equal(typeof cjsApi[ns], "object", `Expected CJS ${ns} to be an object`);
    assert.equal(typeof esmApi[ns], "object", `Expected ESM ${ns} to be an object`);
    assert.deepEqual(
      Object.keys(esmApi[ns]).sort(),
      Object.keys(cjsApi[ns]).sort(),
      `Namespace ${ns} drifted between entrypoints`
    );
  }
});

test("dual entrypoint: transport/context state is shared between CJS and ESM", async () => {
  const esmPath = pathToFileURL(path.join(__dirname, "..", "index.mjs")).href;
  const esmApi = await import(esmPath);

  /** @type {any} */
  let lastCall = null;
  cjsApi.__setTransport({
    postMessage: (message) => {
      lastCall = message;
      if (message?.type === "api_call") {
        queueMicrotask(() => {
          // Respond via the opposite entrypoint to ensure pending state is shared.
          cjsApi.__handleMessage({ type: "api_result", id: message.id, result: 42 });
        });
      }
    }
  });
  cjsApi.__setContext({ extensionId: "formula.test", extensionPath: "/tmp/ext" });

  assert.equal(esmApi.context.extensionId, "formula.test");
  assert.equal(esmApi.context.extensionPath, "/tmp/ext");

  const value = await esmApi.cells.getCell(1, 2);
  assert.equal(value, 42);

  assert.deepEqual(lastCall, {
    type: "api_call",
    id: lastCall.id,
    namespace: "cells",
    method: "getCell",
    args: [1, 2]
  });
});
