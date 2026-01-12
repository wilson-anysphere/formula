const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

async function writeExtension(dir, manifest, entrypointSource) {
  await fs.mkdir(dir, { recursive: true });
  await fs.writeFile(path.join(dir, "package.json"), JSON.stringify(manifest, null, 2));
  await fs.writeFile(path.join(dir, manifest.main), entrypointSource);
}

test("data connectors: invocation activates the extension and returns results", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-data-connector-"));
  const extDir = path.join(dir, "ext");

  const connectorId = "test.connector";
  await writeExtension(
    extDir,
    {
      name: "data-connector-ok",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onDataConnector:${connectorId}`],
      contributes: {
        dataConnectors: [{ id: connectorId, name: "Test Connector" }]
      }
    },
    `
      const formula = require("formula");

      module.exports.activate = async (context) => {
        context.subscriptions.push(await formula.dataConnectors.register(${JSON.stringify(connectorId)}, {
          async browse(config, path) {
            return { config, path: path ?? null, ok: true };
          },
          async query(config, query) {
            return {
              columns: ["configValue", "queryValue"],
              rows: [[config?.value ?? null, query?.value ?? null]]
            };
          }
        }));
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = await host.loadExtension(extDir);
  assert.equal(host.listExtensions().find((e) => e.id === extensionId)?.active, false);

  const browse = await host.invokeDataConnector(connectorId, "browse", { value: 1 });
  assert.deepEqual(browse, { config: { value: 1 }, path: null, ok: true });

  const query = await host.invokeDataConnector(connectorId, "query", { value: "a" }, { value: "b" });
  assert.deepEqual(query, { columns: ["configValue", "queryValue"], rows: [["a", "b"]] });

  assert.equal(host.listExtensions().find((e) => e.id === extensionId)?.active, true);
});

test("data connectors: registration rejected when connector not declared in manifest", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-data-connector-invalid-"));
  const extDir = path.join(dir, "ext");

  await writeExtension(
    extDir,
    {
      name: "data-connector-invalid",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: ["onStartupFinished"]
    },
    `
      const formula = require("formula");

      module.exports.activate = async () => {
        // Not declared in contributes.dataConnectors, so the host should reject registration.
        await formula.dataConnectors.register("test.undeclared", {
          async browse() { return []; },
          async query() { return { columns: [], rows: [] }; }
        });
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    // Activation can be slow on contended runners; keep this high enough to
    // avoid flaking while still exercising the "undeclared connector" guard.
    activationTimeoutMs: 20_000
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);
  await assert.rejects(() => host.startup(), /Data connector not declared in manifest/i);
});

test("data connectors: invoke requires onDataConnector activation event when extension is inactive", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-data-connector-no-event-"));
  const extDir = path.join(dir, "ext");

  const connectorId = "test.connector";
  await writeExtension(
    extDir,
    {
      name: "data-connector-no-event",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [],
      contributes: {
        dataConnectors: [{ id: connectorId, name: "Test Connector" }]
      }
    },
    `
      module.exports.activate = async () => {};
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);
  await assert.rejects(
    () => host.invokeDataConnector(connectorId, "browse", {}),
    /not activated for onDataConnector/i
  );
});

test("data connectors: timeout terminates worker, rejects in-flight requests, and allows restart", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-data-connector-timeout-"));
  const extDir = path.join(dir, "ext");

  const connectorId = "test.connector";
  await writeExtension(
    extDir,
    {
      name: "data-connector-timeout",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onDataConnector:${connectorId}`],
      contributes: {
        dataConnectors: [{ id: connectorId, name: "Timeout Connector" }]
      }
    },
    `
      const formula = require("formula");

      module.exports.activate = async (context) => {
        context.subscriptions.push(await formula.dataConnectors.register(${JSON.stringify(connectorId)}, {
          async browse() {
            return { ok: true };
          },
          async query() {
            // Block the worker thread event loop so subsequent requests cannot be processed.
            // eslint-disable-next-line no-constant-condition
            while (true) {}
          }
        }));
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
    // Worker startup can be slow under heavy CI load; keep activation timeout generous so this
    // test exercises worker termination + restart rather than flaking on activation.
    activationTimeoutMs: 20_000,
    dataConnectorTimeoutMs: 100
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  assert.deepEqual(await host.invokeDataConnector(connectorId, "browse", {}), { ok: true });

  const hangPromise = host.invokeDataConnector(connectorId, "query", {}, {});
  await new Promise((r) => setTimeout(r, 10));
  const pendingPromise = host.invokeDataConnector(connectorId, "browse", {});

  await assert.rejects(() => hangPromise, /timed out/i);
  await assert.rejects(() => pendingPromise, /worker terminated/i);

  assert.deepEqual(await host.invokeDataConnector(connectorId, "browse", {}), { ok: true });
});

test("data connectors: unloadExtension releases connector ids for future loads", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-data-connector-unload-"));
  const extADir = path.join(dir, "ext-a");
  const extBDir = path.join(dir, "ext-b");

  const connectorId = "test.connector";

  await writeExtension(
    extADir,
    {
      name: "data-connector-unload-a",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      contributes: {
        dataConnectors: [{ id: connectorId, name: "Test Connector" }]
      }
    },
    `module.exports.activate = async () => {};`
  );

  await writeExtension(
    extBDir,
    {
      name: "data-connector-unload-b",
      version: "1.0.0",
      publisher: "test",
      main: "extension.js",
      engines: { formula: "^1.0.0" },
      contributes: {
        dataConnectors: [{ id: connectorId, name: "Test Connector" }]
      }
    },
    `module.exports.activate = async () => {};`
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionAId = await host.loadExtension(extADir);
  await host.unloadExtension(extensionAId);

  await assert.doesNotReject(() => host.loadExtension(extBDir));
});

test("sample-hello: sample data connector can be invoked", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sample-connector-"));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");
  await host.loadExtension(extPath);

  const result = await host.invokeDataConnector("sampleHello.connector", "query", {}, { demo: true });
  assert.deepEqual(result, {
    columns: ["id", "label"],
    rows: [
      [1, "hello"],
      [2, "world"]
    ]
  });
});
