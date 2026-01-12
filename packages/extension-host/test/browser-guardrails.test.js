const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const { installFakeWorker } = require("./helpers/fake-browser-worker");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

test("BrowserExtensionHost: init message includes context storage fields", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type !== "init") return;
        assert.ok(typeof msg.extensionUri === "string" && msg.extensionUri.length > 0);
        assert.ok(typeof msg.globalStoragePath === "string" && msg.globalStoragePath.includes("globalStorage"));
        assert.ok(
          typeof msg.workspaceStoragePath === "string" && msg.workspaceStoragePath.includes("workspaceStorage")
        );
        assert.equal(msg.sandbox?.strictImports, true);
        assert.equal(msg.sandbox?.disableEval, true);
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      }
    },
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension({
    extensionId: "test.context",
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "context",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: []
    }
  });
});

test("BrowserExtensionHost: sandbox options are configurable", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type !== "init") return;
        assert.equal(msg.sandbox?.strictImports, false);
        assert.equal(msg.sandbox?.disableEval, false);
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true,
    sandbox: { strictImports: false, disableEval: false }
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension({
    extensionId: "test.sandbox-options",
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "sandbox-options",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: []
    }
  });
});

test("BrowserExtensionHost: terminating a worker clears runtime context menus", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  /** @type {(value?: unknown) => void} */
  let resolveApiResult;
  const apiResult = new Promise((resolve) => {
    resolveApiResult = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_result" && msg.id === "req1") resolveApiResult();
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      }
    },
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.menus";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "menus",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["ui.menus"]
    }
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension);
  assert.ok(extension.worker);

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "ui",
    method: "registerContextMenu",
    args: ["cell/context", [{ command: "test.cmd" }]]
  });

  await apiResult;
  assert.equal(host._contextMenus.size, 1);

  host._terminateWorker(extension, { reason: "crash", cause: new Error("boom") });
  assert.equal(host._contextMenus.size, 0);
});

test("BrowserExtensionHost: startup broadcasts workbookOpened with default workbook payload", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let sawWorkbookOpened = false;

  const scenarios = [
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "event" && msg.event === "workbookOpened") {
          sawWorkbookOpened = true;
          assert.deepEqual(msg.data, {
            workbook: {
              name: "MockWorkbook",
              path: null,
              sheets: [{ id: "sheet1", name: "Sheet1" }],
              activeSheet: { id: "sheet1", name: "Sheet1" }
            }
          });
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension({
    extensionId: "test.startup-workbook",
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "startup-workbook",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: ["onStartupFinished"],
      permissions: []
    }
  });

  await host.startup();
  assert.equal(sawWorkbookOpened, true);
});

test("BrowserExtensionHost: activation timeout terminates worker and allows restart", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const scenarios = [
    // First worker: hang activation.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "activate") {
          // hang forever
          return;
        }
        if (msg?.type === "init") return;
        worker.emitMessage({ type: "log", level: "info", args: ["unexpected", msg] });
      }
    },
    // Second worker: activation succeeds.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      },
      async getCell() {
        return null;
      },
      async setCell() {
        // noop
      }
    },
    permissionPrompt: async () => true,
    activationTimeoutMs: 50
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.activation-timeout";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "activation-timeout",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: ["onStartupFinished"],
      permissions: []
    }
  });

  await assert.rejects(() => host.startup(), /timed out/i);
  assert.equal(host._extensions.get(extensionId).active, false);

  await host.startup();
  assert.equal(host._extensions.get(extensionId).active, true);
});

test("BrowserExtensionHost: activation errors preserve name/code", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const scenarios = [
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({
            type: "activate_error",
            id: msg.id,
            error: { message: "boom", name: "BoomError", code: "BOOM" }
          });
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.activation-error-metadata";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "activation-error-metadata",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: ["onStartupFinished"],
      permissions: []
    }
  });

  await assert.rejects(() => host.startup(), (err) => {
    assert.equal(err?.name, "BoomError");
    assert.equal(err?.code, "BOOM");
    assert.equal(err?.message, "boom");
    return true;
  });
});

test("BrowserExtensionHost: command timeout terminates worker, rejects in-flight requests, and restarts", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const scenarios = [
    // First worker: activation succeeds; command handlers hang forever.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "execute_command") {
          // hang forever
          return;
        }
      }
    },
    // Second worker: activation succeeds; commands resolve.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "execute_command") {
          worker.emitMessage({ type: "command_result", id: msg.id, result: "ok" });
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      },
      async getCell() {
        return null;
      },
      async setCell() {
        // noop
      }
    },
    permissionPrompt: async () => true,
    activationTimeoutMs: 1000,
    commandTimeoutMs: 50
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.command-timeout";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "command-timeout",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: {
        commands: [
          { command: "test.hang", title: "Hang" },
          { command: "test.ok", title: "Ok" }
        ],
        customFunctions: []
      },
      activationEvents: ["onCommand:test.hang", "onCommand:test.ok"],
      permissions: ["ui.commands"]
    }
  });

  const hangPromise = host.executeCommand("test.hang");
  await new Promise((r) => setTimeout(r, 10));
  const pendingPromise = host.executeCommand("test.ok");

  await assert.rejects(() => hangPromise, /timed out/i);
  await assert.rejects(() => pendingPromise, /worker terminated/i);
  assert.equal(host._extensions.get(extensionId).active, false);

  assert.equal(await host.executeCommand("test.ok"), "ok");
});

test("BrowserExtensionHost: custom function timeout terminates worker and allows restart", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const scenarios = [
    // First worker: activation succeeds; custom function invocation hangs.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "invoke_custom_function") {
          return;
        }
      }
    },
    // Second worker: activation succeeds; custom function resolves.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "invoke_custom_function") {
          worker.emitMessage({ type: "custom_function_result", id: msg.id, result: 42 });
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      },
      async getCell() {
        return null;
      },
      async setCell() {
        // noop
      }
    },
    permissionPrompt: async () => true,
    activationTimeoutMs: 1000,
    customFunctionTimeoutMs: 50
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.custom-function-timeout";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "custom-function-timeout",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: {
        commands: [],
        customFunctions: [
          { name: "TEST_HANG", description: "hang", parameters: [], result: { type: "number" } }
        ]
      },
      activationEvents: ["onCustomFunction:TEST_HANG"],
      permissions: []
    }
  });

  await assert.rejects(() => host.invokeCustomFunction("TEST_HANG"), /timed out/i);
  assert.equal(host._extensions.get(extensionId).active, false);

  assert.equal(await host.invokeCustomFunction("TEST_HANG"), 42);
});

test("BrowserExtensionHost: worker termination removes extension-owned context menu registrations", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  const scenarios = [
    // First worker: activation succeeds; command hangs.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "execute_command") return;
      }
    },
    // Second worker: activation succeeds; commands resolve.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "init") return;
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "execute_command") {
          worker.emitMessage({ type: "command_result", id: msg.id, result: "ok" });
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async getSelection() {
        return { startRow: 0, startCol: 0, endRow: 0, endCol: 0, values: [[null]] };
      },
      async getCell() {
        return null;
      },
      async setCell() {
        // noop
      }
    },
    permissionPrompt: async () => true,
    activationTimeoutMs: 1000,
    commandTimeoutMs: 50
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.context-menu-cleanup";
  await host.loadExtension({
    extensionId,
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/main.js",
    manifest: {
      name: "context-menu-cleanup",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: {
        commands: [{ command: "test.hang", title: "Hang" }],
        customFunctions: []
      },
      activationEvents: ["onCommand:test.hang"],
      permissions: ["ui.commands", "ui.menus"]
    }
  });

  const registration = await host._executeApi("ui", "registerContextMenu", [
    "cell/context",
    [{ command: "test.hang" }]
  ], host._extensions.get(extensionId));
  assert.ok(registration?.id);
  assert.equal(host._contextMenus.size, 1);

  await assert.rejects(() => host.executeCommand("test.hang"), /timed out/i);
  assert.equal(host._contextMenus.size, 0);

  assert.equal(await host.executeCommand("test.hang"), "ok");
});
