const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const { installFakeWorker } = require("./helpers/fake-browser-worker");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

test("BrowserExtensionHost: viewActivated is broadcast before attempting activation", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let sawViewActivatedOnStartupExtension = false;
  const order = [];

  const scenarios = [
    // Worker for the startup extension (already running).
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }

        if (msg?.type === "event" && msg.event === "viewActivated") {
          sawViewActivatedOnStartupExtension = true;
        }
      }
    },
    // Worker for the view-gated extension (activation fails).
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "event" && msg.event === "viewActivated") {
          order.push("event:viewActivated");
          return;
        }

        if (msg?.type === "activate") {
          order.push("activate");
          worker.emitMessage({
            type: "activate_error",
            id: msg.id,
            error: { message: "Permission denied: ui.commands" }
          });
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => false
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension({
    extensionId: "test.view-logger",
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/logger.js",
    manifest: {
      name: "view-logger",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], panels: [], customFunctions: [], dataConnectors: [] },
      activationEvents: ["onStartupFinished"],
      permissions: []
    }
  });

  await host.loadExtension({
    extensionId: "test.sample-hello",
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/sample.js",
    manifest: {
      name: "sample-hello",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: {
        commands: [],
        panels: [{ id: "sampleHello.panel", title: "Sample Hello Panel" }],
        customFunctions: [],
        dataConnectors: []
      },
      activationEvents: ["onView:sampleHello.panel"],
      permissions: ["ui.commands", "ui.panels"]
    }
  });

  await host.startup();
  await host.activateView("sampleHello.panel");

  assert.equal(sawViewActivatedOnStartupExtension, true);
  assert.deepEqual(order, ["event:viewActivated", "activate"]);
});

test("BrowserExtensionHost: viewActivated is replayed to newly-activated extensions", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let startupEventCount = 0;
  let viewEventCount = 0;
  const viewOrder = [];

  const scenarios = [
    // Startup extension: should see the broadcast once.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "activate") {
          worker.emitMessage({ type: "activate_result", id: msg.id });
          return;
        }
        if (msg?.type === "event" && msg.event === "viewActivated") {
          startupEventCount += 1;
        }
      }
    },
    // View-gated extension: should see broadcast, then activation, then a replay after activation.
    {
      onPostMessage(msg, worker) {
        if (msg?.type === "event" && msg.event === "viewActivated") {
          viewEventCount += 1;
          viewOrder.push(`event:viewActivated:${viewEventCount}`);
          return;
        }
        if (msg?.type === "activate") {
          viewOrder.push("activate");
          worker.emitMessage({ type: "activate_result", id: msg.id });
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
    extensionId: "test.view-logger",
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/logger.js",
    manifest: {
      name: "view-logger",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], panels: [], customFunctions: [], dataConnectors: [] },
      activationEvents: ["onStartupFinished"],
      permissions: []
    }
  });

  await host.loadExtension({
    extensionId: "test.sample-hello",
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/sample.js",
    manifest: {
      name: "sample-hello",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: {
        commands: [],
        panels: [{ id: "sampleHello.panel", title: "Sample Hello Panel" }],
        customFunctions: [],
        dataConnectors: []
      },
      activationEvents: ["onView:sampleHello.panel"],
      permissions: ["ui.commands", "ui.panels"]
    }
  });

  await host.startup();
  await host.activateView("sampleHello.panel");

  assert.equal(startupEventCount, 1);
  assert.deepEqual(viewOrder, ["event:viewActivated:1", "activate", "event:viewActivated:2"]);
});

test("BrowserExtensionHost: activateView normalizes viewId to string in event payload", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  /** @type {any} */
  let payload = null;

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "event" && msg.event === "viewActivated") {
          payload = msg.data;
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
    extensionId: "test.view-listener",
    extensionPath: "http://example.invalid/",
    mainUrl: "http://example.invalid/listener.js",
    manifest: {
      name: "view-listener",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], panels: [], customFunctions: [], dataConnectors: [] },
      activationEvents: [],
      permissions: []
    }
  });

  await host.activateView(123);
  assert.deepEqual(payload, { viewId: "123" });
});
