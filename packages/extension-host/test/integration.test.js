const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");
const http = require("node:http");

const { ExtensionHost } = require("../src");

// These integration tests spin up extension workers (and sometimes local HTTP servers). Under heavy
// CI load, the default 5s activation/command timeouts can be too aggressive and cause spurious
// EXTENSION_TIMEOUT failures. Keep the tests focused on integration correctness by using a more
// forgiving timeout.
const INTEGRATION_TEST_TIMEOUT_MS = 30_000;

function createHost(options = {}) {
  return new ExtensionHost({
    ...options,
    activationTimeoutMs: options.activationTimeoutMs ?? INTEGRATION_TEST_TIMEOUT_MS,
    commandTimeoutMs: options.commandTimeoutMs ?? INTEGRATION_TEST_TIMEOUT_MS,
    customFunctionTimeoutMs: options.customFunctionTimeoutMs ?? INTEGRATION_TEST_TIMEOUT_MS,
    dataConnectorTimeoutMs: options.dataConnectorTimeoutMs ?? INTEGRATION_TEST_TIMEOUT_MS,
  });
}

test("integration: load sample extension and execute contributed command", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-int-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    // Worker startup can be slow under heavy CI load; keep integration tests focused on
    // correctness rather than the default 5s activation SLA.
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async ({ permissions }) => {
      // Allow command registration, but deny outbound network access.
      if (permissions.includes("network")) return false;
      return true;
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");

  await host.loadExtension(extPath);

  host.spreadsheet.setCell(0, 0, 1);
  host.spreadsheet.setCell(0, 1, 2);
  host.spreadsheet.setCell(1, 0, 3);
  host.spreadsheet.setCell(1, 1, 4);
  host.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });

  const result = await host.executeCommand("sampleHello.sumSelection");
  assert.equal(result, 10);
  assert.equal(host.spreadsheet.getCell(2, 0), 10);

  const messages = host.getMessages();
  assert.ok(messages.some((m) => String(m.message).includes("Sum: 10")));
});

test("integration: panel command creates panel and sets HTML", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-panel-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");

  await host.loadExtension(extPath);

  await host.executeCommand("sampleHello.openPanel");

  const panel = host.getPanel("sampleHello.panel");
  assert.ok(panel);
  assert.ok(panel.html.includes("Sample Hello Panel"));
});

test("integration: view activation creates and renders contributed panel", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-view-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");

  await host.loadExtension(extPath);

  await host.activateView("sampleHello.panel");

  const deadline = Date.now() + 500;
  while (Date.now() < deadline) {
    const panel = host.getPanel("sampleHello.panel");
    if (panel && panel.html.includes("Sample Hello Panel")) {
      return;
    }
    await new Promise((r) => setTimeout(r, 10));
  }

  assert.fail("Timed out waiting for panel HTML after view activation");
});

test("integration: viewActivated is broadcast even if view activation fails", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-view-activated-"));

  // Deny all permission prompts so the Sample Hello extension fails to activate on view open. The
  // goal of this test is to ensure `formula.events.onViewActivated` is still delivered to already
  // running extensions (i.e. it behaves like other event broadcasts, not like an activation hook).
  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async () => false
  });

  t.after(async () => {
    await host.dispose();
  });

  const sampleHelloPath = path.resolve(__dirname, "../../../extensions/sample-hello");
  const viewLoggerPath = path.resolve(__dirname, "./fixtures/view-logger");

  await host.loadExtension(sampleHelloPath);
  await host.loadExtension(viewLoggerPath);
  await host.startup();

  await assert.rejects(() => host.activateView("sampleHello.panel"), /Permission denied/);

  const expected = "[view-logger] viewActivated:string:sampleHello.panel";
  const deadline = Date.now() + 1_000;
  while (Date.now() < deadline) {
    const messages = host.getMessages().map((m) => String(m.message));
    if (messages.some((m) => m.includes(expected))) return;
    await new Promise((r) => setTimeout(r, 10));
  }

  assert.fail("Timed out waiting for viewActivated message from view-logger extension");
});

test("integration: viewActivated payload normalizes viewId to string", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-view-id-normalization-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const viewLoggerPath = path.resolve(__dirname, "./fixtures/view-logger");
  await host.loadExtension(viewLoggerPath);
  await host.startup();

  await host.activateView(123);

  const expected = "[view-logger] viewActivated:string:123";
  const deadline = Date.now() + 1_000;
  while (Date.now() < deadline) {
    const messages = host.getMessages().map((m) => String(m.message));
    if (messages.some((m) => m.includes(expected))) return;
    await new Promise((r) => setTimeout(r, 10));
  }

  assert.fail("Timed out waiting for normalized viewActivated payload");
});

test("integration: panel messaging (webview -> extension -> webview)", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-panel-msg-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");

  await host.loadExtension(extPath);

  await host.executeCommand("sampleHello.openPanel");

  host.dispatchPanelMessage("sampleHello.panel", { type: "ping" });

  const deadline = Date.now() + 500;
  while (Date.now() < deadline) {
    const outgoing = host.getPanelOutgoingMessages("sampleHello.panel");
    if (outgoing.some((m) => m && m.type === "pong")) {
      return;
    }
    await new Promise((r) => setTimeout(r, 10));
  }

  assert.fail("Timed out waiting for pong message");
});

test("integration: invoke custom function activates extension and returns result", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-fn-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");

  await host.loadExtension(extPath);

  const result = await host.invokeCustomFunction("SAMPLEHELLO_DOUBLE", 5);
  assert.equal(result, 10);
});

test("integration: network.fetch is permission gated and can fetch via host", async (t) => {
  const server = http.createServer((req, res) => {
    res.writeHead(200, { "Content-Type": "text/plain" });
    res.end("hello");
  });

  await new Promise((resolve) => server.listen(0, resolve));
  const address = server.address();
  const port = typeof address === "object" && address ? address.port : null;
  if (!port) throw new Error("Failed to allocate test port");

  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-net-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await new Promise((resolve) => server.close(resolve));
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");
  await host.loadExtension(extPath);

  const url = `http://127.0.0.1:${port}/`;
  const text = await host.executeCommand("sampleHello.fetchText", url);
  assert.equal(text, "hello");
  assert.ok(host.getMessages().some((m) => String(m.message).includes("Fetched: hello")));
});

test("integration: denied network permission blocks fetch", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-net-deny-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    // Worker startup can be slow under heavy CI load; keep integration tests focused on
    // correctness rather than the default 5s activation SLA.
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async ({ permissions }) => {
      if (permissions.includes("network")) return false;
      return true;
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");
  await host.loadExtension(extPath);

  await assert.rejects(
    () => host.executeCommand("sampleHello.fetchText", "http://example.invalid/"),
    /Permission denied/
  );
});

test("integration: clipboard API is permission gated and writes clipboard text", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-clipboard-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");
  await host.loadExtension(extPath);

  host.spreadsheet.setCell(0, 0, 1);
  host.spreadsheet.setCell(0, 1, 2);
  host.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 0, endCol: 1 });

  const sum = await host.executeCommand("sampleHello.copySumToClipboard");
  assert.equal(sum, 3);
  assert.equal(host.getClipboardText(), "3");
});

test("integration: denied clipboard permission blocks clipboard writes", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-clipboard-deny-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async ({ permissions }) => {
      if (permissions.includes("clipboard")) return false;
      return true;
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");
  await host.loadExtension(extPath);

  host.spreadsheet.setCell(0, 0, 1);
  host.spreadsheet.setCell(0, 1, 2);
  host.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 0, endCol: 1 });

  await assert.rejects(
    () => host.executeCommand("sampleHello.copySumToClipboard"),
    /Permission denied/
  );
  assert.equal(host.getClipboardText(), "");
});

test("integration: config.get returns contributed default values", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-config-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");
  await host.loadExtension(extPath);

  const greeting = await host.executeCommand("sampleHello.showGreeting");
  assert.equal(greeting, "Hello");
  assert.ok(host.getMessages().some((m) => String(m.message).includes("Greeting: Hello")));
});

test("integration: denied permission prevents side effects", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-deny-"));

  const host = createHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    activationTimeoutMs: INTEGRATION_TEST_TIMEOUT_MS,
    permissionPrompt: async ({ permissions }) => {
      // Allow command registration, but deny write access.
      if (permissions.includes("cells.write")) return false;
      return true;
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  const extPath = path.resolve(__dirname, "../../../extensions/sample-hello");

  await host.loadExtension(extPath);

  host.spreadsheet.setCell(0, 0, 1);
  host.spreadsheet.setCell(0, 1, 2);
  host.spreadsheet.setSelection({ startRow: 0, startCol: 0, endRow: 0, endCol: 1 });

  await assert.rejects(() => host.executeCommand("sampleHello.sumSelection"), /Permission denied/);
  assert.equal(host.spreadsheet.getCell(1, 0), null);
});
