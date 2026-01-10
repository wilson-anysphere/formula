const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");
const http = require("node:http");

const { ExtensionHost } = require("../src");

test("integration: load sample extension and execute contributed command", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-int-"));

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

  await host.executeCommand("sampleHello.openPanel");

  const panel = host.getPanel("sampleHello.panel");
  assert.ok(panel);
  assert.ok(panel.html.includes("Sample Hello Panel"));
});

test("integration: panel messaging (webview -> extension -> webview)", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-panel-msg-"));

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

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
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

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
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

test("integration: denied permission prevents side effects", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-deny-"));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
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
