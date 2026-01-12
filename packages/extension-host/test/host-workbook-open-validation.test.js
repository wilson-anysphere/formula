const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

test("ExtensionHost: workbook.openWorkbook rejects whitespace-only paths", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-workbook-open-whitespace-"));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  await assert.rejects(
    () => host._executeApi("workbook", "openWorkbook", ["   "], { id: "test" }),
    /Workbook path must be a non-empty string/,
  );
});

test("ExtensionHost: workbook.saveAs rejects whitespace-only paths", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-workbook-saveas-whitespace-"));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  await assert.rejects(
    () => host._executeApi("workbook", "saveAs", ["   "], { id: "test" }),
    /Workbook path must be a non-empty string/,
  );
});

test("ExtensionHost: workbook.openWorkbook rejects non-string paths", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-workbook-open-nonstring-"));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  await assert.rejects(
    () => host._executeApi("workbook", "openWorkbook", [123], { id: "test" }),
    /Workbook path must be a non-empty string/,
  );
});

test("ExtensionHost: workbook.saveAs rejects non-string paths", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-workbook-saveas-nonstring-"));

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  await assert.rejects(
    () => host._executeApi("workbook", "saveAs", [123], { id: "test" }),
    /Workbook path must be a non-empty string/,
  );
});

test("ExtensionHost: invalid workbook.openWorkbook path does not prompt for permissions", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-workbook-open-no-prompt-"));

  let promptCalls = 0;
  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  /** @type {any[]} */
  const messages = [];

  const extension = {
    id: "test.no-prompt-open",
    manifest: {
      name: "no-prompt-open",
      permissions: ["workbook.manage"],
    },
    worker: {
      postMessage(msg) {
        messages.push(msg);
      },
    },
  };

  await host._handleApiCall(extension, {
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: ["   "],
  });

  assert.equal(promptCalls, 0);
  assert.equal(messages.length, 1);
  assert.equal(messages[0]?.type, "api_error");
  assert.equal(messages[0]?.error?.message, "Workbook path must be a non-empty string");
});

test("ExtensionHost: invalid workbook.saveAs path does not prompt for permissions", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-workbook-saveas-no-prompt-"));

  let promptCalls = 0;
  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  /** @type {any[]} */
  const messages = [];

  const extension = {
    id: "test.no-prompt-saveas",
    manifest: {
      name: "no-prompt-saveas",
      permissions: ["workbook.manage"],
    },
    worker: {
      postMessage(msg) {
        messages.push(msg);
      },
    },
  };

  await host._handleApiCall(extension, {
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "saveAs",
    args: ["   "],
  });

  assert.equal(promptCalls, 0);
  assert.equal(messages.length, 1);
  assert.equal(messages[0]?.type, "api_error");
  assert.equal(messages[0]?.error?.message, "Workbook path must be a non-empty string");
});

test("ExtensionHost: workbook.openWorkbook missing path argument does not prompt for permissions", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-workbook-open-missing-no-prompt-"));

  let promptCalls = 0;
  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  /** @type {any[]} */
  const messages = [];

  const extension = {
    id: "test.no-prompt-open-missing",
    manifest: {
      name: "no-prompt-open-missing",
      permissions: ["workbook.manage"],
    },
    worker: {
      postMessage(msg) {
        messages.push(msg);
      },
    },
  };

  await host._handleApiCall(extension, {
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: [],
  });

  assert.equal(promptCalls, 0);
  assert.equal(messages.length, 1);
  assert.equal(messages[0]?.type, "api_error");
  assert.equal(messages[0]?.error?.message, "Workbook path must be a non-empty string");
});

test("ExtensionHost: workbook.saveAs missing path argument does not prompt for permissions", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-workbook-saveas-missing-no-prompt-"));

  let promptCalls = 0;
  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  /** @type {any[]} */
  const messages = [];

  const extension = {
    id: "test.no-prompt-saveas-missing",
    manifest: {
      name: "no-prompt-saveas-missing",
      permissions: ["workbook.manage"],
    },
    worker: {
      postMessage(msg) {
        messages.push(msg);
      },
    },
  };

  await host._handleApiCall(extension, {
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "saveAs",
    args: [],
  });

  assert.equal(promptCalls, 0);
  assert.equal(messages.length, 1);
  assert.equal(messages[0]?.type, "api_error");
  assert.equal(messages[0]?.error?.message, "Workbook path must be a non-empty string");
});

test("ExtensionHost: malformed api_call args do not prompt for permissions", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-host-workbook-malformed-args-"));

  let promptCalls = 0;
  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  /** @type {any[]} */
  const messages = [];

  const extension = {
    id: "test.no-prompt-malformed",
    manifest: {
      name: "no-prompt-malformed",
      permissions: ["workbook.manage"],
    },
    worker: {
      postMessage(msg) {
        messages.push(msg);
      },
    },
  };

  // The host should treat non-array args as "no args" rather than indexing into the string.
  await host._handleApiCall(extension, {
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: "Book.xlsx",
  });

  assert.equal(promptCalls, 0);
  assert.equal(messages.length, 1);
  assert.equal(messages[0]?.type, "api_error");
  assert.equal(messages[0]?.error?.message, "Workbook path must be a non-empty string");
});
