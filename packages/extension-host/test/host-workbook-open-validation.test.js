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
