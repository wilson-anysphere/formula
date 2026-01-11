const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { ExtensionHost } = require("../src");

async function writeExtensionFixture(extensionDir, manifest, entrypointCode) {
  await fs.mkdir(path.join(extensionDir, "dist"), { recursive: true });
  await fs.writeFile(path.join(extensionDir, "package.json"), JSON.stringify(manifest, null, 2), "utf8");
  await fs.writeFile(path.join(extensionDir, "dist", "extension.js"), entrypointCode, "utf8");
}

test("api surface: cells.getRange/setRange roundtrip uses A1 refs and serializes values", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-range-roundtrip-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "rangeExt.roundtrip";
  await writeExtensionFixture(
    extDir,
    {
      name: "range-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Range Roundtrip" }] },
      permissions: ["ui.commands", "cells.read", "cells.write"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          await formula.cells.setRange("A1:B2", [[1, 2], [3, 4]]);
          return formula.cells.getRange("A1:B2");
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

  await host.loadExtension(extDir);

  const range = await host.executeCommand(commandId);
  assert.deepEqual(range, {
    startRow: 0,
    startCol: 0,
    endRow: 1,
    endCol: 1,
    values: [
      [1, 2],
      [3, 4]
    ]
  });

  assert.equal(host.spreadsheet.getCell(0, 0), 1);
  assert.equal(host.spreadsheet.getCell(1, 1), 4);
});

test("permissions: cells.getRange requires cells.read", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-range-perm-read-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "rangeExt.getDenied";
  await writeExtensionFixture(
    extDir,
    {
      name: "range-ext-read-denied",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Range Get Denied" }] },
      permissions: ["ui.commands", "cells.read"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          return formula.cells.getRange("A1:A1");
        }));
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async ({ permissions }) => !permissions.includes("cells.read")
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(() => host.executeCommand(commandId), /Permission denied: cells\.read/);
});

test("permissions: cells.setRange requires cells.write and denial prevents writes", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-range-perm-write-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "rangeExt.setDenied";
  await writeExtensionFixture(
    extDir,
    {
      name: "range-ext-write-denied",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Range Set Denied" }] },
      permissions: ["ui.commands", "cells.write"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          await formula.cells.setRange("A1:B2", [[1, 2], [3, 4]]);
          return "ok";
        }));
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async ({ permissions }) => !permissions.includes("cells.write")
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(() => host.executeCommand(commandId), /Permission denied: cells\.write/);
  assert.equal(host.spreadsheet.getCell(0, 0), null);
  assert.equal(host.spreadsheet.getCell(1, 1), null);
});

test("api surface: sheets.createSheet/renameSheet/getSheet manage workbook sheets", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sheets-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "sheetExt.manage";
  await writeExtensionFixture(
    extDir,
    {
      name: "sheet-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Sheet Manage" }] },
      permissions: ["ui.commands", "sheets.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          await formula.sheets.createSheet("Data");
          await formula.sheets.renameSheet("Data", "Data2");
          const sheet = await formula.sheets.getSheet("Data2");
          const missing = await formula.sheets.getSheet("Data");
          return { sheet, missing };
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

  await host.loadExtension(extDir);

  const result = await host.executeCommand(commandId);
  assert.equal(result.sheet.name, "Data2");
  assert.ok(result.sheet.id);
  assert.equal(result.missing, undefined);
});

test("permissions: sheets.createSheet requires sheets.manage", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sheets-deny-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "sheetExt.denied";
  await writeExtensionFixture(
    extDir,
    {
      name: "sheet-ext-denied",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Sheet Denied" }] },
      permissions: ["ui.commands", "sheets.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          return formula.sheets.createSheet("Denied");
        }));
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async ({ permissions }) => !permissions.includes("sheets.manage")
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(() => host.executeCommand(commandId), /Permission denied: sheets\.manage/);
});

test("events: sheets.createSheet emits sheetActivated with stable payload", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sheet-activated-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "sheetExt.activatedEvent";
  await writeExtensionFixture(
    extDir,
    {
      name: "sheet-activated-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Sheet Activated Event" }] },
      permissions: ["ui.commands", "sheets.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          const eventPromise = new Promise((resolve) => {
            const disp = formula.events.onSheetActivated((e) => {
              disp.dispose();
              resolve(e);
            });
          });
          const sheet = await formula.sheets.createSheet("Data");
          const evt = await eventPromise;
          return { sheet, evt };
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

  await host.loadExtension(extDir);

  const result = await host.executeCommand(commandId);
  assert.equal(result.sheet.name, "Data");
  assert.ok(result.sheet.id);
  assert.deepEqual(result.evt, { sheet: result.sheet });
});

test("api surface: workbook.openWorkbook emits workbookOpened with stable payload", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-workbook-open-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "workbookExt.open";
  await writeExtensionFixture(
    extDir,
    {
      name: "workbook-open-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Workbook Open" }] },
      permissions: ["ui.commands", "workbook.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async (workbookPath) => {
          const eventPromise = new Promise((resolve) => {
            const disp = formula.events.onWorkbookOpened((e) => {
              disp.dispose();
              resolve(e);
            });
          });
          const workbook = await formula.workbook.openWorkbook(workbookPath);
          const evt = await eventPromise;
          return { workbook, evt };
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

  await host.loadExtension(extDir);

  const workbookPath = path.join(dir, "Book1.xlsx");
  const result = await host.executeCommand(commandId, workbookPath);
  assert.deepEqual(result.workbook, { name: "Book1.xlsx", path: workbookPath });
  assert.deepEqual(result.evt, { workbook: result.workbook });
});

test("permissions: workbook.openWorkbook requires workbook.manage", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-workbook-open-deny-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "workbookExt.openDenied";
  await writeExtensionFixture(
    extDir,
    {
      name: "workbook-open-denied-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Workbook Open Denied" }] },
      permissions: ["ui.commands", "workbook.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          return formula.workbook.openWorkbook("Denied.xlsx");
        }));
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async ({ permissions }) => !permissions.includes("workbook.manage")
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);

  await assert.rejects(() => host.executeCommand(commandId), /Permission denied: workbook\.manage/);
});
