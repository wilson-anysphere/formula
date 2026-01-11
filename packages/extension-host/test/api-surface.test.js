const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");
const { pathToFileURL } = require("node:url");

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

test("permissions: sheets.activateSheet requires sheets.manage", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sheets-activate-deny-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);
 
  const commandId = "sheetExt.activateDenied";
  await writeExtensionFixture(
    extDir,
    {
      name: "sheet-activate-denied-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Sheet Activate Denied" }] },
      permissions: ["ui.commands", "sheets.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          return formula.sheets.activateSheet("Sheet1");
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

test("api surface: sheets.deleteSheet removes sheets and cannot delete last sheet", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sheets-delete-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "sheetExt.delete";
  await writeExtensionFixture(
    extDir,
    {
      name: "sheet-delete-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Sheet Delete" }] },
      permissions: ["ui.commands", "sheets.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          await formula.sheets.createSheet("Temp");
          await formula.sheets.deleteSheet("Temp");
          const missing = await formula.sheets.getSheet("Temp");
          let lastSheetError = null;
          try {
            await formula.sheets.deleteSheet("Sheet1");
          } catch (err) {
            lastSheetError = String(err && err.message ? err.message : err);
          }
          return { missing, lastSheetError };
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
  assert.equal(result.missing, undefined);
  assert.match(result.lastSheetError, /Cannot delete the last remaining sheet/);
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

test("api surface: sheets.activateSheet switches active sheet and emits sheetActivated", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-sheet-activate-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "sheetExt.activate";
  await writeExtensionFixture(
    extDir,
    {
      name: "sheet-activate-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Sheet Activate" }] },
      permissions: ["ui.commands", "sheets.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          await formula.sheets.createSheet("Data");

          const eventPromise = new Promise((resolve) => {
            const disp = formula.events.onSheetActivated((e) => {
              disp.dispose();
              resolve(e);
            });
          });

          const sheet = await formula.sheets.activateSheet("Sheet1");
          const evt = await eventPromise;
          const active = await formula.sheets.getActiveSheet();
          return { sheet, evt, active };
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
  assert.deepEqual(result, {
    sheet: { id: "sheet1", name: "Sheet1" },
    evt: { sheet: { id: "sheet1", name: "Sheet1" } },
    active: { id: "sheet1", name: "Sheet1" }
  });
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

test("permissions: workbook.saveAs requires workbook.manage and denial prevents updating workbook path", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-workbook-saveas-deny-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "workbookExt.saveAsDenied";
  await writeExtensionFixture(
    extDir,
    {
      name: "workbook-saveas-denied-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Workbook SaveAs Denied" }] },
      permissions: ["ui.commands", "workbook.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async (nextPath) => {
          await formula.workbook.saveAs(nextPath);
          return true;
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

  const initialPath = path.join(dir, "Initial.xlsx");
  host.openWorkbook(initialPath);

  const nextPath = path.join(dir, "Next.xlsx");
  await assert.rejects(() => host.executeCommand(commandId, nextPath), /Permission denied: workbook\.manage/);

  assert.deepEqual(host._workbook, { name: "Initial.xlsx", path: initialPath });
});

test("events: workbook.save emits beforeSave with stable payload", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-before-save-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "beforeSaveExt.save";
  await writeExtensionFixture(
    extDir,
    {
      name: "before-save-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Before Save" }] },
      permissions: ["ui.commands", "workbook.manage"]
    },
    `
      const formula = require("@formula/extension-api");

      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          const eventPromise = new Promise((resolve) => {
            const disp = formula.events.onBeforeSave((e) => {
              disp.dispose();
              resolve(e);
            });
          });

          await formula.workbook.save();
          return eventPromise;
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

  const workbookPath = path.join(dir, "Book2.xlsx");
  host.openWorkbook(workbookPath);
  const evt = await host.executeCommand(commandId);
  assert.deepEqual(evt, { workbook: { name: "Book2.xlsx", path: workbookPath } });
});

test("api surface: config.onDidChange fires after config.update and value persists", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-config-change-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "configExt.update";
  const configKey = "configExt.greeting";
  await writeExtensionFixture(
    extDir,
    {
      name: "config-change-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: {
        commands: [{ command: commandId, title: "Config Change" }],
        configuration: {
          title: "Config Change",
          properties: {
            [configKey]: { type: "string", default: "Hello" }
          }
        }
      },
      permissions: ["ui.commands", "storage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          const eventPromise = new Promise((resolve) => {
            const disp = formula.config.onDidChange((e) => {
              disp.dispose();
              resolve(e);
            });
          });

          await formula.config.update(${JSON.stringify(configKey)}, "Hi");
          const evt = await eventPromise;
          const value = await formula.config.get(${JSON.stringify(configKey)});
          return { evt, value };
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
  assert.deepEqual(result, {
    evt: { key: configKey, value: "Hi" },
    value: "Hi"
  });
});

test("api surface: workbook.saveAs updates workbook path and emits beforeSave", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-workbook-saveas-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "workbookExt.saveAs";
  await writeExtensionFixture(
    extDir,
    {
      name: "workbook-saveas-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Workbook SaveAs" }] },
      permissions: ["ui.commands", "workbook.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async (nextPath) => {
          const eventPromise = new Promise((resolve) => {
            const disp = formula.events.onBeforeSave((e) => {
              disp.dispose();
              resolve(e);
            });
          });

          await formula.workbook.saveAs(nextPath);
          const evt = await eventPromise;
          const workbook = await formula.workbook.getActiveWorkbook();
          return { evt, workbook };
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

  const initialPath = path.join(dir, "Initial.xlsx");
  host.openWorkbook(initialPath);

  const nextPath = path.join(dir, "Next.xlsx");
  const result = await host.executeCommand(commandId, nextPath);

  assert.deepEqual(result, {
    evt: { workbook: { name: "Next.xlsx", path: nextPath } },
    workbook: { name: "Next.xlsx", path: nextPath }
  });
});

test("api surface: workbook.close resets to default workbook and emits workbookOpened", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-workbook-close-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "workbookExt.close";
  await writeExtensionFixture(
    extDir,
    {
      name: "workbook-close-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Workbook Close" }] },
      permissions: ["ui.commands", "workbook.manage"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          const eventPromise = new Promise((resolve) => {
            const disp = formula.events.onWorkbookOpened((e) => {
              disp.dispose();
              resolve(e);
            });
          });

          await formula.workbook.close();
          const evt = await eventPromise;
          const workbook = await formula.workbook.getActiveWorkbook();
          return { evt, workbook };
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

  host.openWorkbook(path.join(dir, "Book3.xlsx"));

  const result = await host.executeCommand(commandId);
  assert.deepEqual(result, {
    evt: { workbook: { name: "MockWorkbook", path: null } },
    workbook: { name: "MockWorkbook", path: null }
  });
});

test("api surface: ui.showInputBox/showQuickPick return deterministic placeholder values", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-ui-prompts-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "uiExt.prompts";
  await writeExtensionFixture(
    extDir,
    {
      name: "ui-prompts-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "UI Prompts" }] },
      permissions: ["ui.commands"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          const input = await formula.ui.showInputBox({ prompt: "Name", value: "Alice" });
          const pick = await formula.ui.showQuickPick([
            { label: "One", value: 1 },
            { label: "Two", value: 2 }
          ]);
          return { input, pick };
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
  assert.deepEqual(result, { input: "Alice", pick: 1 });
});

test("api surface: ui.registerContextMenu adds and removes runtime menu items", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-ui-menus-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const registerCmd = "uiExt.registerMenu";
  const unregisterCmd = "uiExt.unregisterMenu";
  await writeExtensionFixture(
    extDir,
    {
      name: "ui-menus-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${registerCmd}`, `onCommand:${unregisterCmd}`],
      contributes: {
        commands: [
          { command: registerCmd, title: "Register Menu" },
          { command: unregisterCmd, title: "Unregister Menu" }
        ]
      },
      permissions: ["ui.commands", "ui.menus"]
    },
    `
      const formula = require("@formula/extension-api");
      let disposable = null;

      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          registerCmd
        )}, async () => {
          disposable = await formula.ui.registerContextMenu("cell/context", [
            { command: ${JSON.stringify(registerCmd)}, when: "cellHasValue", group: "extensions" }
          ]);
          return true;
        }));

        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          unregisterCmd
        )}, async () => {
          disposable?.dispose();
          disposable = null;
          return true;
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

  await host.executeCommand(registerCmd);
  assert.deepEqual(host.getContributedMenu("cell/context"), [
    {
      extensionId,
      command: registerCmd,
      when: "cellHasValue",
      group: "extensions"
    }
  ]);

  await host.executeCommand(unregisterCmd);
  const deadline = Date.now() + 500;
  while (Date.now() < deadline) {
    if (host.getContributedMenu("cell/context").length === 0) return;
    await new Promise((r) => setTimeout(r, 10));
  }
  assert.fail("Timed out waiting for context menu unregister");
});

test("permissions: ui.registerContextMenu requires ui.menus", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-ui-menus-deny-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "uiExt.registerDenied";
  await writeExtensionFixture(
    extDir,
    {
      name: "ui-menus-denied-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Register Menu Denied" }] },
      permissions: ["ui.commands", "ui.menus"]
    },
    `
      const formula = require("@formula/extension-api");
      exports.activate = async (context) => {
        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          await formula.ui.registerContextMenu("cell/context", [{ command: ${JSON.stringify(commandId)} }]);
          return true;
        }));
      };
    `
  );

  const host = new ExtensionHost({
    engineVersion: "1.0.0",
    permissionsStoragePath: path.join(dir, "permissions.json"),
    extensionStoragePath: path.join(dir, "storage.json"),
    permissionPrompt: async ({ permissions }) => !permissions.includes("ui.menus")
  });

  t.after(async () => {
    await host.dispose();
  });

  await host.loadExtension(extDir);
  await assert.rejects(() => host.executeCommand(commandId), /Permission denied: ui\.menus/);
});

test("api surface: formula.context and activation context expose storage paths", async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "formula-ext-context-"));
  const extDir = path.join(dir, "ext");
  await fs.mkdir(extDir);

  const commandId = "contextExt.read";
  await writeExtensionFixture(
    extDir,
    {
      name: "context-ext",
      version: "1.0.0",
      publisher: "formula-test",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      activationEvents: [`onCommand:${commandId}`],
      contributes: { commands: [{ command: commandId, title: "Context Read" }] },
      permissions: ["ui.commands"]
    },
    `
      const formula = require("@formula/extension-api");
      let ctx = null;

      exports.activate = async (context) => {
        ctx = {
          extensionId: context.extensionId,
          extensionPath: context.extensionPath,
          extensionUri: context.extensionUri,
          globalStoragePath: context.globalStoragePath,
          workspaceStoragePath: context.workspaceStoragePath
        };

        context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(
          commandId
        )}, async () => {
          return {
            ctx,
            api: {
              extensionId: formula.context.extensionId,
              extensionPath: formula.context.extensionPath,
              extensionUri: formula.context.extensionUri,
              globalStoragePath: formula.context.globalStoragePath,
              workspaceStoragePath: formula.context.workspaceStoragePath
            }
          };
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

  const expected = {
    extensionId,
    extensionPath: extDir,
    extensionUri: pathToFileURL(extDir).href,
    globalStoragePath: path.join(dir, "extension-data", extensionId, "globalStorage"),
    workspaceStoragePath: path.join(dir, "extension-data", extensionId, "workspaceStorage")
  };

  const result = await host.executeCommand(commandId);
  assert.deepEqual(result.ctx, expected);
  assert.deepEqual(result.api, expected);

  await assert.doesNotReject(() => fs.stat(expected.globalStoragePath));
  await assert.doesNotReject(() => fs.stat(expected.workspaceStoragePath));
});
