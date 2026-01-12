const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const { installFakeWorker } = require("./helpers/fake-browser-worker");

async function importBrowserHost() {
  const moduleUrl = pathToFileURL(path.resolve(__dirname, "../src/browser/index.mjs")).href;
  return import(moduleUrl);
}

test("BrowserExtensionHost: workbook.createWorkbook delegates to spreadsheetApi and emits workbookOpened", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let createCalls = 0;
  let active = { name: "Initial", path: "/tmp/initial.xlsx" };
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

  /** @type {any} */
  let apiResult;
  /** @type {any} */
  let workbookOpened;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_result" && msg.id === "req1") {
          apiResult = msg.result;
          resolveDone();
          return;
        }
        if (msg?.type === "event" && msg.event === "workbookOpened") {
          workbookOpened = msg.data;
        }
      },
    },
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async getActiveWorkbook() {
        return active;
      },
      listSheets() {
        return sheets;
      },
      getActiveSheet() {
        return activeSheet;
      },
      async createWorkbook() {
        createCalls += 1;
        active = { name: "Created", path: null };
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-delegate";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-delegate/",
    mainUrl: "memory://workbook-delegate/main.mjs",
    manifest: {
      name: "workbook-delegate",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["workbook.manage"],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "createWorkbook",
    args: [],
  });

  await done;

  assert.equal(createCalls, 1);
  assert.deepEqual(apiResult, {
    name: "Created",
    path: null,
    sheets,
    activeSheet,
  });

  assert.deepEqual(workbookOpened, {
    workbook: {
      name: "Created",
      path: null,
      sheets,
      activeSheet,
    },
  });
});

test("BrowserExtensionHost: workbook.saveAs delegates to spreadsheetApi and emits beforeSave with updated path", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let saveAsCalls = 0;
  /** @type {string | null} */
  let saveAsPath = null;
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

  /** @type {any} */
  let beforeSave;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "event" && msg.event === "beforeSave") {
          beforeSave = msg.data;
          return;
        }
        if (msg?.type === "api_result" && msg.id === "req1") {
          resolveDone();
          return;
        }
      },
    },
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      listSheets() {
        return sheets;
      },
      getActiveSheet() {
        return activeSheet;
      },
      async saveWorkbookAs(pathArg) {
        saveAsCalls += 1;
        saveAsPath = String(pathArg);
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-saveas";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-saveas/",
    mainUrl: "memory://workbook-saveas/main.mjs",
    manifest: {
      name: "workbook-saveas",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["workbook.manage"],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "saveAs",
    args: ["/tmp/test.xlsx"],
  });

  await done;

  assert.equal(saveAsCalls, 1);
  assert.equal(saveAsPath, "/tmp/test.xlsx");

  assert.deepEqual(beforeSave, {
    workbook: {
      name: "test.xlsx",
      path: "/tmp/test.xlsx",
      sheets,
      activeSheet,
    },
  });
});

