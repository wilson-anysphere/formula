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

test("BrowserExtensionHost: workbook.openWorkbook delegates to spreadsheetApi and emits workbookOpened", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let openCalls = 0;
  /** @type {string | null} */
  let openedPath = null;
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
        if (msg?.type === "event" && msg.event === "workbookOpened") {
          workbookOpened = msg.data;
          return;
        }
        if (msg?.type === "api_result" && msg.id === "req1") {
          apiResult = msg.result;
          resolveDone();
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
      async getActiveWorkbook() {
        return active;
      },
      async openWorkbook(pathArg) {
        openCalls += 1;
        openedPath = String(pathArg);
        active = { name: "opened.xlsx", path: openedPath };
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-open";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-open/",
    mainUrl: "memory://workbook-open/main.mjs",
    manifest: {
      name: "workbook-open",
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
    method: "openWorkbook",
    args: ["/tmp/opened.xlsx"],
  });

  await done;

  assert.equal(openCalls, 1);
  assert.equal(openedPath, "/tmp/opened.xlsx");

  assert.deepEqual(apiResult, {
    name: "opened.xlsx",
    path: "/tmp/opened.xlsx",
    sheets,
    activeSheet,
  });

  assert.deepEqual(workbookOpened, {
    workbook: {
      name: "opened.xlsx",
      path: "/tmp/opened.xlsx",
      sheets,
      activeSheet,
    },
  });
});

test("BrowserExtensionHost: workbook.close delegates to spreadsheetApi and emits workbookOpened", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let closeCalls = 0;
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

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
        if (msg?.type === "event" && msg.event === "workbookOpened") {
          workbookOpened = msg.data;
          return;
        }
        if (msg?.type === "api_result" && msg.id === "req1") {
          resolveDone();
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
      async closeWorkbook() {
        closeCalls += 1;
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-close";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-close/",
    mainUrl: "memory://workbook-close/main.mjs",
    manifest: {
      name: "workbook-close",
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
    method: "close",
    args: [],
  });

  await done;

  assert.equal(closeCalls, 1);

  assert.deepEqual(workbookOpened, {
    workbook: {
      name: "MockWorkbook",
      path: null,
      sheets,
      activeSheet,
    },
  });
});

test("BrowserExtensionHost: workbook.save delegates to spreadsheetApi and emits beforeSave", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let saveCalls = 0;
  const active = { name: "Book.xlsx", path: "/tmp/book.xlsx" };
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
      async getActiveWorkbook() {
        return active;
      },
      async saveWorkbook() {
        saveCalls += 1;
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-save";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-save/",
    mainUrl: "memory://workbook-save/main.mjs",
    manifest: {
      name: "workbook-save",
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
    method: "save",
    args: [],
  });

  await done;

  assert.equal(saveCalls, 1);
  assert.deepEqual(beforeSave, {
    workbook: {
      name: active.name,
      path: active.path,
      sheets,
      activeSheet,
    },
  });
});

test("BrowserExtensionHost: workbook.save does not emit beforeSave when workbook has no path", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let saveCalls = 0;
  const active = { name: "Untitled", path: null };
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
      async getActiveWorkbook() {
        return active;
      },
      async saveWorkbook() {
        saveCalls += 1;
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-save-pathless";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-save-pathless/",
    mainUrl: "memory://workbook-save-pathless/main.mjs",
    manifest: {
      name: "workbook-save-pathless",
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
    method: "save",
    args: [],
  });

  await done;

  assert.equal(saveCalls, 1);
  assert.equal(beforeSave, undefined);
});

test("BrowserExtensionHost: workbook.getActiveWorkbook overwrites stored path when spreadsheetApi returns path=null", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let active = { name: "Initial", path: "/tmp/initial.xlsx" };
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

  /** @type {any} */
  let apiResult;

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
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-clear-path";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-clear-path/",
    mainUrl: "memory://workbook-clear-path/main.mjs",
    manifest: {
      name: "workbook-clear-path",
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

  // Prime internal workbook metadata with a path, then simulate the host returning a
  // pathless workbook snapshot.
  host.openWorkbook("/tmp/primed.xlsx");
  active = { name: "Untitled", path: null };

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "getActiveWorkbook",
    args: [],
  });

  await done;

  assert.equal(apiResult?.path ?? null, null);
});

test("BrowserExtensionHost: workbookOpened events update the fallback workbook snapshot when getActiveWorkbook is not implemented", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  /** @type {((event: any) => void) | null} */
  let onWorkbookOpened = null;

  /** @type {any} */
  let apiResult;

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
        }
      },
    },
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      onWorkbookOpened(handler) {
        onWorkbookOpened = handler;
        return () => {
          onWorkbookOpened = null;
        };
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-opened-updates-fallback";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-opened-updates-fallback/",
    mainUrl: "memory://workbook-opened-updates-fallback/main.mjs",
    manifest: {
      name: "workbook-opened-updates-fallback",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: [],
    },
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);

  assert.equal(typeof onWorkbookOpened, "function");
  onWorkbookOpened({
    workbook: {
      name: "opened.xlsx",
      path: "/tmp/opened.xlsx",
      sheets: [{ id: "sheet1", name: "Sheet1" }],
      activeSheet: { id: "sheet1", name: "Sheet1" },
    },
  });

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "getActiveWorkbook",
    args: [],
  });

  await done;

  assert.equal(apiResult?.path, "/tmp/opened.xlsx");
  assert.equal(apiResult?.name, "opened.xlsx");
});

test("BrowserExtensionHost: dispose unsubscribes from spreadsheetApi event hooks", async () => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let selectionDisposed = false;
  let cellDisposed = false;
  let sheetDisposed = false;
  let workbookDisposed = false;
  let beforeSaveDisposed = false;

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      onSelectionChanged() {
        return () => {
          selectionDisposed = true;
        };
      },
      onCellChanged() {
        return { dispose: () => {
          cellDisposed = true;
        } };
      },
      onSheetActivated() {
        return () => {
          sheetDisposed = true;
        };
      },
      onWorkbookOpened() {
        return () => {
          workbookDisposed = true;
        };
      },
      onBeforeSave() {
        return () => {
          beforeSaveDisposed = true;
        };
      },
    },
    permissionPrompt: async () => true,
  });

  await host.dispose();

  assert.equal(selectionDisposed, true);
  assert.equal(cellDisposed, true);
  assert.equal(sheetDisposed, true);
  assert.equal(workbookDisposed, true);
  assert.equal(beforeSaveDisposed, true);
});
