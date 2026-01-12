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
  extension.active = true;

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

test("BrowserExtensionHost: workbook.createWorkbook cancellation does not emit workbookOpened", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let createCalls = 0;
  let active = { name: "Initial", path: "/tmp/initial.xlsx" };
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

  /** @type {any} */
  let apiError;
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
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
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
      async createWorkbook() {
        createCalls += 1;
        throw { name: "AbortError", message: "Create workbook cancelled" };
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  // Ensure host has a stable baseline to restore to if create fails.
  await host._getActiveWorkbook();

  const extensionId = "test.workbook-create-cancel";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-create-cancel/",
    mainUrl: "memory://workbook-create-cancel/main.mjs",
    manifest: {
      name: "workbook-create-cancel",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "createWorkbook",
    args: [],
  });

  await done;

  assert.equal(createCalls, 1);
  assert.equal(apiError?.name, "AbortError");
  assert.equal(apiError?.message, "Create workbook cancelled");
  assert.equal(workbookOpened, undefined);
  assert.equal(host._workbook.path, "/tmp/initial.xlsx");
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
  extension.active = true;

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

test("BrowserExtensionHost: workbook.saveAs permission denial does not call spreadsheetApi or emit beforeSave", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let saveAsCalls = 0;
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

  /** @type {any} */
  let beforeSave;
  /** @type {any} */
  let apiError;

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
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
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
      async saveWorkbookAs() {
        saveAsCalls += 1;
      },
    },
    permissionPrompt: async () => false,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-saveas-permission-denied";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-saveas-permission-denied/",
    mainUrl: "memory://workbook-saveas-permission-denied/main.mjs",
    manifest: {
      name: "workbook-saveas-permission-denied",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "saveAs",
    args: ["/tmp/test.xlsx"],
  });

  await done;

  assert.equal(saveAsCalls, 0);
  assert.equal(beforeSave, undefined);
  assert.equal(apiError?.name, "PermissionError");
  assert.equal(apiError?.message, "Permission denied: workbook.manage");
});

test("BrowserExtensionHost: workbook.saveAs rejects when workbook.manage is not declared in manifest", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let promptCalls = 0;
  let saveAsCalls = 0;
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

  /** @type {any} */
  let beforeSave;
  /** @type {any} */
  let apiError;

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
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
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
      async saveWorkbookAs() {
        saveAsCalls += 1;
      },
    },
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-saveas-undeclared";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-saveas-undeclared/",
    mainUrl: "memory://workbook-saveas-undeclared/main.mjs",
    manifest: {
      name: "workbook-saveas-undeclared",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "saveAs",
    args: ["/tmp/test.xlsx"],
  });

  await done;

  assert.equal(promptCalls, 0);
  assert.equal(saveAsCalls, 0);
  assert.equal(beforeSave, undefined);
  assert.equal(apiError?.name, "PermissionError");
  assert.equal(apiError?.message, "Permission not declared in manifest: workbook.manage");
});

test("BrowserExtensionHost: workbook.saveAs rejects whitespace-only paths", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let promptCalls = 0;

  /** @type {any} */
  let apiError;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
          resolveDone();
        }
      },
    },
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-saveas-whitespace";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-saveas-whitespace/",
    mainUrl: "memory://workbook-saveas-whitespace/main.mjs",
    manifest: {
      name: "workbook-saveas-whitespace",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "saveAs",
    args: ["   "],
  });

  await done;

  assert.equal(promptCalls, 0);
  assert.equal(apiError?.message, "Workbook path must be a non-empty string");
});

test("BrowserExtensionHost: workbook.saveAs rejects empty paths", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let promptCalls = 0;

  /** @type {any} */
  let apiError;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
          resolveDone();
        }
      },
    },
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-saveas-empty";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-saveas-empty/",
    mainUrl: "memory://workbook-saveas-empty/main.mjs",
    manifest: {
      name: "workbook-saveas-empty",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "saveAs",
    args: [""],
  });

  await done;

  assert.equal(promptCalls, 0);
  assert.equal(apiError?.message, "Workbook path must be a non-empty string");
});

test("BrowserExtensionHost: workbook.saveAs rejects missing path arguments", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let promptCalls = 0;

  /** @type {any} */
  let apiError;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
          resolveDone();
        }
      },
    },
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-saveas-missing";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-saveas-missing/",
    mainUrl: "memory://workbook-saveas-missing/main.mjs",
    manifest: {
      name: "workbook-saveas-missing",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "saveAs",
    args: [],
  });

  await done;

  assert.equal(promptCalls, 0);
  assert.equal(apiError?.message, "Workbook path must be a non-empty string");
});

test("BrowserExtensionHost: workbook.saveAs rejects non-string paths", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let promptCalls = 0;

  /** @type {any} */
  let apiError;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
          resolveDone();
        }
      },
    },
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-saveas-nonstring";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-saveas-nonstring/",
    mainUrl: "memory://workbook-saveas-nonstring/main.mjs",
    manifest: {
      name: "workbook-saveas-nonstring",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "saveAs",
    args: [123],
  });

  await done;

  assert.equal(promptCalls, 0);
  assert.equal(apiError?.message, "Workbook path must be a non-empty string");
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
  extension.active = true;

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

test("BrowserExtensionHost: workbook.openWorkbook permission denial does not call spreadsheetApi or emit workbookOpened", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let openCalls = 0;
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];
  const active = { name: "Initial", path: "/tmp/initial.xlsx" };

  /** @type {any} */
  let apiError;
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
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
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
      async openWorkbook() {
        openCalls += 1;
      },
    },
    permissionPrompt: async () => false,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-open-permission-denied";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-open-permission-denied/",
    mainUrl: "memory://workbook-open-permission-denied/main.mjs",
    manifest: {
      name: "workbook-open-permission-denied",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: ["/tmp/opened.xlsx"],
  });

  await done;

  assert.equal(openCalls, 0);
  assert.equal(workbookOpened, undefined);
  assert.equal(apiError?.name, "PermissionError");
  assert.equal(apiError?.message, "Permission denied: workbook.manage");
});

test("BrowserExtensionHost: workbook.openWorkbook rejects when workbook.manage is not declared in manifest", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let promptCalls = 0;
  let openCalls = 0;
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];
  const active = { name: "Initial", path: "/tmp/initial.xlsx" };

  /** @type {any} */
  let apiError;
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
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
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
      async openWorkbook() {
        openCalls += 1;
      },
    },
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    },
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-open-undeclared";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-open-undeclared/",
    mainUrl: "memory://workbook-open-undeclared/main.mjs",
    manifest: {
      name: "workbook-open-undeclared",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: ["/tmp/opened.xlsx"],
  });

  await done;

  assert.equal(promptCalls, 0);
  assert.equal(openCalls, 0);
  assert.equal(workbookOpened, undefined);
  assert.equal(apiError?.name, "PermissionError");
  assert.equal(apiError?.message, "Permission not declared in manifest: workbook.manage");
});

test("BrowserExtensionHost: workbook.openWorkbook cancellation does not emit workbookOpened", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let openCalls = 0;
  let active = { name: "Initial", path: "/tmp/initial.xlsx" };
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

  /** @type {any} */
  let apiError;
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
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
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
      async openWorkbook() {
        openCalls += 1;
        throw { name: "AbortError", message: "Open workbook cancelled" };
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  // Ensure host has a stable baseline to restore to if open fails.
  await host._getActiveWorkbook();

  const extensionId = "test.workbook-open-cancel";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-open-cancel/",
    mainUrl: "memory://workbook-open-cancel/main.mjs",
    manifest: {
      name: "workbook-open-cancel",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: ["/tmp/opened.xlsx"],
  });

  await done;

  assert.equal(openCalls, 1);
  assert.equal(apiError?.name, "AbortError");
  assert.equal(apiError?.message, "Open workbook cancelled");
  assert.equal(workbookOpened, undefined);
  assert.equal(host._workbook.path, "/tmp/initial.xlsx");
});

test("BrowserExtensionHost: workbook.openWorkbook rejects empty paths", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let promptCalls = 0;

  /** @type {any} */
  let apiError;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
          resolveDone();
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-open-empty";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-open-empty/",
    mainUrl: "memory://workbook-open-empty/main.mjs",
    manifest: {
      name: "workbook-open-empty",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["workbook.manage"]
    }
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: [""]
  });

  await done;

  assert.equal(promptCalls, 0);
  assert.equal(apiError?.message, "Workbook path must be a non-empty string");
});

test("BrowserExtensionHost: workbook.openWorkbook rejects missing path arguments", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let promptCalls = 0;

  /** @type {any} */
  let apiError;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
          resolveDone();
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-open-missing";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-open-missing/",
    mainUrl: "memory://workbook-open-missing/main.mjs",
    manifest: {
      name: "workbook-open-missing",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["workbook.manage"]
    }
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: []
  });

  await done;

  assert.equal(promptCalls, 0);
  assert.equal(apiError?.message, "Workbook path must be a non-empty string");
});

test("BrowserExtensionHost: workbook.openWorkbook rejects whitespace-only paths", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let promptCalls = 0;

  /** @type {any} */
  let apiError;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
          resolveDone();
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-open-whitespace";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-open-whitespace/",
    mainUrl: "memory://workbook-open-whitespace/main.mjs",
    manifest: {
      name: "workbook-open-whitespace",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["workbook.manage"]
    }
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: ["   "]
  });

  await done;

  assert.equal(promptCalls, 0);
  assert.equal(apiError?.message, "Workbook path must be a non-empty string");
});

test("BrowserExtensionHost: workbook.openWorkbook rejects non-string paths", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let promptCalls = 0;

  /** @type {any} */
  let apiError;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
          resolveDone();
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {},
    permissionPrompt: async () => {
      promptCalls += 1;
      return true;
    }
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-open-nonstring";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-open-nonstring/",
    mainUrl: "memory://workbook-open-nonstring/main.mjs",
    manifest: {
      name: "workbook-open-nonstring",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["workbook.manage"]
    }
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: [123]
  });

  await done;

  assert.equal(promptCalls, 0);
  assert.equal(apiError?.message, "Workbook path must be a non-empty string");
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
  extension.active = true;

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

test("BrowserExtensionHost: workbook.close cancellation does not emit workbookOpened", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let closeCalls = 0;
  let active = { name: "Initial", path: "/tmp/initial.xlsx" };
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

  /** @type {any} */
  let apiError;
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
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
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
      async closeWorkbook() {
        closeCalls += 1;
        throw { name: "AbortError", message: "Close workbook cancelled" };
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  // Ensure host has a stable baseline to restore to if close fails.
  await host._getActiveWorkbook();

  const extensionId = "test.workbook-close-cancel";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-close-cancel/",
    mainUrl: "memory://workbook-close-cancel/main.mjs",
    manifest: {
      name: "workbook-close-cancel",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "close",
    args: [],
  });

  await done;

  assert.equal(closeCalls, 1);
  assert.equal(apiError?.name, "AbortError");
  assert.equal(apiError?.message, "Close workbook cancelled");
  assert.equal(workbookOpened, undefined);
  assert.equal(host._workbook.path, "/tmp/initial.xlsx");
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
  extension.active = true;

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

test("BrowserExtensionHost: workbook.save cancellation does not emit beforeSave when workbook has no path", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let saveCalls = 0;
  const active = { name: "Untitled", path: null };
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

  /** @type {any} */
  let beforeSave;
  /** @type {any} */
  let apiError;

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
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
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
        throw { name: "AbortError", message: "Save cancelled" };
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-save-cancel-pathless";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-save-cancel-pathless/",
    mainUrl: "memory://workbook-save-cancel-pathless/main.mjs",
    manifest: {
      name: "workbook-save-cancel-pathless",
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
  extension.active = true;

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
  assert.equal(apiError?.name, "AbortError");
  assert.equal(apiError?.message, "Save cancelled");
});

test("BrowserExtensionHost: workbook.save updates workbook snapshot when a pathless workbook is saved", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  let saveCalls = 0;
  let active = { name: "Untitled", path: null };
  const sheets = [{ id: "sheet1", name: "Sheet1" }];
  const activeSheet = sheets[0];

  /** @type {any} */
  let beforeSave;
  /** @type {any} */
  let apiResult;

  /** @type {(value?: unknown) => void} */
  let resolveSaveDone;
  const saveDone = new Promise((resolve) => {
    resolveSaveDone = resolve;
  });

  /** @type {(value?: unknown) => void} */
  let resolveGetDone;
  const getDone = new Promise((resolve) => {
    resolveGetDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "event" && msg.event === "beforeSave") {
          beforeSave = msg.data;
          return;
        }
        if (msg?.type === "api_result" && msg.id === "req1") {
          resolveSaveDone();
          return;
        }
        if (msg?.type === "api_result" && msg.id === "req2") {
          apiResult = msg.result;
          resolveGetDone();
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
        active = { name: "Saved.xlsx", path: "/tmp/saved.xlsx" };
      },
    },
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-save-pathless-updates";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-save-pathless-updates/",
    mainUrl: "memory://workbook-save-pathless-updates/main.mjs",
    manifest: {
      name: "workbook-save-pathless-updates",
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
  extension.active = true;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "save",
    args: [],
  });

  await saveDone;

  extension.worker.emitMessage({
    type: "api_call",
    id: "req2",
    namespace: "workbook",
    method: "getActiveWorkbook",
    args: [],
  });

  await getDone;

  assert.equal(saveCalls, 1);
  assert.equal(beforeSave, undefined);
  assert.deepEqual(apiResult, {
    name: "Saved.xlsx",
    path: "/tmp/saved.xlsx",
    sheets,
    activeSheet,
  });
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

test("BrowserExtensionHost: workbookOpened sent via _sendEventToExtension updates the fallback workbook snapshot", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

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
    spreadsheetApi: {},
    permissionPrompt: async () => true,
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-opened-targeted";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-opened-targeted/",
    mainUrl: "memory://workbook-opened-targeted/main.mjs",
    manifest: {
      name: "workbook-opened-targeted",
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

  host._sendEventToExtension(extension, "workbookOpened", {
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

test("BrowserExtensionHost: api_error preserves name/code for non-Error throws", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

  /** @type {any} */
  let apiError;

  /** @type {(value?: unknown) => void} */
  let resolveDone;
  const done = new Promise((resolve) => {
    resolveDone = resolve;
  });

  const scenarios = [
    {
      onPostMessage(msg) {
        if (msg?.type === "api_error" && msg.id === "req1") {
          apiError = msg.error;
          resolveDone();
        }
      }
    }
  ];

  installFakeWorker(t, scenarios);

  const host = new BrowserExtensionHost({
    engineVersion: "1.0.0",
    spreadsheetApi: {
      async getActiveWorkbook() {
        return { name: "Initial", path: null };
      },
      async openWorkbook() {
        // Simulate a DOMException-like value (not an Error instance) that still has
        // useful `name`/`code` metadata.
        throw { name: "AbortError", message: "Open workbook cancelled", code: { reason: "cancelled" } };
      }
    },
    permissionPrompt: async () => true
  });

  t.after(async () => {
    await host.dispose();
  });

  const extensionId = "test.workbook-error-serialization";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-error-serialization/",
    mainUrl: "memory://workbook-error-serialization/main.mjs",
    manifest: {
      name: "workbook-error-serialization",
      publisher: "test",
      version: "1.0.0",
      engines: { formula: "^1.0.0" },
      contributes: { commands: [], customFunctions: [] },
      activationEvents: [],
      permissions: ["workbook.manage"]
    }
  });

  const extension = host._extensions.get(extensionId);
  assert.ok(extension?.worker);

  extension.worker.emitMessage({
    type: "api_call",
    id: "req1",
    namespace: "workbook",
    method: "openWorkbook",
    args: ["/tmp/Example.xlsx"]
  });

  await done;

  assert.equal(apiError?.message, "Open workbook cancelled");
  assert.equal(apiError?.name, "AbortError");
  assert.equal(typeof apiError?.code, "string");
});

test("BrowserExtensionHost: getActiveWorkbook derives name from path when the host omits workbook.name", async (t) => {
  const { BrowserExtensionHost } = await importBrowserHost();

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
        return { path: "/tmp/opened.xlsx" };
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

  const extensionId = "test.workbook-name-from-path";
  await host.loadExtension({
    extensionId,
    extensionPath: "memory://workbook-name-from-path/",
    mainUrl: "memory://workbook-name-from-path/main.mjs",
    manifest: {
      name: "workbook-name-from-path",
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
