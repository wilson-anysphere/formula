const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const cjsApi = require("..");

const esmPath = pathToFileURL(path.join(__dirname, "..", "index.mjs")).href;

function stripApiCall(message) {
  return {
    type: message.type,
    namespace: message.namespace,
    method: message.method,
    args: message.args
  };
}

function createMockHost() {
  /** @type {any} */
  let activeWorkbook = {
    name: "Book1",
    path: "/tmp/book1.xlsx",
    sheets: [{ id: "sheet-1", name: "Sheet1" }],
    activeSheet: { id: "sheet-1", name: "Sheet1" }
  };

  /** @type {any} */
  let activeSheet = { id: "sheet-1", name: "Sheet1" };

  /** @param {any} message */
  function handleCall(message) {
    const key = `${message?.namespace}.${message?.method}`;
    const args = Array.isArray(message?.args) ? message.args : [];

    switch (key) {
      case "cells.getSelection":
      case "cells.getRange":
        return {
          startRow: 0,
          startCol: 0,
          endRow: 1,
          endCol: 1,
          values: [
            [1, 2],
            [3, 4]
          ]
        };
      case "cells.getCell":
        return 42;
      case "cells.setRange":
      case "sheets.renameSheet":
      case "workbook.save":
      case "workbook.close":
      case "workbook.saveAs":
        if (key === "workbook.close") {
          activeWorkbook = { ...activeWorkbook, name: "BookClosed", path: null };
        }
        if (key === "workbook.save") {
          // Model Save prompting for a path (Save As) when the workbook is currently unsaved.
          if (!activeWorkbook?.path) {
            activeWorkbook = { ...activeWorkbook, name: "BookSavedFromSave", path: "/tmp/book-saved-from-save.xlsx" };
          }
        }
        if (key === "workbook.saveAs") {
          const workbookPath = String(args[0]);
          activeWorkbook = { ...activeWorkbook, name: "BookSaved", path: workbookPath };
        }
        if (key === "sheets.renameSheet") {
          const from = String(args[0]);
          const to = String(args[1]);
          if (activeSheet?.name === from) activeSheet = { ...activeSheet, name: to };
          activeWorkbook = {
            ...activeWorkbook,
            sheets: Array.isArray(activeWorkbook?.sheets)
              ? activeWorkbook.sheets.map((s) => (s?.name === from ? { ...s, name: to } : s))
              : activeWorkbook.sheets,
            activeSheet:
              activeWorkbook?.activeSheet?.name === from
                ? { ...activeWorkbook.activeSheet, name: to }
                : activeWorkbook.activeSheet
          };
        }
        return null;
      case "sheets.getActiveSheet":
        return activeSheet;
      case "sheets.activateSheet": {
        const name = String(args[0]);
        activeSheet = { id: `sheet-active:${name}`, name };
        activeWorkbook = { ...activeWorkbook, activeSheet };
        return activeSheet;
      }
      case "sheets.getSheet": {
        const name = String(args[0]);
        if (name === "Missing") return null;
        return { id: `sheet:${name}`, name };
      }
      case "workbook.getActiveWorkbook":
        return activeWorkbook;
      default:
        throw new Error(`Unhandled api call in test transport: ${key}`);
    }
  }

  return { handleCall };
}

function setMockTransport(handleCall) {
  /** @type {any[]} */
  const calls = [];
  cjsApi.__setTransport({
    postMessage: (message) => {
      calls.push(message);
      if (message?.type !== "api_call") return;
      queueMicrotask(() => {
        try {
          const result = handleCall(message);
          cjsApi.__handleMessage({ type: "api_result", id: message.id, result });
        } catch (error) {
          cjsApi.__handleMessage({
            type: "api_error",
            id: message.id,
            error: { message: String(error?.message ?? error), stack: error?.stack }
          });
        }
      });
    }
  });
  return calls;
}

test("dual entrypoint: CJS + ESM stay in lockstep", async (t) => {
  const esmApi = await import(esmPath);

  await t.test("exports stay in sync", async () => {
    const cjsKeys = Object.keys(cjsApi).sort();
    const esmKeys = Object.keys(esmApi).filter((k) => k !== "default").sort();

    assert.deepEqual(
      esmKeys,
      cjsKeys,
      `ESM export surface drifted.\nCJS: ${cjsKeys.join(", ")}\nESM: ${esmKeys.join(", ")}`
    );

    const namespaces = [
      "workbook",
      "sheets",
      "cells",
      "commands",
      "functions",
      "dataConnectors",
      "network",
      "clipboard",
      "ui",
      "storage",
      "config",
      "events"
    ];

    for (const ns of namespaces) {
      assert.equal(typeof cjsApi[ns], "object", `Expected CJS ${ns} to be an object`);
      assert.equal(typeof esmApi[ns], "object", `Expected ESM ${ns} to be an object`);
      assert.deepEqual(
        Object.keys(esmApi[ns]).sort(),
        Object.keys(cjsApi[ns]).sort(),
        `Namespace ${ns} drifted between entrypoints`
      );
    }
  });

  await t.test("runtime behavior matches the .d.ts contract in both entrypoints", async () => {
    const mockHost = createMockHost();
    const calls = setMockTransport(mockHost.handleCall);

    // Range shaping (address + formulas) for both entrypoints.
    const rangeCjs = await cjsApi.cells.getRange("A1:B2");
    const rangeEsm = await esmApi.cells.getRange("A1:B2");

    for (const [label, range] of [
      ["CJS", rangeCjs],
      ["ESM", rangeEsm]
    ]) {
      assert.equal(range.address, "A1:B2", `${label} range.address should be A1 notation`);
      assert.deepEqual(
        range.formulas,
        [
          [null, null],
          [null, null]
        ],
        `${label} range.formulas should default to a null matrix`
      );
    }

    // Workbook shaping + helper methods.
    const workbookCjs = await cjsApi.workbook.getActiveWorkbook();
    const workbookEsm = await esmApi.workbook.getActiveWorkbook();

    for (const [label, wb] of [
      ["CJS", workbookCjs],
      ["ESM", workbookEsm]
    ]) {
      assert.equal(typeof wb.save, "function", `${label} Workbook.save should exist`);
      assert.equal(typeof wb.saveAs, "function", `${label} Workbook.saveAs should exist`);
      assert.equal(typeof wb.close, "function", `${label} Workbook.close should exist`);
      assert.equal(
        typeof wb.activeSheet?.getRange,
        "function",
        `${label} Workbook.activeSheet should be enhanced`
      );
    }

    // CJS workbook helpers should call through to RPC.
    calls.length = 0;
    await workbookCjs.save();
    assert.deepEqual(calls.filter((m) => m.type === "api_call").map(stripApiCall), [
      { type: "api_call", namespace: "workbook", method: "save", args: [] }
    ]);

    // When saving an unsaved workbook (pathless), the workbook instance should refresh fields.
    calls.length = 0;
    await workbookCjs.close();
    assert.equal(workbookCjs.path, null);

    calls.length = 0;
    await workbookCjs.save();
    assert.equal(workbookCjs.path, "/tmp/book-saved-from-save.xlsx");
    assert.equal(workbookCjs.name, "BookSavedFromSave");
    assert.deepEqual(calls.filter((m) => m.type === "api_call").map(stripApiCall), [
      { type: "api_call", namespace: "workbook", method: "save", args: [] },
      { type: "api_call", namespace: "workbook", method: "getActiveWorkbook", args: [] }
    ]);

    // ESM workbook helpers should call RPC and refresh local fields.
    calls.length = 0;
    await workbookEsm.saveAs("/tmp/book2.xlsx");
    assert.equal(workbookEsm.name, "BookSaved");
    assert.equal(workbookEsm.path, "/tmp/book2.xlsx");
    assert.deepEqual(
      calls.filter((m) => m.type === "api_call").map(stripApiCall),
      [
        { type: "api_call", namespace: "workbook", method: "saveAs", args: ["/tmp/book2.xlsx"] },
        { type: "api_call", namespace: "workbook", method: "getActiveWorkbook", args: [] }
      ]
    );

    // Workbook path validation should reject non-string/empty args without sending RPC calls.
    calls.length = 0;
    await assert.rejects(() => cjsApi.workbook.openWorkbook(123), {
      message: "Workbook path must be a non-empty string"
    });
    assert.equal(calls.filter((m) => m.type === "api_call").length, 0);

    calls.length = 0;
    await assert.rejects(() => esmApi.workbook.saveAs({}), {
      message: "Workbook path must be a non-empty string"
    });
    assert.equal(calls.filter((m) => m.type === "api_call").length, 0);

    calls.length = 0;
    await assert.rejects(() => workbookEsm.saveAs(123), {
      message: "Workbook path must be a non-empty string"
    });
    assert.equal(calls.filter((m) => m.type === "api_call").length, 0);

    // Sheet shaping + helper methods.
    const sheetCjs = await cjsApi.sheets.getActiveSheet();
    const sheetEsm = await esmApi.sheets.getActiveSheet();

    for (const [label, sheet] of [
      ["CJS", sheetCjs],
      ["ESM", sheetEsm]
    ]) {
      assert.equal(typeof sheet.getRange, "function", `${label} Sheet.getRange should exist`);
      assert.equal(typeof sheet.setRange, "function", `${label} Sheet.setRange should exist`);
      assert.equal(typeof sheet.activate, "function", `${label} Sheet.activate should exist`);
      assert.equal(typeof sheet.rename, "function", `${label} Sheet.rename should exist`);
    }

    // CJS sheet helpers should prefix refs with sheet name.
    calls.length = 0;
    await sheetCjs.setRange("A1", [[99]]);
    assert.deepEqual(calls.filter((m) => m.type === "api_call").map(stripApiCall), [
      { type: "api_call", namespace: "cells", method: "setRange", args: [`${sheetCjs.name}!A1`, [[99]]] }
    ]);

    calls.length = 0;
    const fromName = sheetEsm.name;
    await sheetEsm.rename("RenamedSheet");
    assert.equal(sheetEsm.name, "RenamedSheet");
    assert.deepEqual(calls.filter((m) => m.type === "api_call").map(stripApiCall), [
      { type: "api_call", namespace: "sheets", method: "renameSheet", args: [fromName, "RenamedSheet"] }
    ]);

    calls.length = 0;
    const range = await sheetEsm.getRange("A1");
    assert.equal(range.address, "A1:B2");
    assert.deepEqual(calls.filter((m) => m.type === "api_call").map(stripApiCall), [
      { type: "api_call", namespace: "cells", method: "getRange", args: ["RenamedSheet!A1"] }
    ]);

    // Event payload shaping (selection -> Range, workbook -> Workbook)
    const selectionEvent = await new Promise((resolve) => {
      const disposable = esmApi.events.onSelectionChanged((e) => {
        disposable.dispose();
        resolve(e);
      });
      cjsApi.__handleMessage({
        type: "event",
        event: "selectionChanged",
        data: {
          selection: {
            startRow: 0,
            startCol: 0,
            endRow: 0,
            endCol: 0,
            values: [[1]]
          }
        }
      });
    });

    assert.equal(selectionEvent.selection.address, "A1");
    assert.deepEqual(selectionEvent.selection.formulas, [[null]]);

    // When hosts omit `selection.formulas`, the runtime backfills it with a null matrix for
    // API consistency. That behavior must be capped so selecting very large ranges (eg: entire
    // Excel sheets) can't OOM the worker by allocating millions of entries.
    const hugeSelectionEvent = await new Promise((resolve) => {
      const disposable = esmApi.events.onSelectionChanged((e) => {
        disposable.dispose();
        resolve(e);
      });
      cjsApi.__handleMessage({
        type: "event",
        event: "selectionChanged",
        data: {
          selection: {
            startRow: 0,
            startCol: 0,
            endRow: 449,
            endCol: 449,
            values: []
          }
        }
      });
    });

    assert.deepEqual(hugeSelectionEvent.selection.formulas, []);
    assert.equal(hugeSelectionEvent.selection.truncated, true);

    const workbookOpenedEvent = await new Promise((resolve) => {
      const disposable = cjsApi.events.onWorkbookOpened((e) => {
        disposable.dispose();
        resolve(e);
      });

      cjsApi.__handleMessage({
        type: "event",
        event: "workbookOpened",
        data: {
          workbook: {
            name: "EventBook",
            path: "/tmp/event.xlsx",
            sheets: [{ id: "sheet-1", name: "Sheet1" }],
            activeSheet: { id: "sheet-1", name: "Sheet1" }
          }
        }
      });
    });

    assert.equal(typeof workbookOpenedEvent.workbook.save, "function");
    assert.equal(typeof workbookOpenedEvent.workbook.activeSheet.getRange, "function");

    // `sheets.getSheet()` should normalize falsy host responses to `undefined`.
    const missingSheet = await esmApi.sheets.getSheet("Missing");
    assert.equal(missingSheet, undefined);
  });

  await t.test("transport/context state is shared between CJS and ESM", async () => {
    /** @type {any} */
    let lastCall = null;
    cjsApi.__setTransport({
      postMessage: (message) => {
        lastCall = message;
        if (message?.type === "api_call") {
          queueMicrotask(() => {
            // Respond via the opposite entrypoint to ensure pending state is shared.
            cjsApi.__handleMessage({ type: "api_result", id: message.id, result: 42 });
          });
        }
      }
    });
    cjsApi.__setContext({
      extensionId: "formula.test",
      extensionPath: "/tmp/ext",
      extensionUri: "file:///tmp/ext",
      globalStoragePath: "/tmp/ext/globalStorage",
      workspaceStoragePath: "/tmp/ext/workspaceStorage"
    });

    assert.equal(esmApi.context.extensionId, "formula.test");
    assert.equal(esmApi.context.extensionPath, "/tmp/ext");
    assert.equal(esmApi.context.extensionUri, "file:///tmp/ext");
    assert.equal(esmApi.context.globalStoragePath, "/tmp/ext/globalStorage");
    assert.equal(esmApi.context.workspaceStoragePath, "/tmp/ext/workspaceStorage");

    const value = await esmApi.cells.getCell(1, 2);
    assert.equal(value, 42);

    assert.deepEqual(lastCall, {
      type: "api_call",
      id: lastCall.id,
      namespace: "cells",
      method: "getCell",
      args: [1, 2]
    });
  });
});
