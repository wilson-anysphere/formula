import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";
import { requireYjsCjs } from "../packages/collab/yjs-utils/test/require-yjs-cjs.js";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";
import { createUndoService } from "../packages/collab/undo/index.js";

async function waitForCondition(predicate, timeoutMs = 2_000, intervalMs = 5) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (predicate()) return;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for condition");
}

function findSheetEntry(ydoc, sheetId) {
  const sheets = ydoc.getArray("sheets");
  for (const entry of sheets.toArray()) {
    const id = entry?.get?.("id") ?? entry?.id;
    if (id === sheetId) return entry;
  }
  return null;
}

test("binder: DocumentController→Yjs syncs sheet view state (freeze panes + row/col sizes)", async () => {
  const ydoc = new Y.Doc();
  const documentController = new DocumentController();

  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    documentController.setFrozen("Sheet1", 2, 3);
    documentController.setColWidth("Sheet1", 0, 120);
    documentController.setRowHeight("Sheet1", 5, 40);

    await waitForCondition(() => {
      const entry = findSheetEntry(ydoc, "Sheet1");
      const view = entry?.get?.("view") ?? entry?.view;
      return (
        view?.frozenRows === 2 &&
        view?.frozenCols === 3 &&
        view?.colWidths?.["0"] === 120 &&
        view?.rowHeights?.["5"] === 40
      );
    });

    const entry = findSheetEntry(ydoc, "Sheet1");
    assert.ok(entry, "expected a Yjs sheets entry for Sheet1");
    assert.deepEqual(entry.get?.("view") ?? entry.view, {
      frozenRows: 2,
      frozenCols: 3,
      colWidths: { "0": 120 },
      rowHeights: { "5": 40 },
    });

    // Backwards compatibility: mirror frozen panes at the sheet root (legacy schema).
    assert.equal(entry.get?.("frozenRows") ?? entry.frozenRows, 2);
    assert.equal(entry.get?.("frozenCols") ?? entry.frozenCols, 3);

    // Axis overrides live under `sheet.view.*` and should not be written at the root.
    assert.equal(entry.get?.("colWidths") ?? entry.colWidths, undefined);
    assert.equal(entry.get?.("rowHeights") ?? entry.rowHeights, undefined);
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: DocumentController→Yjs upgrades existing plain-object sheet entries when writing view state", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");
  ydoc.transact(() => {
    sheets.push([
      {
        id: "Sheet1",
        name: "Sheet1",
        // Include extra view metadata so we can verify the binder preserves unknown keys
        // when upgrading a plain-object sheet entry into a Y.Map.
        view: { frozenRows: 0, frozenCols: 0, drawings: [{ id: "drawing-1" }], defaultFormat: { font: { bold: true } } },
      },
    ]);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    documentController.setFrozen("Sheet1", 2, 0);

    await waitForCondition(() => {
      if (sheets.length !== 1) return false;
      const entry = sheets.get(0);
      if (!entry || typeof entry.get !== "function") return false;
      const view = entry.get("view");
      return entry.get("id") === "Sheet1" && view?.frozenRows === 2 && view?.frozenCols === 0;
    });

    assert.equal(sheets.length, 1);
    const entry = sheets.get(0);
    assert.ok(entry instanceof Y.Map, "expected sheet entry to be upgraded to a Y.Map");
    assert.deepEqual(entry.get("view"), {
      frozenRows: 2,
      frozenCols: 0,
      defaultFormat: { font: { bold: true } },
      drawings: [{ id: "drawing-1" }],
    });
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: preserves unknown view keys when writing sheet view deltas", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");
  ydoc.transact(() => {
    const entry = new Y.Map();
    entry.set("id", "Sheet1");
    entry.set("name", "Sheet1");
    entry.set("view", {
      frozenRows: 0,
      frozenCols: 0,
      defaultFormat: { font: { bold: true } },
      rowFormats: { "0": { numberFormat: "percent" } },
      colFormats: { "0": { numberFormat: "currency" } },
    });
    sheets.push([entry]);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    documentController.setFrozen("Sheet1", 2, 1);

    await waitForCondition(() => {
      const entry = sheets.get(0);
      if (!(entry instanceof Y.Map)) return false;
      const view = entry.get("view");
      return view?.frozenRows === 2 && view?.frozenCols === 1 && view?.defaultFormat?.font?.bold === true;
    });

    const entry = sheets.get(0);
    assert.ok(entry instanceof Y.Map);
    assert.deepEqual(entry.get("view"), {
      frozenRows: 2,
      frozenCols: 1,
      defaultFormat: { font: { bold: true } },
      rowFormats: { "0": { numberFormat: "percent" } },
      colFormats: { "0": { numberFormat: "currency" } },
    });
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: prefers non-local sheet entries when duplicates exist", async () => {
  const ydoc = new Y.Doc();

  // Simulate a remote client inserting a Sheet1 entry, then the local client
  // also inserting a duplicate Sheet1 entry (a common race during schema init).
  const remoteDoc = new Y.Doc();
  remoteDoc.transact(() => {
    const remoteSheets = remoteDoc.getArray("sheets");
    const entry = new Y.Map();
    entry.set("id", "Sheet1");
    entry.set("name", "Sheet1");
    remoteSheets.push([entry]);
  });

  Y.applyUpdate(ydoc, Y.encodeStateAsUpdate(remoteDoc));

  ydoc.transact(() => {
    const localSheets = ydoc.getArray("sheets");
    const entry = new Y.Map();
    entry.set("id", "Sheet1");
    entry.set("name", "Sheet1");
    localSheets.push([entry]);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    documentController.setFrozen("Sheet1", 1, 0);

    await waitForCondition(() => {
      const sheets = ydoc.getArray("sheets");
      for (const entry of sheets.toArray()) {
        if (!(entry instanceof Y.Map)) continue;
        if (entry.get("id") !== "Sheet1") continue;
        const view = entry.get("view");
        if (view?.frozenRows === 1 && view?.frozenCols === 0) return true;
      }
      return false;
    });

    const sheets = ydoc.getArray("sheets");
    assert.ok(sheets.length >= 2);

    /** @type {Y.Map<any>[]} */
    const entries = [];
    for (const entry of sheets.toArray()) {
      if (entry instanceof Y.Map && entry.get("id") === "Sheet1") entries.push(entry);
    }
    assert.ok(entries.length >= 2, "expected duplicate Sheet1 entries");

    const localClient = ydoc.clientID;
    const localEntries = entries.filter((e) => e?._item?.id?.client === localClient);
    const nonLocalEntries = entries.filter((e) => e?._item?.id?.client !== localClient);
    assert.ok(localEntries.length > 0, "expected at least one local duplicate entry");
    assert.ok(nonLocalEntries.length > 0, "expected at least one non-local duplicate entry");

    // The binder should apply view state to all duplicates so whichever entry
    // survives schema normalization retains the view settings.
    for (const entry of entries) {
      assert.deepEqual(entry.get("view"), { frozenRows: 1, frozenCols: 0 });
    }
  } finally {
    binder.destroy();
    ydoc.destroy();
    remoteDoc.destroy();
  }
});

test("binder: hydrates sheet view state from plain-object Yjs sheet entries", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");
  ydoc.transact(() => {
    sheets.push([{ id: "Sheet1", name: "Sheet1", view: { frozenRows: 1, frozenCols: 2, colWidths: { "0": 111 } } }]);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 1 && view.frozenCols === 2 && view.colWidths?.["0"] === 111;
    });

    assert.deepEqual(documentController.getSheetView("Sheet1"), {
      frozenRows: 1,
      frozenCols: 2,
      colWidths: { "0": 111 },
    });
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: hydrates sheet view state from legacy top-level frozenRows/frozenCols fields", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");
  ydoc.transact(() => {
    const entry = new Y.Map();
    entry.set("id", "Sheet1");
    entry.set("name", "Sheet1");
    // Legacy format: view fields stored at the top-level.
    entry.set("frozenRows", 2);
    entry.set("frozenCols", 1);
    sheets.push([entry]);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 2 && view.frozenCols === 1;
    });

    assert.deepEqual(documentController.getSheetView("Sheet1"), { frozenRows: 2, frozenCols: 1 });
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: hydrates sheet view state from sheets created by a different Yjs instance (CJS applyUpdate)", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.transact(() => {
    const sheets = remote.getArray("sheets");
    const sheet = new Ycjs.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const view = new Ycjs.Map();
    view.set("frozenRows", 1);
    view.set("frozenCols", 2);
    const colWidths = new Ycjs.Map();
    colWidths.set("0", 111);
    view.set("colWidths", colWidths);
    sheet.set("view", view);

    sheets.push([sheet]);
  });

  const ydoc = new Y.Doc();
  // Apply update via the CJS build to simulate y-websocket/provider behavior.
  const update = Ycjs.encodeStateAsUpdate(remote);
  Ycjs.applyUpdate(ydoc, update);

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 1 && view.frozenCols === 2 && view.colWidths?.["0"] === 111;
    });

    assert.deepEqual(documentController.getSheetView("Sheet1"), {
      frozenRows: 1,
      frozenCols: 2,
      colWidths: { "0": 111 },
    });

    // Remote collaborator updates nested map.
    remote.transact(() => {
      const sheets = remote.getArray("sheets");
      const sheet = sheets.get(0);
      const view = sheet.get("view");
      const colWidths = view.get("colWidths");
      colWidths.set("1", 222);
    });

    const remoteOrigin = { type: "remote-test" };
    Ycjs.applyUpdate(ydoc, Ycjs.encodeStateAsUpdate(remote), remoteOrigin);

    await waitForCondition(() => documentController.getSheetView("Sheet1").colWidths?.["1"] === 222);
    assert.deepEqual(documentController.getSheetView("Sheet1"), {
      frozenRows: 1,
      frozenCols: 2,
      colWidths: { "0": 111, "1": 222 },
    });
  } finally {
    binder.destroy();
    ydoc.destroy();
    remote.destroy();
  }
});

test("binder: collab undo/redo reverts local sheet view changes when sheet.view is a foreign Y.Map (CJS applyUpdate)", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.transact(() => {
    const sheets = remote.getArray("sheets");
    const sheet = new Ycjs.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const view = new Ycjs.Map();
    view.set("frozenRows", 1);
    view.set("frozenCols", 0);
    sheet.set("view", view);

    sheets.push([sheet]);
  });

  const ydoc = new Y.Doc();
  // Ensure the root exists in the ESM Yjs instance so the update only introduces
  // foreign nested sheet maps (not a foreign `sheets` root).
  ydoc.getArray("sheets");
  Ycjs.applyUpdate(ydoc, Ycjs.encodeStateAsUpdate(remote));

  const sheets = ydoc.getArray("sheets");
  const undo = createUndoService({ mode: "collab", doc: ydoc, scope: sheets });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({
    ydoc,
    documentController,
    undoService: undo,
    defaultSheetId: "Sheet1",
  });

  try {
    // Initial hydration should apply the remote frozen state.
    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 1 && view.frozenCols === 0;
    });

    const entry = findSheetEntry(ydoc, "Sheet1");
    assert.ok(entry, "expected Sheet1 entry");
    const rawView = entry.get("view");
    assert.ok(rawView && typeof rawView === "object");
    assert.equal(rawView instanceof Y.Map, false, "expected sheet.view to be a foreign (non-ESM) Y.Map");
    assert.equal(typeof rawView.get, "function", "expected sheet.view to behave like a Y.Map");

    // Local edit.
    documentController.setFrozen("Sheet1", 2, 1);
    await waitForCondition(() => documentController.getSheetView("Sheet1").frozenRows === 2);

    // Undo should revert to the initial remote state.
    undo.undo();
    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 1 && view.frozenCols === 0;
    });
  } finally {
    binder.destroy();
    ydoc.destroy();
    remote.destroy();
  }
});

test("binder: collab undo/redo reverts local col width changes when sheet.view.colWidths is a foreign Y.Map (CJS applyUpdate)", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.transact(() => {
    const sheets = remote.getArray("sheets");
    const sheet = new Ycjs.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const view = new Ycjs.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);

    const colWidths = new Ycjs.Map();
    colWidths.set("0", 111);
    view.set("colWidths", colWidths);
    sheet.set("view", view);

    sheets.push([sheet]);
  });

  const ydoc = new Y.Doc();
  // Ensure the root exists in the ESM Yjs instance so the update only introduces
  // foreign nested sheet/view maps (not a foreign `sheets` root).
  ydoc.getArray("sheets");
  Ycjs.applyUpdate(ydoc, Ycjs.encodeStateAsUpdate(remote));

  const sheets = ydoc.getArray("sheets");
  const undo = createUndoService({ mode: "collab", doc: ydoc, scope: sheets });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({
    ydoc,
    documentController,
    undoService: undo,
    defaultSheetId: "Sheet1",
  });

  try {
    await waitForCondition(() => documentController.getSheetView("Sheet1").colWidths?.["0"] === 111);

    const entry = findSheetEntry(ydoc, "Sheet1");
    assert.ok(entry, "expected Sheet1 entry");
    const rawView = entry.get("view");
    assert.ok(rawView && typeof rawView === "object");
    assert.equal(rawView instanceof Y.Map, false, "expected sheet.view to be a foreign (non-ESM) Y.Map");
    const rawColWidths = rawView?.get?.("colWidths");
    assert.ok(rawColWidths && typeof rawColWidths === "object");
    assert.equal(rawColWidths instanceof Y.Map, false, "expected view.colWidths to be a foreign (non-ESM) Y.Map");

    // Local edit (overwrite col 0).
    documentController.setColWidth("Sheet1", 0, 222);

    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.colWidths?.["0"] === 222;
    });

    // Ensure the binder wrote the edit into Yjs.
    await waitForCondition(() => {
      const entry = findSheetEntry(ydoc, "Sheet1");
      const view = entry?.get?.("view");
      const colWidths = view?.get?.("colWidths");
      const width = Number(colWidths?.get?.("0"));
      return width === 222;
    });

    undo.undo();

    await waitForCondition(() => documentController.getSheetView("Sheet1").colWidths?.["0"] === 111);
    assert.deepEqual(documentController.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0, colWidths: { "0": 111 } });
  } finally {
    binder.destroy();
    ydoc.destroy();
    remote.destroy();
  }
});

test("binder: collab undo/redo reverts local row height changes when sheet.view.rowHeights is a foreign Y.Map (CJS applyUpdate)", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.transact(() => {
    const sheets = remote.getArray("sheets");
    const sheet = new Ycjs.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const view = new Ycjs.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);

    const rowHeights = new Ycjs.Map();
    rowHeights.set("10", 44);
    view.set("rowHeights", rowHeights);
    sheet.set("view", view);

    sheets.push([sheet]);
  });

  const ydoc = new Y.Doc();
  ydoc.getArray("sheets");
  Ycjs.applyUpdate(ydoc, Ycjs.encodeStateAsUpdate(remote));

  const sheets = ydoc.getArray("sheets");
  const undo = createUndoService({ mode: "collab", doc: ydoc, scope: sheets });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({
    ydoc,
    documentController,
    undoService: undo,
    defaultSheetId: "Sheet1",
  });

  try {
    await waitForCondition(() => documentController.getSheetView("Sheet1").rowHeights?.["10"] === 44);

    const entry = findSheetEntry(ydoc, "Sheet1");
    assert.ok(entry, "expected Sheet1 entry");
    const rawView = entry.get("view");
    assert.ok(rawView && typeof rawView === "object");
    assert.equal(rawView instanceof Y.Map, false, "expected sheet.view to be a foreign (non-ESM) Y.Map");
    const rawRowHeights = rawView?.get?.("rowHeights");
    assert.ok(rawRowHeights && typeof rawRowHeights === "object");
    assert.equal(rawRowHeights instanceof Y.Map, false, "expected view.rowHeights to be a foreign (non-ESM) Y.Map");

    documentController.setRowHeight("Sheet1", 10, 55);

    await waitForCondition(() => documentController.getSheetView("Sheet1").rowHeights?.["10"] === 55);

    await waitForCondition(() => {
      const entry = findSheetEntry(ydoc, "Sheet1");
      const view = entry?.get?.("view");
      const rowHeights = view?.get?.("rowHeights");
      const height = Number(rowHeights?.get?.("10"));
      return height === 55;
    });

    undo.undo();

    await waitForCondition(() => documentController.getSheetView("Sheet1").rowHeights?.["10"] === 44);
    assert.deepEqual(documentController.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0, rowHeights: { "10": 44 } });
  } finally {
    binder.destroy();
    ydoc.destroy();
    remote.destroy();
  }
});

test("binder: collab undo/redo reverts local background image changes when sheet.view is a foreign Y.Map (CJS applyUpdate)", async () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.transact(() => {
    const sheets = remote.getArray("sheets");
    const sheet = new Ycjs.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const view = new Ycjs.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);
    view.set("backgroundImageId", "bg1.png");
    sheet.set("view", view);

    sheets.push([sheet]);
  });

  const ydoc = new Y.Doc();
  ydoc.getArray("sheets");
  Ycjs.applyUpdate(ydoc, Ycjs.encodeStateAsUpdate(remote));

  const sheets = ydoc.getArray("sheets");
  const undo = createUndoService({ mode: "collab", doc: ydoc, scope: sheets });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({
    ydoc,
    documentController,
    undoService: undo,
    defaultSheetId: "Sheet1",
  });

  try {
    await waitForCondition(() => documentController.getSheetView("Sheet1").backgroundImageId === "bg1.png");

    const entry = findSheetEntry(ydoc, "Sheet1");
    assert.ok(entry, "expected Sheet1 entry");
    const rawView = entry.get("view");
    assert.ok(rawView && typeof rawView === "object");
    assert.equal(rawView instanceof Y.Map, false, "expected sheet.view to be a foreign (non-ESM) Y.Map");

    documentController.setSheetBackgroundImageId("Sheet1", "bg2.png");

    await waitForCondition(() => documentController.getSheetView("Sheet1").backgroundImageId === "bg2.png");

    await waitForCondition(() => {
      const entry = findSheetEntry(ydoc, "Sheet1");
      const view = entry?.get?.("view");
      const id = view?.get?.("backgroundImageId");
      return id === "bg2.png";
    });

    undo.undo();

    await waitForCondition(() => documentController.getSheetView("Sheet1").backgroundImageId === "bg1.png");
    assert.deepEqual(documentController.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0, backgroundImageId: "bg1.png" });
  } finally {
    binder.destroy();
    ydoc.destroy();
    remote.destroy();
  }
});

test("binder: applies view state when sheet id is set after view", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");

  ydoc.transact(() => {
    const entry = new Y.Map();
    entry.set("name", "Sheet1");
    entry.set("view", { frozenRows: 2, frozenCols: 1 });
    sheets.push([entry]);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    // No id yet, so there is no sheet key to apply against.
    assert.deepEqual(documentController.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0 });

    const remoteOrigin = { type: "remote-test" };
    ydoc.transact(
      () => {
        const entry = sheets.get(0);
        assert.ok(entry instanceof Y.Map);
        entry.set("id", "Sheet1");
      },
      remoteOrigin,
    );

    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 2 && view.frozenCols === 1;
    });

    assert.deepEqual(documentController.getSheetView("Sheet1"), { frozenRows: 2, frozenCols: 1 });
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: observes deep view mutations when view is stored as nested Y.Maps", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");

  const view = new Y.Map();
  view.set("frozenRows", 0);
  view.set("frozenCols", 0);
  const colWidths = new Y.Map();
  colWidths.set("0", 111);
  view.set("colWidths", colWidths);

  ydoc.transact(() => {
    const entry = new Y.Map();
    entry.set("id", "Sheet1");
    entry.set("name", "Sheet1");
    entry.set("view", view);
    sheets.push([entry]);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    await waitForCondition(() => documentController.getSheetView("Sheet1").colWidths?.["0"] === 111);

    const remoteOrigin = { type: "remote-test" };
    ydoc.transact(
      () => {
        colWidths.set("1", 222);
      },
      remoteOrigin,
    );

    await waitForCondition(() => documentController.getSheetView("Sheet1").colWidths?.["1"] === 222);
    assert.deepEqual(documentController.getSheetView("Sheet1"), {
      frozenRows: 0,
      frozenCols: 0,
      colWidths: { "0": 111, "1": 222 },
    });
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: collab undo/redo reverts sheet view changes (sheets in undo scope)", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");
  ydoc.transact(() => {
    const entry = new Y.Map();
    entry.set("id", "Sheet1");
    entry.set("name", "Sheet1");
    sheets.push([entry]);
  });

  const undo = createUndoService({ mode: "collab", doc: ydoc, scope: sheets });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({
    ydoc,
    documentController,
    undoService: undo,
    defaultSheetId: "Sheet1",
  });

  try {
    documentController.setFrozen("Sheet1", 2, 1);

    // Wait until the binder has applied the local view write into Yjs so the
    // UndoManager has something to undo.
    await waitForCondition(() => {
      const entry = sheets.get(0);
      if (!(entry instanceof Y.Map)) return false;
      const view = entry.get("view");
      return view?.frozenRows === 2 && view?.frozenCols === 1;
    });

    undo.stopCapturing();
    assert.equal(undo.canUndo(), true);

    undo.undo();
    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 0 && view.frozenCols === 0;
    });
    assert.deepEqual(documentController.getSheetView("Sheet1"), { frozenRows: 0, frozenCols: 0 });

    assert.equal(undo.canRedo(), true);
    undo.redo();
    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 2 && view.frozenCols === 1;
    });
    assert.deepEqual(documentController.getSheetView("Sheet1"), { frozenRows: 2, frozenCols: 1 });
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: Yjs→DocumentController syncs sheet view state (initial hydration + remote changes)", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");

  ydoc.transact(() => {
    const entry = new Y.Map();
    entry.set("id", "Sheet1");
    entry.set("name", "Sheet1");
    entry.set("view", { frozenRows: 1, frozenCols: 1, colWidths: { "0": 111 } });
    sheets.push([entry]);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    // Remote hydration should not create undo history entries.
    assert.deepEqual(documentController.getStackDepths(), { undo: 0, redo: 0 });
    assert.equal(documentController.canUndo, false);

    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 1 && view.frozenCols === 1 && view.colWidths?.["0"] === 111;
    });

    assert.deepEqual(documentController.getSheetView("Sheet1"), {
      frozenRows: 1,
      frozenCols: 1,
      colWidths: { "0": 111 },
    });
    assert.deepEqual(documentController.getStackDepths(), { undo: 0, redo: 0 });
    assert.equal(documentController.canUndo, false);

    // Simulate a remote collaborator updating the view object.
    const remoteOrigin = { type: "remote-test" };
    ydoc.transact(
      () => {
        const entry = findSheetEntry(ydoc, "Sheet1");
        assert.ok(entry, "expected Sheet1 entry");
        if (typeof entry.set === "function") {
          entry.set("view", { frozenRows: 0, frozenCols: 2, rowHeights: { "10": 55 } });
        } else {
          const sheets = ydoc.getArray("sheets");
          const idx = sheets.toArray().findIndex((e) => (e?.get?.("id") ?? e?.id) === "Sheet1");
          sheets.delete(idx, 1);
          sheets.insert(idx, [{ id: "Sheet1", name: "Sheet1", view: { frozenRows: 0, frozenCols: 2, rowHeights: { "10": 55 } } }]);
        }
      },
      remoteOrigin,
    );

    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 0 && view.frozenCols === 2 && view.rowHeights?.["10"] === 55 && !view.colWidths;
    });

    assert.deepEqual(documentController.getSheetView("Sheet1"), {
      frozenRows: 0,
      frozenCols: 2,
      rowHeights: { "10": 55 },
    });
    assert.deepEqual(documentController.getStackDepths(), { undo: 0, redo: 0 });
    assert.equal(documentController.canUndo, false);
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: local sheet view changes do not echo back as external changes", async () => {
  const ydoc = new Y.Doc();
  const documentController = new DocumentController();

  /** @type {any[]} */
  const changeEvents = [];
  const unsubscribe = documentController.on("change", (payload) => changeEvents.push(payload));

  const binder = bindYjsToDocumentController({ ydoc, documentController, defaultSheetId: "Sheet1" });

  try {
    // Give hydration a chance to run (it should be a no-op for empty docs).
    await new Promise((r) => setTimeout(r, 25));
    changeEvents.length = 0;

    documentController.setFrozen("Sheet1", 1, 1);

    await waitForCondition(() => {
      const entry = findSheetEntry(ydoc, "Sheet1");
      const view = entry?.get("view");
      return view?.frozenRows === 1 && view?.frozenCols === 1;
    });

    // Allow any potential echo to schedule.
    await new Promise((r) => setTimeout(r, 25));

    const collabEvents = changeEvents.filter((evt) => evt?.source === "collab");
    assert.equal(collabEvents.length, 0);
  } finally {
    unsubscribe();
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: syncs layered full-column formatting via sheet metadata (no per-cell materialization)", async () => {
  const ydoc = new Y.Doc();
  const documentControllerA = new DocumentController();
  const documentControllerB = new DocumentController();

  const binderA = bindYjsToDocumentController({ ydoc, documentController: documentControllerA, defaultSheetId: "Sheet1" });
  const binderB = bindYjsToDocumentController({ ydoc, documentController: documentControllerB, defaultSheetId: "Sheet1" });

  try {
    // Apply formatting to the full height of column A. This should update the column layer,
    // not materialize 1M cells in the sparse `cells` map.
    documentControllerA.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

    await waitForCondition(() => {
      const entry = findSheetEntry(ydoc, "Sheet1");
      const colFormats = entry?.get?.("colFormats") ?? entry?.colFormats;
      const col0 = colFormats?.get?.("0") ?? colFormats?.["0"];
      return col0?.font?.bold === true;
    });

    await waitForCondition(() => documentControllerB.getCellFormat("Sheet1", "A1")?.font?.bold === true);

    assert.equal(documentControllerB.getCell("Sheet1", "A1").styleId, 0);
    assert.equal(documentControllerB.getCellFormat("Sheet1", "A1")?.font?.bold, true);
  } finally {
    binderA.destroy();
    binderB.destroy();
    ydoc.destroy();
  }
});

test("binder: syncs range-run formatting via sheet metadata (no per-cell materialization)", async () => {
  const ydoc = new Y.Doc();
  const documentControllerA = new DocumentController();
  const documentControllerB = new DocumentController();

  const binderA = bindYjsToDocumentController({ ydoc, documentController: documentControllerA, defaultSheetId: "Sheet1" });
  const binderB = bindYjsToDocumentController({ ydoc, documentController: documentControllerB, defaultSheetId: "Sheet1" });

  try {
    // Use a large range that exceeds the range-run threshold so DocumentController stores
    // the patch in `formatRunsByCol` instead of per-cell styles.
    documentControllerA.setRangeFormat("Sheet1", "A1:A50001", { font: { italic: true } });

    await waitForCondition(() => {
      const entry = findSheetEntry(ydoc, "Sheet1");
      const runsByCol = entry?.get?.("formatRunsByCol") ?? entry?.formatRunsByCol;
      const col0 = runsByCol?.get?.("0") ?? runsByCol?.["0"];
      return Array.isArray(col0) && col0[0]?.format?.font?.italic === true;
    });

    await waitForCondition(() => documentControllerB.getCellFormat("Sheet1", "A1")?.font?.italic === true);

    assert.equal(documentControllerB.getCell("Sheet1", "A1").styleId, 0);
    assert.equal(documentControllerB.getCellFormat("Sheet1", "A1")?.font?.italic, true);
    assert.equal(documentControllerB.getCellFormat("Sheet1", "A50002")?.font?.italic, undefined);
  } finally {
    binderA.destroy();
    binderB.destroy();
    ydoc.destroy();
  }
});

test("binder: clearing top-level formats overrides legacy view-based formats", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");
  ydoc.transact(() => {
    const entry = new Y.Map();
    entry.set("id", "Sheet1");
    entry.set("name", "Sheet1");
    // Legacy encoding: layered formats stored inside `view` (older BranchService docs).
    entry.set("view", { frozenRows: 0, frozenCols: 0, colFormats: { "0": { font: { bold: true } } } });
    sheets.push([entry]);
  });

  const documentControllerA = new DocumentController();
  const documentControllerB = new DocumentController();

  const binderA = bindYjsToDocumentController({ ydoc, documentController: documentControllerA, defaultSheetId: "Sheet1" });
  const binderB = bindYjsToDocumentController({ ydoc, documentController: documentControllerB, defaultSheetId: "Sheet1" });

  try {
    await waitForCondition(() => documentControllerA.getCellFormat("Sheet1", "A1")?.font?.bold === true);
    await waitForCondition(() => documentControllerB.getCellFormat("Sheet1", "A1")?.font?.bold === true);

    // Clear the formatting. Binder should write an explicit empty top-level `colFormats`
    // map so we don't fall back to the stale `view.colFormats` value.
    documentControllerA.setRangeFormat("Sheet1", "A1:A1048576", null);

    await waitForCondition(() => {
      const entry = findSheetEntry(ydoc, "Sheet1");
      const topLevel = entry?.get?.("colFormats");
      if (!topLevel || typeof topLevel.get !== "function") return false;
      return topLevel.get("0") == null;
    });

    await waitForCondition(() => documentControllerB.getCellFormat("Sheet1", "A1")?.font?.bold !== true);
    assert.equal(documentControllerB.getCell("Sheet1", "A1").styleId, 0);
  } finally {
    binderA.destroy();
    binderB.destroy();
    ydoc.destroy();
  }
});
