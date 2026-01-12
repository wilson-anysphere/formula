import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";

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

    // Binder should not write legacy top-level frozen rows/cols.
    assert.equal(entry.get?.("frozenRows") ?? entry.frozenRows, undefined);
    assert.equal(entry.get?.("frozenCols") ?? entry.frozenCols, undefined);
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: DocumentController→Yjs upgrades existing plain-object sheet entries when writing view state", async () => {
  const ydoc = new Y.Doc();
  const sheets = ydoc.getArray("sheets");
  ydoc.transact(() => {
    sheets.push([{ id: "Sheet1", name: "Sheet1" }]);
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
    assert.deepEqual(entry.get("view"), { frozenRows: 2, frozenCols: 0 });
  } finally {
    binder.destroy();
    ydoc.destroy();
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
    await waitForCondition(() => {
      const view = documentController.getSheetView("Sheet1");
      return view.frozenRows === 1 && view.frozenCols === 1 && view.colWidths?.["0"] === 111;
    });

    assert.deepEqual(documentController.getSheetView("Sheet1"), {
      frozenRows: 1,
      frozenCols: 1,
      colWidths: { "0": 111 },
    });

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
