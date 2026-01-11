import assert from "node:assert/strict";
import test from "node:test";
import { createRequire } from "node:module";
  
import * as Y from "yjs";
  
import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";
import { maskCellValue } from "../packages/collab/permissions/index.js";
 
async function waitForCondition(predicate, timeoutMs = 2_000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (predicate()) return;
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error("Timed out waiting for condition");
}
 
async function waitForCell(documentController, sheetId, coord, expected) {
  await waitForCondition(() => {
    const cell = documentController.getCell(sheetId, coord);
    return (cell.value ?? null) === (expected.value ?? null) && (cell.formula ?? null) === (expected.formula ?? null);
  });
}
 
test("binder: masks unreadable cells and blocks disallowed edits", async () => {
  const ydoc = new Y.Doc();
  const cells = ydoc.getMap("cells");
 
  ydoc.transact(() => {
    const a1 = new Y.Map();
    a1.set("value", "allowed");
    cells.set("Sheet1:0:0", a1);
 
    const b1 = new Y.Map();
    b1.set("value", "secret");
    cells.set("Sheet1:0:1", b1);
  });
 
  const documentController = new DocumentController();
 
  /** @type {any[] | null} */
  let rejected = null;
 
  const binder = bindYjsToDocumentController({
    ydoc,
    documentController,
    defaultSheetId: "Sheet1",
    permissions: (cell) => {
      if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 1) {
        return { canRead: false, canEdit: false };
      }
      return { canRead: true, canEdit: true };
    },
    onEditRejected: (deltas) => {
      rejected = deltas;
    },
  });
 
  try {
    await waitForCell(documentController, "Sheet1", "A1", { value: "allowed", formula: null });
    await waitForCell(documentController, "Sheet1", "B1", { value: maskCellValue("secret"), formula: null });
 
    // Remote updates to unreadable cells should stay masked.
    const remoteOrigin = { type: "remote-test" };
    ydoc.transact(() => {
      const b1 = cells.get("Sheet1:0:1");
      assert.ok(b1 instanceof Y.Map);
      b1.set("value", "new-secret");
    }, remoteOrigin);
 
    await waitForCell(documentController, "Sheet1", "B1", { value: maskCellValue("new-secret"), formula: null });
 
    // Simulate a buggy caller that bypasses DocumentController.canEditCell by using applyExternalDeltas.
    // The binder must still reject + revert (and refuse to write into Yjs).
    const before = documentController.getCell("Sheet1", "B1");
    documentController.applyExternalDeltas([
      {
        sheetId: "Sheet1",
        row: 0,
        col: 1,
        before,
        after: { value: "hacked", formula: null, styleId: before.styleId },
      },
    ], { recalc: false });
 
    assert.ok(Array.isArray(rejected), "expected onEditRejected to be called");
    assert.equal(rejected?.length, 1);
    assert.equal(rejected?.[0]?.row, 0);
    assert.equal(rejected?.[0]?.col, 1);
 
    assert.equal(cells.get("Sheet1:0:1")?.get("value"), "new-secret");
    await waitForCell(documentController, "Sheet1", "B1", { value: maskCellValue("new-secret"), formula: null });
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});
 
test("binder: encrypted Yjs cells are masked and refuse plaintext writes", async () => {
  const ydoc = new Y.Doc();
  const cells = ydoc.getMap("cells");

  ydoc.transact(() => {
    const a1 = new Y.Map();
    a1.set("enc", { alg: "test", ciphertext: "deadbeef" });
    cells.set("Sheet1:0:0", a1);
  });

  const documentController = new DocumentController();

  /** @type {any[] | null} */
  let rejected = null;

  const binder = bindYjsToDocumentController({
    ydoc,
    documentController,
    defaultSheetId: "Sheet1",
    permissions: () => ({ canRead: true, canEdit: true }),
    onEditRejected: (deltas) => {
      rejected = deltas;
    },
  });

  try {
    await waitForCell(documentController, "Sheet1", "A1", { value: maskCellValue(null), formula: null });
 
    documentController.setCellValue("Sheet1", "A1", "hacked");
 
    assert.ok(Array.isArray(rejected), "expected onEditRejected to be called");
    const yCell = cells.get("Sheet1:0:0");
    assert.ok(yCell instanceof Y.Map);
    assert.ok(yCell.get("enc"), "expected encrypted payload to remain in Yjs");
    assert.equal(yCell.get("value"), undefined);
    assert.equal(yCell.get("formula"), undefined);

    assert.equal(documentController.getCell("Sheet1", "A1").value, maskCellValue(null));
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});

test("binder: initializes when cells root was created by a different Yjs instance (CJS Doc.getMap)", async () => {
  const require = createRequire(import.meta.url);
  // eslint-disable-next-line import/no-named-as-default-member
  const Ycjs = require("yjs");

  const ydoc = new Y.Doc();
  const cells = Ycjs.Doc.prototype.getMap.call(ydoc, "cells");

  // Populate a cell via the foreign Yjs instance to ensure the binder can read it.
  Ycjs.Doc.prototype.transact.call(ydoc, () => {
    const a1 = new Ycjs.Map();
    a1.set("value", "hello");
    a1.set("formula", null);
    cells.set("Sheet1:0:0", a1);
  });

  const documentController = new DocumentController();
  const binder = bindYjsToDocumentController({
    ydoc,
    documentController,
    defaultSheetId: "Sheet1",
  });

  try {
    await waitForCell(documentController, "Sheet1", "A1", { value: "hello", formula: null });
  } finally {
    binder.destroy();
    ydoc.destroy();
  }
});
