import test from "node:test";
import assert from "node:assert/strict";

import { createCollabSession } from "../src/index.ts";

test("CollabSession writes formula=null markers when clearing formulas without conflict monitors", async () => {
  const session = createCollabSession();
  try {
    await session.setCellFormula("Sheet1:0:0", "=1");
    await session.setCellFormula("Sheet1:0:0", null);

    const cell = session.cells.get("Sheet1:0:0");
    assert.ok(cell, "expected Yjs cell map to exist");
    assert.equal(cell.get("formula"), null);
  } finally {
    session.destroy();
    session.doc.destroy();
  }
});

test("CollabSession.getCell treats formula=null the same as an absent formula key", async () => {
  const session = createCollabSession();
  try {
    await session.setCellFormula("Sheet1:0:0", "=1");
    await session.setCellValue("Sheet1:0:0", "literal");

    // With the explicit marker, `getCell` should still surface `formula: null`.
    const readWithMarker = await session.getCell("Sheet1:0:0");
    assert.ok(readWithMarker);
    assert.equal(readWithMarker.value, "literal");
    assert.equal(readWithMarker.formula, null);

    // Simulate a legacy writer that deletes the formula key entirely: read API
    // should behave the same.
    session.doc.transact(() => {
      const cell = session.cells.get("Sheet1:0:0");
      cell?.delete?.("formula");
    });

    const readWithDelete = await session.getCell("Sheet1:0:0");
    assert.ok(readWithDelete);
    assert.equal(readWithDelete.value, "literal");
    assert.equal(readWithDelete.formula, null);
  } finally {
    session.destroy();
    session.doc.destroy();
  }
});

test("CollabSession clears formulas via formula=null markers when writing literal values without conflict monitors", async () => {
  const session = createCollabSession();
  try {
    await session.setCellFormula("Sheet1:0:0", "=1");
    await session.setCellValue("Sheet1:0:0", "literal");

    const cell = session.cells.get("Sheet1:0:0");
    assert.ok(cell, "expected Yjs cell map to exist");
    assert.equal(cell.get("formula"), null);
    assert.equal(cell.get("value"), "literal");
  } finally {
    session.destroy();
    session.doc.destroy();
  }
});

test("CollabSession clears formulas via formula=null markers when using cellValueConflicts monitor", async () => {
  const session = createCollabSession({
    cellValueConflicts: {
      localUserId: "u",
      onConflict: () => {},
    },
  });
  try {
    await session.setCellFormula("Sheet1:0:0", "=1");
    await session.setCellValue("Sheet1:0:0", "literal");

    const cell = session.cells.get("Sheet1:0:0");
    assert.ok(cell, "expected Yjs cell map to exist");
    assert.equal(cell.get("formula"), null);
    assert.equal(cell.get("value"), "literal");
  } finally {
    session.destroy();
    session.doc.destroy();
  }
});
