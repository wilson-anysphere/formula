import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { createCollabSession } from "../src/index.ts";

test("CollabSession.compactCells deletes truly-empty cells", () => {
  const session = createCollabSession({ doc: new Y.Doc() });
  try {
    session.doc.transact(() => {
      const emptyMarker = new Y.Map();
      emptyMarker.set("value", null);
      emptyMarker.set("formula", null);
      emptyMarker.set("modified", 1);
      session.cells.set("Sheet1:0:0", emptyMarker);

      const blankValue = new Y.Map();
      blankValue.set("value", { t: "blank" });
      blankValue.set("formula", "");
      session.cells.set("Sheet1:0:1", blankValue);

      const nonEmpty = new Y.Map();
      nonEmpty.set("value", "x");
      nonEmpty.set("formula", null);
      session.cells.set("Sheet1:0:2", nonEmpty);
    });

    assert.equal(session.cells.size, 3);

    const result = session.compactCells();
    assert.deepEqual(result, { scanned: 3, deleted: 2 });

    assert.equal(session.cells.has("Sheet1:0:0"), false);
    assert.equal(session.cells.has("Sheet1:0:1"), false);
    assert.equal(session.cells.has("Sheet1:0:2"), true);
  } finally {
    session.destroy();
    session.doc.destroy();
  }
});

test("CollabSession.compactCells does not delete encrypted cells", () => {
  const session = createCollabSession({ doc: new Y.Doc() });
  try {
    session.doc.transact(() => {
      const encrypted = new Y.Map();
      encrypted.set("enc", { keyId: "k1", payload: "..." });
      encrypted.set("value", null);
      encrypted.set("formula", null);
      session.cells.set("Sheet1:1:0", encrypted);

      const empty = new Y.Map();
      empty.set("value", null);
      empty.set("formula", null);
      session.cells.set("Sheet1:1:1", empty);
    });

    const result = session.compactCells();
    assert.equal(result.deleted, 1);
    assert.equal(session.cells.has("Sheet1:1:0"), true);
    assert.equal(session.cells.has("Sheet1:1:1"), false);
  } finally {
    session.destroy();
    session.doc.destroy();
  }
});

test("CollabSession.compactCells does not delete format-only cells", () => {
  const session = createCollabSession({ doc: new Y.Doc() });
  try {
    session.doc.transact(() => {
      const formatted = new Y.Map();
      formatted.set("format", { bold: true });
      formatted.set("value", null);
      formatted.set("formula", null);
      session.cells.set("Sheet1:2:0", formatted);

      const empty = new Y.Map();
      empty.set("value", null);
      empty.set("formula", null);
      session.cells.set("Sheet1:2:1", empty);
    });

    const result = session.compactCells();
    assert.equal(result.deleted, 1);
    assert.equal(session.cells.has("Sheet1:2:0"), true);
    assert.equal(session.cells.has("Sheet1:2:1"), false);
  } finally {
    session.destroy();
    session.doc.destroy();
  }
});

test("CollabSession.compactCells is idempotent", () => {
  const session = createCollabSession({ doc: new Y.Doc() });
  try {
    session.doc.transact(() => {
      const empty = new Y.Map();
      empty.set("value", null);
      empty.set("formula", null);
      session.cells.set("Sheet1:3:0", empty);
    });

    const first = session.compactCells();
    assert.equal(first.deleted, 1);
    assert.equal(session.cells.size, 0);

    const second = session.compactCells();
    assert.equal(second.deleted, 0);
    assert.equal(session.cells.size, 0);
  } finally {
    session.destroy();
    session.doc.destroy();
  }
});

