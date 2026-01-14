import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabSession } from "../packages/collab/session/src/index.ts";
import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindSheetViewToCollabSession } from "../apps/desktop/src/collab/sheetViewBinder.ts";

const REMOTE_ORIGIN = Symbol("remote");

/**
 * @param {Y.Doc} docA
 * @param {Y.Doc} docB
 */
function connectDocs(docA, docB) {
  const forwardA = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docB, update, REMOTE_ORIGIN);
  };
  const forwardB = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docA, update, REMOTE_ORIGIN);
  };

  docA.on("update", forwardA);
  docB.on("update", forwardB);

  Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB), REMOTE_ORIGIN);
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), REMOTE_ORIGIN);

  return () => {
    docA.off("update", forwardA);
    docB.off("update", forwardB);
  };
}

async function waitForCondition(fn, timeoutMs = 2000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const ok = await fn();
      if (ok) return;
    } catch {
      // Ignore transient errors while waiting for async state to settle.
    }
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error("Timed out waiting for condition");
}

test("CollabSession sheet view binder syncs frozen panes + axis overrides without polluting undo history", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA });
  const sessionB = createCollabSession({ doc: docB });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = bindSheetViewToCollabSession({ session: sessionA, documentController: dcA });
  const binderB = bindSheetViewToCollabSession({ session: sessionB, documentController: dcB });

  dcA.setFrozen("Sheet1", 2, 1, { label: "Freeze" });
  dcA.setColWidth("Sheet1", 0, 120, { label: "Resize Column" });
  dcA.setRowHeight("Sheet1", 1, 40, { label: "Resize Row" });

  await waitForCondition(() => {
    const view = dcB.getSheetView("Sheet1");
    return (
      view.frozenRows === 2 &&
      view.frozenCols === 1 &&
      view.colWidths?.["0"] === 120 &&
      view.rowHeights?.["1"] === 40
    );
  });

  // Remote changes should not create local undo history entries.
  assert.deepEqual(dcB.getStackDepths(), { undo: 0, redo: 0 });

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession sheet view binder syncs drawings without polluting undo history", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA });
  const sessionB = createCollabSession({ doc: docB });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = bindSheetViewToCollabSession({ session: sessionA, documentController: dcA });
  const binderB = bindSheetViewToCollabSession({ session: sessionB, documentController: dcB });

  const drawings = [
    {
      id: "drawing-1",
      zOrder: 0,
      kind: { type: "image", imageId: "img-1" },
      anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
    },
  ];

  dcA.setSheetDrawings("Sheet1", drawings, { label: "Insert Picture" });

  await waitForCondition(() => {
    assert.deepEqual(dcB.getSheetDrawings("Sheet1"), drawings);
    return true;
  });

  // Remote changes should not create local undo history entries.
  assert.deepEqual(dcB.getStackDepths(), { undo: 0, redo: 0 });

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession sheet view binder does not write view state into Yjs when session role is read-only", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc });
  // Viewer/commenter roles must not mutate the shared doc (even if a caller applies local-only
  // view changes in the UI).
  session.setPermissions({ role: "viewer", rangeRestrictions: [], userId: "viewer-1" });

  const dc = new DocumentController();
  const binder = bindSheetViewToCollabSession({ session, documentController: dc });

  /** @type {number} */
  let updates = 0;
  const onUpdate = () => {
    updates += 1;
  };

  try {
    // Give initial schema/hydration work a chance to settle before we start counting updates.
    await new Promise((r) => setTimeout(r, 25));
    updates = 0;
    doc.on("update", onUpdate);

    // Local UI changes should update DocumentController but not persist into Yjs.
    dc.setFrozen("Sheet1", 2, 1, { label: "Freeze (local-only)" });
    dc.setColWidth("Sheet1", 0, 120, { label: "Resize Column (local-only)" });
    dc.setRowHeight("Sheet1", 1, 40, { label: "Resize Row (local-only)" });
    dc.setSheetDrawings(
      "Sheet1",
      [
        {
          id: "drawing-local",
          zOrder: 0,
          kind: { type: "image", imageId: "img-local" },
          anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
        },
      ],
      { label: "Insert Drawing (local-only)" },
    );

    await new Promise((r) => setTimeout(r, 25));
    assert.equal(updates, 0, "expected no Yjs updates from local sheet view changes in read-only role");

    assert.deepEqual(dc.getSheetView("Sheet1"), {
      frozenRows: 2,
      frozenCols: 1,
      colWidths: { "0": 120 },
      rowHeights: { "1": 40 },
    });
    assert.deepEqual(dc.getSheetDrawings("Sheet1"), [
      {
        id: "drawing-local",
        zOrder: 0,
        kind: { type: "image", imageId: "img-local" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
      },
    ]);

    // Remote Yjs updates should still apply to the DocumentController.
    updates = 0;
    doc.transact(
      () => {
        const sheet = session.sheets.toArray().find((s) => (s?.get?.("id") ?? s?.id) === "Sheet1") ?? null;
        assert.ok(sheet, "expected Sheet1 entry in Yjs");
        // Simulate a remote collaborator overwriting the view state.
        sheet.set("view", { frozenRows: 0, frozenCols: 3, colWidths: { "0": 200 }, rowHeights: { "1": 10 } });
      },
      REMOTE_ORIGIN,
    );

    await waitForCondition(() => {
      const view = dc.getSheetView("Sheet1");
      return (
        view.frozenRows === 0 &&
        view.frozenCols === 3 &&
        view.colWidths?.["0"] === 200 &&
        view.rowHeights?.["1"] === 10
      );
    });

    assert.ok(updates > 0, "expected remote Yjs changes to produce doc updates");
    assert.deepEqual(dc.getSheetView("Sheet1"), {
      frozenRows: 0,
      frozenCols: 3,
      colWidths: { "0": 200 },
      rowHeights: { "1": 10 },
    });

    // Remote drawings updates should still apply to the DocumentController in read-only mode.
    updates = 0;
    doc.transact(
      () => {
        const sheet = session.sheets.toArray().find((s) => (s?.get?.("id") ?? s?.id) === "Sheet1") ?? null;
        assert.ok(sheet, "expected Sheet1 entry in Yjs");
        const view = sheet.get("view");
        const drawings = [
          {
            id: "drawing-remote",
            zOrder: 1,
            kind: { type: "image", imageId: "img-remote" },
            anchor: { type: "absolute", pos: { xEmu: 10, yEmu: 10 }, size: { cx: 2, cy: 3 } },
          },
        ];
        if (view && typeof view.set === "function") {
          view.set("drawings", drawings);
        } else {
          sheet.set("view", { ...(view ?? {}), drawings });
        }
      },
      REMOTE_ORIGIN,
    );

    await waitForCondition(() => {
      assert.deepEqual(dc.getSheetDrawings("Sheet1"), [
        {
          id: "drawing-remote",
          zOrder: 1,
          kind: { type: "image", imageId: "img-remote" },
          anchor: { type: "absolute", pos: { xEmu: 10, yEmu: 10 }, size: { cx: 2, cy: 3 } },
        },
      ]);
      return true;
    });

    assert.ok(updates > 0, "expected remote Yjs drawing changes to produce doc updates");
  } finally {
    doc.off("update", onUpdate);
    binder.destroy();
    session.destroy();
    doc.destroy();
  }
});
