import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { DocumentController } from "../../../../apps/desktop/src/document/documentController.js";
import { bindCollabSessionToDocumentController, createCollabSession } from "../src/index.ts";

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

test("CollabSession↔DocumentController binder masks unreadable remote values/formulas", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA });
  const sessionB = createCollabSession({ doc: docB });

  sessionA.setPermissions({ role: "editor", userId: "u-a", rangeRestrictions: [] });
  sessionB.setPermissions({
    role: "editor",
    userId: "u-b",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 1,
        readAllowlist: ["u-a"],
        editAllowlist: [],
      },
    ],
  });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = await bindCollabSessionToDocumentController({ session: sessionA, documentController: dcA });
  const binderB = await bindCollabSessionToDocumentController({
    session: sessionB,
    documentController: dcB,
    maskCellFormat: true,
  });

  // Perform edits via DocumentController (typical UI path) so we exercise
  // DocumentController→Yjs propagation as well.
  dcA.setCellValue("Sheet1", "A1", "super secret");
  dcA.setRangeFormat("Sheet1", "A1", { font: { bold: true } });
  dcA.setCellFormula("Sheet1", "B1", "=TOP_SECRET()");
  dcA.setRangeFormat("Sheet1", "B1", { fill: { color: "#ff0000" } });

  await waitForCondition(() => {
    const a1 = sessionB.cells.get("Sheet1:0:0");
    const b1 = sessionB.cells.get("Sheet1:0:1");
    return (
      a1 &&
      typeof a1.get === "function" &&
      a1.get("format") != null &&
      b1 &&
      typeof b1.get === "function" &&
      b1.get("format") != null
    );
  });

  await waitForCondition(() => {
    const cellA = dcA.getCell("Sheet1", "A1");
    return cellA.value === "super secret" && cellA.formula == null && cellA.styleId !== 0;
  });

  await waitForCondition(() => {
    const cellB = dcB.getCell("Sheet1", "A1");
    return cellB.value === "###" && cellB.formula == null && cellB.styleId === 0;
  });

  await waitForCondition(() => {
    const cellA = dcA.getCell("Sheet1", "B1");
    return cellA.formula === "=TOP_SECRET()" && cellA.value == null && cellA.styleId !== 0;
  });

  await waitForCondition(() => {
    const cellB = dcB.getCell("Sheet1", "B1");
    return cellB.value === "###" && cellB.formula == null && cellB.styleId === 0;
  });

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession↔DocumentController binder blocks edits to non-editable cells", async () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA });
  const sessionB = createCollabSession({ doc: docB });

  sessionA.setPermissions({ role: "editor", userId: "u-a", rangeRestrictions: [] });
  sessionB.setPermissions({
    role: "editor",
    userId: "u-b",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        // Anyone can read, but only u-a can edit.
        readAllowlist: [],
        editAllowlist: ["u-a"],
      },
    ],
  });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = await bindCollabSessionToDocumentController({ session: sessionA, documentController: dcA });
  const binderB = await bindCollabSessionToDocumentController({ session: sessionB, documentController: dcB });

  // Seed a value as the editable user.
  dcA.setCellValue("Sheet1", "A1", "original");
  await waitForCondition(() => dcB.getCell("Sheet1", "A1").value === "original");
  await waitForCondition(async () => (await sessionA.getCell("Sheet1:0:0"))?.value === "original");

  // Attempt an edit as the restricted user.
  dcB.setCellValue("Sheet1", "A1", "hacked");
  dcB.setCellFormula("Sheet1", "A1", "=HACK()");

  // Local UI and shared Yjs document should remain unchanged.
  await new Promise((r) => setTimeout(r, 25));
  assert.equal(dcB.getCell("Sheet1", "A1").value, "original");
  assert.equal(dcB.getCell("Sheet1", "A1").formula, null);
  assert.equal(dcA.getCell("Sheet1", "A1").value, "original");
  assert.equal(dcA.getCell("Sheet1", "A1").formula, null);
  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.value, "original");

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession↔DocumentController binder blocks sheet-level formatting when it intersects restricted ranges", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc });
  session.setPermissions({
    role: "editor",
    userId: "u-b",
    rangeRestrictions: [
      {
        // Restrict a range that does *not* include A1 so naive "probe A1" checks
        // would incorrectly allow sheet-level formatting.
        sheetName: "Sheet1",
        startRow: 1,
        startCol: 1,
        endRow: 2,
        endCol: 2,
        readAllowlist: [],
        editAllowlist: ["u-a"],
      },
    ],
  });

  const dc = new DocumentController();
  const binder = await bindCollabSessionToDocumentController({ session, documentController: dc });

  assert.equal(dc.getSheetDefaultStyleId("Sheet1"), 0);

  const before = Buffer.from(Y.encodeStateAsUpdate(doc));

  let updateCount = 0;
  const onUpdate = () => {
    updateCount += 1;
  };
  doc.on("update", onUpdate);

  // Attempt a sheet-wide format edit (affects all cells, including restricted ones).
  dc.setSheetFormat("Sheet1", { font: { bold: true } });

  await new Promise((r) => setTimeout(r, 25));

  doc.off("update", onUpdate);

  // The local UI should be reverted and the shared Yjs document must not change.
  assert.equal(dc.getSheetDefaultStyleId("Sheet1"), 0);
  assert.equal(updateCount, 0);
  assert.deepEqual(Buffer.from(Y.encodeStateAsUpdate(doc)), before);

  binder.destroy();
  session.destroy();
  doc.destroy();
});

test("CollabSession↔DocumentController binder encrypts protected cells and decrypts when key is available", async () => {
  const docId = "collab-session-documentController-encryption-test-doc";
  const docA = new Y.Doc({ guid: docId });
  const docB = new Y.Doc({ guid: docId });
  const disconnect = connectDocs(docA, docB);

  const keyBytes = new Uint8Array(32).fill(7);
  const keyForA1 = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  const sessionA = createCollabSession({ doc: docA, encryption: { keyForCell: keyForA1 } });
  const sessionB = createCollabSession({ doc: docB });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = await bindCollabSessionToDocumentController({ session: sessionA, documentController: dcA });
  /** @type {any[] | null} */
  let rejected = null;
  let binderB = await bindCollabSessionToDocumentController({
    session: sessionB,
    documentController: dcB,
    maskCellFormat: true,
    onEditRejected: (deltas) => {
      rejected = deltas;
    },
  });

  // Edit through the DocumentController path so binder is responsible for encryption.
  dcA.setCellValue("Sheet1", "A1", "top-secret");
  dcA.setRangeFormat("Sheet1", "A1", { font: { italic: true } });

  await waitForCondition(() => {
    const cellA = dcA.getCell("Sheet1", "A1");
    return cellA.value === "top-secret" && cellA.formula == null && cellA.styleId !== 0;
  });

  // Ensure formatting has propagated into B's underlying Yjs doc before asserting
  // the UI-level styleId is suppressed.
  await waitForCondition(() => {
    const cellMapB = sessionB.cells.get("Sheet1:0:0");
    return (
      cellMapB &&
      typeof cellMapB.get === "function" &&
      cellMapB.get("enc") != null &&
      cellMapB.get("format") != null
    );
  });

  await waitForCondition(() => {
    const cellB = dcB.getCell("Sheet1", "A1");
    return cellB.value === "###" && cellB.formula == null && cellB.styleId === 0;
  });

  // Attempt an edit without the encryption key. This simulates a UI path that bypasses
  // `DocumentController.canEditCell` (e.g. via `applyExternalDeltas`), and should be
  // rejected with an explicit encryption reason.
  const before = dcB.getCell("Sheet1", "A1");
  dcB.applyExternalDeltas(
    [
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        before,
        after: { value: "hacked", formula: null, styleId: before.styleId },
      },
    ],
    { recalc: false },
  );

  assert.ok(Array.isArray(rejected), "expected onEditRejected to be called for missing encryption key");
  assert.equal(rejected?.[0]?.rejectionReason, "encryption");

  await waitForCondition(() => {
    const cellMap = sessionA.cells.get("Sheet1:0:0");
    return cellMap && typeof cellMap.get === "function" && cellMap.get("enc") != null;
  });

  // Raw Yjs should not contain plaintext.
  const cellMap = sessionA.cells.get("Sheet1:0:0");
  assert.ok(cellMap, "expected Yjs cell map to exist");
  assert.equal(cellMap.get("value"), undefined);
  assert.equal(cellMap.get("formula"), undefined);
  assert.ok(cellMap.get("enc"), "expected encrypted payload under `enc`");
  assert.ok(cellMap.get("format"), "expected plaintext `format` alongside encrypted payload");
  assert.equal(JSON.stringify(cellMap.toJSON()).includes("top-secret"), false);

  // Now "grant" the key on B by recreating the session and re-binding.
  binderB.destroy();
  sessionB.destroy();
  const sessionBWithKey = createCollabSession({ doc: docB, encryption: { keyForCell: keyForA1 } });
  binderB = await bindCollabSessionToDocumentController({ session: sessionBWithKey, documentController: dcB, maskCellFormat: true });

  await waitForCondition(() => {
    const cell = dcB.getCell("Sheet1", "A1");
    return cell.value === "top-secret" && cell.styleId !== 0;
  });

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionBWithKey.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession↔DocumentController binder encryptFormat encrypts cell formatting and removes plaintext `format`", async () => {
  const docId = "collab-session-documentController-encryption-test-doc-encryptFormat";
  const docA = new Y.Doc({ guid: docId });
  const docB = new Y.Doc({ guid: docId });
  const disconnect = connectDocs(docA, docB);

  const keyBytes = new Uint8Array(32).fill(7);
  const keyForA1 = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  const sessionA = createCollabSession({ doc: docA, encryption: { keyForCell: keyForA1, encryptFormat: true } });
  const sessionB = createCollabSession({ doc: docB, encryption: { keyForCell: () => null, encryptFormat: true } });

  const dcA = new DocumentController();
  const dcB = new DocumentController();

  const binderA = await bindCollabSessionToDocumentController({ session: sessionA, documentController: dcA });
  let binderB = await bindCollabSessionToDocumentController({ session: sessionB, documentController: dcB });

  // Seed a legacy plaintext `style` key to ensure encryptFormat writes scrub it.
  docA.transact(() => {
    const cell = new Y.Map();
    cell.set("style", { font: { underline: true } });
    sessionA.cells.set("Sheet1:0:0", cell);
  });

  dcA.setCellValue("Sheet1", "A1", "top-secret");
  dcA.setRangeFormat("Sheet1", "A1", { font: { bold: true } });

  await waitForCondition(() => {
    const cellA = dcA.getCell("Sheet1", "A1");
    return cellA.value === "top-secret" && cellA.formula == null && cellA.styleId !== 0;
  });

  await waitForCondition(() => {
    const cellB = dcB.getCell("Sheet1", "A1");
    return cellB.value === "###" && cellB.formula == null && cellB.styleId === 0;
  });

  await waitForCondition(() => {
    const cellMap = sessionA.cells.get("Sheet1:0:0");
    return cellMap && typeof cellMap.get === "function" && cellMap.get("enc") != null;
  });

  const cellMap = sessionA.cells.get("Sheet1:0:0");
  assert.ok(cellMap, "expected Yjs cell map to exist");
  assert.equal(cellMap.get("value"), undefined);
  assert.equal(cellMap.get("formula"), undefined);
  assert.equal(cellMap.get("format"), undefined);
  assert.equal(cellMap.get("style"), undefined);
  assert.ok(cellMap.get("enc"), "expected encrypted payload under `enc`");

  // Now "grant" the key on B by recreating the session and re-binding.
  binderB.destroy();
  sessionB.destroy();
  const sessionBWithKey = createCollabSession({ doc: docB, encryption: { keyForCell: keyForA1, encryptFormat: true } });
  binderB = await bindCollabSessionToDocumentController({ session: sessionBWithKey, documentController: dcB });

  await waitForCondition(() => {
    const cellB = dcB.getCell("Sheet1", "A1");
    const format = dcB.styleTable.get(cellB.styleId);
    return cellB.value === "top-secret" && cellB.styleId !== 0 && format?.font?.bold === true;
  });

  binderA.destroy();
  binderB.destroy();
  sessionA.destroy();
  sessionBWithKey.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession↔DocumentController binder normalizes formula text (trimming + bare '=')", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc });
  session.setPermissions({ role: "editor", userId: "u-a", rangeRestrictions: [] });

  const dc = new DocumentController();
  const binder = await bindCollabSessionToDocumentController({ session, documentController: dc });

  await session.setCellFormula("Sheet1:0:0", "=  SUM(A1:A3)  ");
  await waitForCondition(() => dc.getCell("Sheet1", "A1").formula === "=SUM(A1:A3)");

  // A bare "=" (or whitespace-only formula) is treated as empty and clears the cell.
  await session.setCellFormula("Sheet1:0:1", "=");
  await waitForCondition(() => {
    const cell = dc.getCell("Sheet1", "B1");
    return cell.formula == null && cell.value == null;
  });

  binder.destroy();
  session.destroy();
  doc.destroy();
});

test("bindCollabSessionToDocumentController forwards onEditRejected (permission rejection)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc });
  session.setPermissions({
    role: "editor",
    userId: "u-b",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        // Anyone can read, but only "other-user" can edit.
        readAllowlist: [],
        editAllowlist: ["other-user"],
      },
    ],
  });

  const dc = new DocumentController();

  /** @type {any[] | null} */
  let rejected = null;
  const binder = await bindCollabSessionToDocumentController({
    session,
    documentController: dc,
    onEditRejected: (deltas) => {
      rejected = deltas;
    },
  });

  const before = dc.getCell("Sheet1", "A1");
  // Bypass DocumentController.canEditCell (simulates a buggy caller / special edit path).
  dc.applyExternalDeltas(
    [
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        before,
        after: { value: "hacked", formula: null, styleId: before.styleId },
      },
    ],
    { recalc: false },
  );

  assert.ok(Array.isArray(rejected), "expected onEditRejected to be called");
  assert.equal(rejected?.length, 1);
  assert.equal(rejected?.[0]?.sheetId, "Sheet1");
  assert.equal(rejected?.[0]?.row, 0);
  assert.equal(rejected?.[0]?.col, 0);
  // Binder annotates the rejection cause so UIs can differentiate permission vs encryption failures.
  assert.equal(rejected?.[0]?.rejectionReason, "permission");

  binder.destroy();
  session.destroy();
  doc.destroy();
});
