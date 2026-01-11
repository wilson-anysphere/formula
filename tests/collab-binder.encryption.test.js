import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";

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

async function waitForCondition(condition, timeoutMs, intervalMs = 25) {
  const start = Date.now();
  while (Date.now() - start <= timeoutMs) {
    if (await condition()) return;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for condition");
}

test("Yjs↔DocumentController binder encrypts protected cell contents and masks without key (in-memory sync)", async (t) => {
  const docId = "collab-binder-encryption-test-doc";
  const docA = new Y.Doc({ guid: docId });
  const docB = new Y.Doc({ guid: docId });
  const disconnect = connectDocs(docA, docB);

  const keyBytes = new Uint8Array(32).fill(5);
  const keyForA1 = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  const controllerA = new DocumentController();
  const controllerB = new DocumentController();

  const binderA = bindYjsToDocumentController({
    ydoc: docA,
    documentController: controllerA,
    defaultSheetId: "Sheet1",
    userId: "u-a",
    encryption: { keyForCell: keyForA1 },
  });

  /** @type {ReturnType<typeof bindYjsToDocumentController>} */
  let binderB = bindYjsToDocumentController({
    ydoc: docB,
    documentController: controllerB,
    defaultSheetId: "Sheet1",
    userId: "u-b",
    // No key on B.
    // Use a pass-through mask function to ensure encrypted cells still show a fixed
    // placeholder even when masking isn't being applied due to read restrictions.
    maskCellValue: (value) => value,
  });

  t.after(() => {
    binderA.destroy();
    binderB.destroy();
    disconnect();
    docA.destroy();
    docB.destroy();
  });

  controllerA.setCellValue("Sheet1", "A1", "top-secret");

  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value === "###", 10_000);
  assert.equal(controllerB.getCell("Sheet1", "A1").value, "###");
  assert.equal(controllerB.getCell("Sheet1", "A1").formula, null);

  // Raw Yjs should not contain plaintext.
  const cellMap = docA.getMap("cells").get("Sheet1:0:0");
  assert.ok(cellMap, "expected Yjs cell map to exist");
  assert.equal(cellMap.get("value"), undefined);
  assert.equal(cellMap.get("formula"), undefined);
  assert.ok(cellMap.get("enc"), "expected encrypted payload under `enc`");
  assert.equal(JSON.stringify(cellMap.toJSON()).includes("top-secret"), false);

  // Now "grant" the key by rebinding on B.
  binderB.destroy();
  binderB = bindYjsToDocumentController({
    ydoc: docB,
    documentController: controllerB,
    defaultSheetId: "Sheet1",
    userId: "u-b",
    encryption: { keyForCell: keyForA1 },
  });

  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value === "top-secret", 10_000);
  assert.equal(controllerB.getCell("Sheet1", "A1").value, "top-secret");

  // Clearing the cell should keep it encrypted so unauthorized clients can't
  // infer whether the cell is empty.
  controllerA.clearCell("Sheet1", "A1");

  await waitForCondition(() => controllerA.getCell("Sheet1", "A1").value == null, 10_000);
  assert.equal(controllerA.getCell("Sheet1", "A1").value, null);

  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value == null, 10_000);
  assert.equal(controllerB.getCell("Sheet1", "A1").value, null);

  // Rebind without a key: should still be masked (enc marker remains).
  binderB.destroy();
  binderB = bindYjsToDocumentController({
    ydoc: docB,
    documentController: controllerB,
    defaultSheetId: "Sheet1",
    userId: "u-b",
  });

  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value === "###", 10_000);
  assert.equal(controllerB.getCell("Sheet1", "A1").value, "###");

  // Rebind with the key again: should decrypt back to an empty cell.
  binderB.destroy();
  binderB = bindYjsToDocumentController({
    ydoc: docB,
    documentController: controllerB,
    defaultSheetId: "Sheet1",
    userId: "u-b",
    encryption: { keyForCell: keyForA1 },
  });

  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value == null, 10_000);
  assert.equal(controllerB.getCell("Sheet1", "A1").value, null);
});

test("Yjs↔DocumentController binder uses maskCellValue hook when encrypted cell cannot be decrypted", async (t) => {
  const docId = "collab-binder-encryption-test-doc-mask-hook";
  const docA = new Y.Doc({ guid: docId });
  const docB = new Y.Doc({ guid: docId });
  const disconnect = connectDocs(docA, docB);

  const keyBytes = new Uint8Array(32).fill(5);
  const keyForA1 = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  const controllerA = new DocumentController();
  const controllerB = new DocumentController();

  const maskCellValue = () => "MASKED";

  const binderA = bindYjsToDocumentController({
    ydoc: docA,
    documentController: controllerA,
    defaultSheetId: "Sheet1",
    userId: "u-a",
    encryption: { keyForCell: keyForA1 },
    maskCellValue,
  });

  const binderB = bindYjsToDocumentController({
    ydoc: docB,
    documentController: controllerB,
    defaultSheetId: "Sheet1",
    userId: "u-b",
    // No key on B, so it should fall back to maskCellValue.
    maskCellValue,
  });

  t.after(() => {
    binderA.destroy();
    binderB.destroy();
    disconnect();
    docA.destroy();
    docB.destroy();
  });

  controllerA.setCellValue("Sheet1", "A1", "top-secret");

  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value === "MASKED", 10_000);
  assert.equal(controllerB.getCell("Sheet1", "A1").formula, null);
});
