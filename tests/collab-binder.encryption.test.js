import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";
import { encryptCellPlaintext } from "../packages/collab/encryption/src/index.node.js";
import { EncryptedRangeManager, createEncryptionPolicyFromDoc } from "../packages/collab/encrypted-ranges/src/index.ts";

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

  // Unauthorized clients should not be able to mutate encrypted cells. The binder
  // should revert local DocumentController edits and avoid writing plaintext into Yjs.
  controllerB.setCellValue("Sheet1", "A1", "hacked");
  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value === "###", 10_000);
  assert.equal(controllerA.getCell("Sheet1", "A1").value, "top-secret");

  // Raw Yjs should not contain plaintext.
  const cellMap = docA.getMap("cells").get("Sheet1:0:0");
  assert.ok(cellMap, "expected Yjs cell map to exist");
  assert.equal(cellMap.get("value"), undefined);
  assert.equal(cellMap.get("formula"), undefined);
  assert.ok(cellMap.get("enc"), "expected encrypted payload under `enc`");
  assert.equal(JSON.stringify(cellMap.toJSON()).includes("top-secret"), false);
  assert.equal(JSON.stringify(cellMap.toJSON()).includes("hacked"), false);

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

test("Yjs↔DocumentController binder refuses writes when encryption key id does not match existing enc payload", async (t) => {
  const docId = "collab-binder-encryption-test-doc-keyid-mismatch";
  const docA = new Y.Doc({ guid: docId });
  const docB = new Y.Doc({ guid: docId });
  const disconnect = connectDocs(docA, docB);

  const keyBytesA = new Uint8Array(32).fill(5);
  const keyBytesB = new Uint8Array(32).fill(6);
  const keyForA1Correct = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes: keyBytesA };
    }
    return null;
  };
  const keyForA1Wrong = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-2", keyBytes: keyBytesB };
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
    encryption: { keyForCell: keyForA1Correct },
  });

  const binderB = bindYjsToDocumentController({
    ydoc: docB,
    documentController: controllerB,
    defaultSheetId: "Sheet1",
    userId: "u-b",
    encryption: { keyForCell: keyForA1Wrong },
    // Ensure masked encrypted cells still show a fixed placeholder even when masking isn't
    // applied due to read restrictions.
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

  // B has a key, but it does not match the payload key id, so it should still see a mask.
  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value === "###", 10_000);
  assert.equal(controllerB.getCell("Sheet1", "A1").value, "###");

  // Attempt an unauthorized overwrite from B. This should be rejected and must not clobber
  // the encrypted payload (or make A lose the ability to decrypt).
  controllerB.setCellValue("Sheet1", "A1", "hacked");
  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value === "###", 10_000);
  assert.equal(controllerA.getCell("Sheet1", "A1").value, "top-secret");

  const cellMap = docA.getMap("cells").get("Sheet1:0:0");
  assert.ok(cellMap, "expected Yjs cell map to exist");
  assert.ok(cellMap.get("enc"), "expected encrypted payload under `enc`");
  assert.equal(cellMap.get("enc").keyId, "k-range-1");
  assert.equal(JSON.stringify(cellMap.toJSON()).includes("hacked"), false);
});

test("Yjs↔DocumentController binder refuses writes when enc payload schema is unsupported", async (t) => {
  const docId = "collab-binder-encryption-test-doc-unsupported-payload";
  const doc = new Y.Doc({ guid: docId });

  // Simulate a future encryption payload version that this binder cannot decrypt.
  doc.transact(() => {
    const cells = doc.getMap("cells");
    const cell = new Y.Map();
    cell.set("enc", {
      v: 2,
      alg: "AES-256-GCM",
      keyId: "k-range-1",
      ivBase64: "AA==",
      tagBase64: "AA==",
      ciphertextBase64: "AA==",
    });
    cells.set("Sheet1:0:0", cell);
  });

  const controller = new DocumentController();
  const keyBytes = new Uint8Array(32).fill(7);
  const keyForA1 = () => ({ keyId: "k-range-1", keyBytes });

  const binder = bindYjsToDocumentController({
    ydoc: doc,
    documentController: controller,
    defaultSheetId: "Sheet1",
    userId: "u-a",
    encryption: { keyForCell: keyForA1 },
    maskCellValue: (value) => value,
  });

  t.after(() => {
    binder.destroy();
    doc.destroy();
  });

  await waitForCondition(() => controller.getCell("Sheet1", "A1").value === "###", 10_000);

  // Attempt to overwrite. This should be rejected to avoid clobbering ciphertext
  // written by a newer client with an unknown schema.
  controller.setCellValue("Sheet1", "A1", "hacked");
  await waitForCondition(() => controller.getCell("Sheet1", "A1").value === "###", 10_000);

  const cellMap = doc.getMap("cells").get("Sheet1:0:0");
  assert.ok(cellMap, "expected Yjs cell map to exist");
  assert.equal(cellMap.get("enc").v, 2);
  assert.equal(JSON.stringify(cellMap.toJSON()).includes("hacked"), false);
});

test("Yjs↔DocumentController binder prefers encrypted payloads over plaintext duplicates across legacy key encodings", async (t) => {
  const docId = "collab-binder-encryption-test-doc-legacy-keys";
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

  const binderB = bindYjsToDocumentController({
    ydoc: docB,
    documentController: controllerB,
    defaultSheetId: "Sheet1",
    userId: "u-b",
    // No key on B.
    maskCellValue: (value) => value,
  });

  t.after(() => {
    binderA.destroy();
    binderB.destroy();
    disconnect();
    docA.destroy();
    docB.destroy();
  });

  const enc = await encryptCellPlaintext({
    plaintext: { value: "top-secret", formula: null },
    key: { keyId: "k-range-1", keyBytes },
    context: { docId, sheetId: "Sheet1", row: 0, col: 0 },
  });

  // Simulate a doc with historical key encodings: encrypted content stored under the
  // legacy `${sheetId}:${row},${col}` key while a plaintext duplicate exists under the
  // canonical `${sheetId}:${row}:${col}` key.
  docA.transact(() => {
    const cells = docA.getMap("cells");

    const encCell = new Y.Map();
    encCell.set("enc", enc);
    cells.set("Sheet1:0,0", encCell);

    const plaintext = new Y.Map();
    plaintext.set("value", "leaked");
    cells.set("Sheet1:0:0", plaintext);
  });

  await waitForCondition(() => controllerA.getCell("Sheet1", "A1").value === "top-secret", 10_000);
  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value === "###", 10_000);

  assert.equal(controllerB.getCell("Sheet1", "A1").value, "###");
  assert.equal(controllerB.getCell("Sheet1", "A1").formula, null);
});

test("Yjs↔DocumentController binder blocks plaintext writes when shouldEncryptCell requires encryption but key is missing (encryptedRanges policy)", async (t) => {
  const docId = "collab-binder-encryption-test-doc-encryptedRanges-policy";
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

  // Define the protected/encrypted range in the shared workbook metadata.
  const ranges = new EncryptedRangeManager({ doc: docA });
  ranges.add({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k-range-1" });

  const policy = createEncryptionPolicyFromDoc(docB);
  assert.equal(policy.shouldEncryptCell({ sheetId: "Sheet1", row: 0, col: 0 }), true);

  const controllerA = new DocumentController();
  const controllerB = new DocumentController();

  const binderA = bindYjsToDocumentController({
    ydoc: docA,
    documentController: controllerA,
    defaultSheetId: "Sheet1",
    userId: "u-a",
    encryption: { keyForCell: keyForA1 },
  });

  const binderB = bindYjsToDocumentController({
    ydoc: docB,
    documentController: controllerB,
    defaultSheetId: "Sheet1",
    userId: "u-b",
    // No key material on B, but B still knows the encryption policy (ranges + key ids).
    encryption: { keyForCell: () => null, shouldEncryptCell: policy.shouldEncryptCell },
  });

  t.after(() => {
    binderA.destroy();
    binderB.destroy();
    disconnect();
    docA.destroy();
    docB.destroy();
  });

  // Before any encrypted payload exists in Yjs, a client without keys should still refuse
  // to write plaintext into a cell that must be encrypted.
  controllerB.setCellValue("Sheet1", "A1", "hacked");
  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value !== "hacked", 10_000);
  assert.equal(controllerB.getCell("Sheet1", "A1").value, null);

  // Ensure no plaintext leaked into the shared Yjs doc.
  assert.equal(docA.getMap("cells").has("Sheet1:0:0"), false);

  // Now write from an authorized client. It should encrypt the cell.
  controllerA.setCellValue("Sheet1", "A1", "top-secret");
  await waitForCondition(() => controllerB.getCell("Sheet1", "A1").value === "###", 10_000);
  assert.equal(controllerB.getCell("Sheet1", "A1").value, "###");

  const cellMap = docA.getMap("cells").get("Sheet1:0:0");
  assert.ok(cellMap, "expected Yjs cell map to exist");
  assert.equal(cellMap.get("value"), undefined);
  assert.equal(cellMap.get("formula"), undefined);
  assert.ok(cellMap.get("enc"), "expected encrypted payload under `enc`");
  assert.equal(JSON.stringify(cellMap.toJSON()).includes("top-secret"), false);
});
