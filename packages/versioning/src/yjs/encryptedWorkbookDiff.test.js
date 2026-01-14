import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { diffYjsWorkbookSnapshots } from "./diffWorkbookSnapshots.js";

function createWorkbookDoc() {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");
  const sheet = new Y.Map();
  sheet.set("id", "Sheet1");
  sheet.set("name", "Sheet1");
  sheets.push([sheet]);
  doc.getMap("cells");
  return doc;
}

test("diffYjsWorkbookSnapshots: encrypted cell payload changes are classified as modified", () => {
  const enc1 = { v: 1, alg: "AES-256-GCM", keyId: "k1", ivBase64: "iv1", tagBase64: "tag1", ciphertextBase64: "ct1" };
  const enc2 = { v: 1, alg: "AES-256-GCM", keyId: "k1", ivBase64: "iv2", tagBase64: "tag2", ciphertextBase64: "ct2" };

  const doc = createWorkbookDoc();
  const cells = doc.getMap("cells");

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("enc", enc1);
    cells.set("Sheet1:0:0", cell);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    const cell = cells.get("Sheet1:0:0");
    assert.ok(cell instanceof Y.Map);
    cell.set("enc", enc2);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheetDiff = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1")?.diff;
  assert.ok(sheetDiff);

  assert.equal(sheetDiff.modified.length, 1);
  assert.deepEqual(sheetDiff.modified[0].cell, { row: 0, col: 0 });
  assert.equal(sheetDiff.modified[0].oldEncrypted, true);
  assert.equal(sheetDiff.modified[0].newEncrypted, true);
  assert.equal(sheetDiff.modified[0].oldKeyId, "k1");
  assert.equal(sheetDiff.modified[0].newKeyId, "k1");
  // User-facing diff records should not include ciphertext payloads.
  assert.equal("enc" in sheetDiff.modified[0], false);
});

test("diffYjsWorkbookSnapshots: encrypted cell moves are detected via encrypted payload signature", () => {
  const enc = { v: 1, alg: "AES-256-GCM", keyId: "k1", ivBase64: "iv", tagBase64: "tag", ciphertextBase64: "ct" };

  const doc = createWorkbookDoc();
  const cells = doc.getMap("cells");

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("enc", enc);
    cells.set("Sheet1:0:0", cell);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    cells.delete("Sheet1:0:0");
    const moved = new Y.Map();
    moved.set("enc", enc);
    cells.set("Sheet1:0:1", moved);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheetDiff = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1")?.diff;
  assert.ok(sheetDiff);

  assert.equal(sheetDiff.moved.length, 1);
  assert.deepEqual(sheetDiff.moved[0].oldLocation, { row: 0, col: 0 });
  assert.deepEqual(sheetDiff.moved[0].newLocation, { row: 0, col: 1 });
  assert.equal(sheetDiff.moved[0].encrypted, true);
  assert.equal(sheetDiff.moved[0].keyId, "k1");
  assert.equal(sheetDiff.added.length, 0);
  assert.equal(sheetDiff.removed.length, 0);
});

test("diffYjsWorkbookSnapshots: encrypted moves work across legacy cell key encodings", () => {
  const enc = { v: 1, alg: "AES-256-GCM", keyId: "k1", ivBase64: "iv", tagBase64: "tag", ciphertextBase64: "ct" };

  const doc = createWorkbookDoc();
  const cells = doc.getMap("cells");

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("enc", enc);
    // Store under a legacy key encoding.
    cells.set("Sheet1:0,0", cell);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    cells.delete("Sheet1:0,0");
    const moved = new Y.Map();
    moved.set("enc", enc);
    // Store under canonical key encoding at a new coordinate.
    cells.set("Sheet1:0:1", moved);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheetDiff = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1")?.diff;
  assert.ok(sheetDiff);

  assert.equal(sheetDiff.moved.length, 1);
  assert.deepEqual(sheetDiff.moved[0].oldLocation, { row: 0, col: 0 });
  assert.deepEqual(sheetDiff.moved[0].newLocation, { row: 0, col: 1 });
  assert.equal(sheetDiff.moved[0].encrypted, true);
  assert.equal(sheetDiff.moved[0].keyId, "k1");
  assert.equal(sheetDiff.added.length, 0);
  assert.equal(sheetDiff.removed.length, 0);
});

test("diffYjsWorkbookSnapshots: encrypted format-only changes are classified as formatOnly", () => {
  const enc = { v: 1, alg: "AES-256-GCM", keyId: "k1", ivBase64: "iv", tagBase64: "tag", ciphertextBase64: "ct" };

  const doc = createWorkbookDoc();
  const cells = doc.getMap("cells");

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("enc", enc);
    cell.set("format", { bold: true });
    cells.set("Sheet1:0:0", cell);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    const cell = cells.get("Sheet1:0:0");
    assert.ok(cell instanceof Y.Map);
    cell.set("format", { bold: false });
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheetDiff = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1")?.diff;
  assert.ok(sheetDiff);

  assert.equal(sheetDiff.formatOnly.length, 1);
  assert.deepEqual(sheetDiff.formatOnly[0].cell, { row: 0, col: 0 });
  assert.equal(sheetDiff.formatOnly[0].oldEncrypted, true);
  assert.equal(sheetDiff.formatOnly[0].newEncrypted, true);
  assert.equal(sheetDiff.formatOnly[0].oldKeyId, "k1");
  assert.equal(sheetDiff.formatOnly[0].newKeyId, "k1");
  assert.equal(sheetDiff.modified.length, 0);
});

test("diffYjsWorkbookSnapshots: encrypted cells win over plaintext duplicates across legacy keys", () => {
  const enc1 = { v: 1, alg: "AES-256-GCM", keyId: "k1", ivBase64: "iv1", tagBase64: "tag1", ciphertextBase64: "ct1" };
  const enc2 = { v: 1, alg: "AES-256-GCM", keyId: "k1", ivBase64: "iv2", tagBase64: "tag2", ciphertextBase64: "ct2" };

  const doc = createWorkbookDoc();
  const cells = doc.getMap("cells");

  doc.transact(() => {
    const encrypted = new Y.Map();
    encrypted.set("enc", enc1);
    // Encrypted payload stored under a legacy key encoding.
    cells.set("Sheet1:0,0", encrypted);

    // Plaintext duplicate stored under the canonical key (should be ignored by diffs).
    const plain = new Y.Map();
    plain.set("value", "leaked");
    cells.set("Sheet1:0:0", plain);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    const encrypted = cells.get("Sheet1:0,0");
    assert.ok(encrypted instanceof Y.Map);
    encrypted.set("enc", enc2);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheetDiff = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1")?.diff;
  assert.ok(sheetDiff);

  assert.equal(sheetDiff.modified.length, 1);
  assert.deepEqual(sheetDiff.modified[0].cell, { row: 0, col: 0 });
  assert.equal(sheetDiff.modified[0].oldValue, null);
  assert.equal(sheetDiff.modified[0].newValue, null);
  assert.equal(sheetDiff.modified[0].oldEncrypted, true);
  assert.equal(sheetDiff.modified[0].newEncrypted, true);
  assert.equal(sheetDiff.modified[0].oldKeyId, "k1");
  assert.equal(sheetDiff.modified[0].newKeyId, "k1");
});

test("diffYjsWorkbookSnapshots: enc=null markers win over plaintext duplicates across legacy keys", () => {
  const doc = createWorkbookDoc();
  const cells = doc.getMap("cells");

  doc.transact(() => {
    const encrypted = new Y.Map();
    // Explicit null marker (should still be treated as encrypted).
    encrypted.set("enc", null);
    // Stored under a legacy key encoding.
    cells.set("Sheet1:0,0", encrypted);

    const plain = new Y.Map();
    plain.set("value", "leaked-1");
    cells.set("Sheet1:0:0", plain);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  // Mutate the plaintext duplicate; diff should ignore it because an `enc` marker exists.
  doc.transact(() => {
    const plain = cells.get("Sheet1:0:0");
    assert.ok(plain instanceof Y.Map);
    plain.set("value", "leaked-2");
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheetDiff = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1")?.diff;
  assert.ok(sheetDiff);

  assert.equal(sheetDiff.added.length, 0);
  assert.equal(sheetDiff.removed.length, 0);
  assert.equal(sheetDiff.modified.length, 0);
  assert.equal(sheetDiff.formatOnly.length, 0);
  assert.equal(sheetDiff.moved.length, 0);
});

test("diffYjsWorkbookSnapshots: enc=null markers are not matched as moves against plaintext format-only cells", () => {
  const doc = createWorkbookDoc();
  const cells = doc.getMap("cells");

  doc.transact(() => {
    const encryptedMarker = new Y.Map();
    encryptedMarker.set("enc", null);
    encryptedMarker.set("format", { bold: true });
    cells.set("Sheet1:0:0", encryptedMarker);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    cells.delete("Sheet1:0:0");
    const formatOnly = new Y.Map();
    formatOnly.set("format", { bold: true });
    cells.set("Sheet1:0:1", formatOnly);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheetDiff = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1")?.diff;
  assert.ok(sheetDiff);

  assert.equal(sheetDiff.moved.length, 0);
  assert.equal(sheetDiff.removed.length, 1);
  assert.deepEqual(sheetDiff.removed[0].cell, { row: 0, col: 0 });
  assert.equal(sheetDiff.removed[0].oldEncrypted, true);

  assert.equal(sheetDiff.added.length, 1);
  assert.deepEqual(sheetDiff.added[0].cell, { row: 0, col: 1 });
  assert.equal("newEncrypted" in sheetDiff.added[0], false);

  assert.equal(sheetDiff.modified.length, 0);
  assert.equal(sheetDiff.formatOnly.length, 0);
});

test("diffYjsWorkbookSnapshots: encrypted format-only changes work even when format is only present on a legacy duplicate", () => {
  const enc = { v: 1, alg: "AES-256-GCM", keyId: "k1", ivBase64: "iv", tagBase64: "tag", ciphertextBase64: "ct" };

  const doc = createWorkbookDoc();
  const cells = doc.getMap("cells");

  // Insert legacy key first (so it appears first in map iteration order),
  // then insert canonical key with the same ciphertext but without `format`.
  doc.transact(() => {
    const legacy = new Y.Map();
    legacy.set("enc", enc);
    legacy.set("format", { bold: true });
    cells.set("Sheet1:0,0", legacy);

    const canonical = new Y.Map();
    canonical.set("enc", enc);
    cells.set("Sheet1:0:0", canonical);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    const legacy = cells.get("Sheet1:0,0");
    assert.ok(legacy instanceof Y.Map);
    legacy.set("format", { bold: false });
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheetDiff = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1")?.diff;
  assert.ok(sheetDiff);

  assert.equal(sheetDiff.formatOnly.length, 1);
  assert.deepEqual(sheetDiff.formatOnly[0].cell, { row: 0, col: 0 });
  assert.equal(sheetDiff.formatOnly[0].oldEncrypted, true);
  assert.equal(sheetDiff.formatOnly[0].newEncrypted, true);
  assert.equal(sheetDiff.formatOnly[0].oldKeyId, "k1");
  assert.equal(sheetDiff.formatOnly[0].newKeyId, "k1");
  assert.equal(sheetDiff.modified.length, 0);
});

test("diffYjsWorkbookSnapshots: encrypted cell additions include encrypted metadata", () => {
  const enc = { v: 1, alg: "AES-256-GCM", keyId: "k1", ivBase64: "iv", tagBase64: "tag", ciphertextBase64: "ct" };

  const doc = createWorkbookDoc();
  const cells = doc.getMap("cells");

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("enc", enc);
    cells.set("Sheet1:0:0", cell);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheetDiff = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1")?.diff;
  assert.ok(sheetDiff);

  assert.equal(sheetDiff.added.length, 1);
  assert.deepEqual(sheetDiff.added[0].cell, { row: 0, col: 0 });
  assert.equal(sheetDiff.added[0].newEncrypted, true);
  assert.equal(sheetDiff.added[0].newKeyId, "k1");
  assert.equal("enc" in sheetDiff.added[0], false);
});

test("diffYjsWorkbookSnapshots: encrypted cell removals include encrypted metadata", () => {
  const enc = { v: 1, alg: "AES-256-GCM", keyId: "k1", ivBase64: "iv", tagBase64: "tag", ciphertextBase64: "ct" };

  const doc = createWorkbookDoc();
  const cells = doc.getMap("cells");

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("enc", enc);
    cells.set("Sheet1:0:0", cell);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    cells.delete("Sheet1:0:0");
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheetDiff = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1")?.diff;
  assert.ok(sheetDiff);

  assert.equal(sheetDiff.removed.length, 1);
  assert.deepEqual(sheetDiff.removed[0].cell, { row: 0, col: 0 });
  assert.equal(sheetDiff.removed[0].oldEncrypted, true);
  assert.equal(sheetDiff.removed[0].oldKeyId, "k1");
  assert.equal("enc" in sheetDiff.removed[0], false);
});
