import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { diffYjsWorkbookSnapshots } from "../index.js";

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
  assert.equal(sheetDiff.modified.length, 0);
});

