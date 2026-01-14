import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { workbookStateFromYjsDoc } from "../packages/versioning/src/yjs/workbookState.js";

function createSheet1(doc) {
  const sheets = doc.getArray("sheets");
  const sheet = new Y.Map();
  sheet.set("id", "Sheet1");
  sheet.set("name", "Sheet1");
  sheets.push([sheet]);
}

test("workbookStateFromYjsDoc prefers encrypted payloads over plaintext duplicates across legacy key encodings", () => {
  const doc = new Y.Doc();
  createSheet1(doc);

  doc.transact(() => {
    const cells = doc.getMap("cells");

    const encCell = new Y.Map();
    encCell.set("enc", { keyId: "k1", ciphertextBase64: "ct" });
    // Encrypted payload stored under a legacy key encoding.
    cells.set("Sheet1:0,0", encCell);

    const plaintext = new Y.Map();
    plaintext.set("value", "leaked");
    // Plaintext duplicate stored under the canonical key.
    cells.set("Sheet1:0:0", plaintext);
  });

  const state = workbookStateFromYjsDoc(doc);
  const cell = state.cellsBySheet.get("Sheet1")?.cells.get("r0c0");
  assert.ok(cell);
  assert.deepEqual(cell.enc, { keyId: "k1", ciphertextBase64: "ct" });
  assert.equal(cell.value, null);
  assert.equal(cell.formula, null);
});

test("workbookStateFromYjsDoc prefers canonical encrypted records among duplicate encrypted keys", () => {
  const doc = new Y.Doc();
  createSheet1(doc);

  doc.transact(() => {
    const cells = doc.getMap("cells");

    const legacyEnc = new Y.Map();
    legacyEnc.set("enc", { keyId: "k1", ciphertextBase64: "legacy" });
    cells.set("Sheet1:0,0", legacyEnc);

    const canonicalEnc = new Y.Map();
    canonicalEnc.set("enc", { keyId: "k1", ciphertextBase64: "canonical" });
    cells.set("Sheet1:0:0", canonicalEnc);
  });

  const state = workbookStateFromYjsDoc(doc);
  const cell = state.cellsBySheet.get("Sheet1")?.cells.get("r0c0");
  assert.ok(cell);
  assert.deepEqual(cell.enc, { keyId: "k1", ciphertextBase64: "canonical" });
});

