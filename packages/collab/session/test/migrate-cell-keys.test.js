import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { migrateLegacyCellKeys } from "../src/index.ts";

function makePlainCell(value) {
  const cell = new Y.Map();
  cell.set("value", value);
  cell.set("formula", null);
  return cell;
}

function makeEncryptedCell(enc) {
  const cell = new Y.Map();
  cell.set("enc", enc);
  cell.set("modified", Date.now());
  return cell;
}

test("migrateLegacyCellKeys rewrites `${sheetId}:${row},${col}` keys to canonical `${sheetId}:${row}:${col}`", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  cells.set("Sheet1:1,2", makePlainCell(123));

  const result = migrateLegacyCellKeys(doc);
  assert.deepEqual(result, { migrated: 1, removed: 1, collisions: 0 });

  assert.equal(cells.has("Sheet1:1,2"), false);
  assert.equal(cells.has("Sheet1:1:2"), true);
  const migratedCell = /** @type {any} */ (cells.get("Sheet1:1:2"));
  assert.equal(migratedCell.get("value"), 123);
});

test("migrateLegacyCellKeys rewrites `r{row}c{col}` keys using defaultSheetId", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  cells.set("r0c0", makePlainCell("hello"));

  const result = migrateLegacyCellKeys(doc, { defaultSheetId: "Main" });
  assert.deepEqual(result, { migrated: 1, removed: 1, collisions: 0 });

  assert.equal(cells.has("r0c0"), false);
  assert.equal(cells.has("Main:0:0"), true);
  const migratedCell = /** @type {any} */ (cells.get("Main:0:0"));
  assert.equal(migratedCell.get("value"), "hello");
});

test("migrateLegacyCellKeys collision: prefers canonical by default", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  cells.set("Sheet1:0:0", makePlainCell("canonical"));
  cells.set("Sheet1:0,0", makePlainCell("legacy"));

  const result = migrateLegacyCellKeys(doc);
  assert.deepEqual(result, { migrated: 0, removed: 1, collisions: 1 });

  assert.equal(cells.has("Sheet1:0,0"), false);
  const cell = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.equal(cell.get("value"), "canonical");
});

test("migrateLegacyCellKeys encrypted legacy cell must not leave plaintext behind", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  // Plaintext canonical cell (should be removed/overwritten).
  cells.set("Sheet1:0:0", makePlainCell("plaintext-secret"));

  // Encrypted legacy duplicate.
  const enc = {
    v: 1,
    alg: "AES-256-GCM",
    keyId: "k1",
    ivBase64: "AA==",
    tagBase64: "AA==",
    ciphertextBase64: "AA==",
  };
  cells.set("Sheet1:0,0", makeEncryptedCell(enc));

  const result = migrateLegacyCellKeys(doc);
  assert.deepEqual(result, { migrated: 1, removed: 1, collisions: 1 });

  assert.equal(cells.has("Sheet1:0,0"), false);
  const migratedCell = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.ok(migratedCell);
  assert.deepEqual(migratedCell.get("enc"), enc);
  assert.equal(migratedCell.get("value"), undefined);
  assert.equal(migratedCell.get("formula"), undefined);
});

test("migrateLegacyCellKeys is idempotent", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  cells.set("Sheet1:5,6", makePlainCell(999));

  const first = migrateLegacyCellKeys(doc);
  assert.deepEqual(first, { migrated: 1, removed: 1, collisions: 0 });

  const second = migrateLegacyCellKeys(doc);
  assert.deepEqual(second, { migrated: 0, removed: 0, collisions: 0 });
});

test("migrateLegacyCellKeys deep-clones nested Yjs types (avoids integration errors)", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  const legacyCell = new Y.Map();
  legacyCell.set("value", 1);
  const format = new Y.Map();
  format.set("bold", true);
  legacyCell.set("format", format);

  cells.set("Sheet1:0,0", legacyCell);

  const result = migrateLegacyCellKeys(doc);
  assert.deepEqual(result, { migrated: 1, removed: 1, collisions: 0 });

  assert.equal(cells.has("Sheet1:0,0"), false);
  const migrated = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.ok(migrated);
  const migratedFormat = migrated.get("format");
  assert.ok(migratedFormat instanceof Y.Map);
  assert.notEqual(migratedFormat, format);
  assert.equal(migratedFormat.get("bold"), true);
});
