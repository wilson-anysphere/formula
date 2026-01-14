import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";
import { requireYjsCjs } from "../../yjs-utils/test/require-yjs-cjs.js";

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

test("migrateLegacyCellKeys prefers encrypted payloads over enc=null markers across duplicate keys", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  const marker = new Y.Map();
  marker.set("enc", null);
  // Simulate a corrupt/foreign writer leaving plaintext behind alongside the marker.
  marker.set("value", "plaintext-leak");
  marker.set("formula", "=1");
  cells.set("Sheet1:0:0", marker);

  const enc = {
    v: 1,
    alg: "AES-256-GCM",
    keyId: "k1",
    ivBase64: "AA==",
    tagBase64: "AA==",
    ciphertextBase64: "AA==",
  };
  // Real ciphertext stored under legacy encoding.
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

test("migrateLegacyCellKeys dryRun computes counts without mutating the doc", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");
  cells.set("Sheet1:1,2", makePlainCell(123));

  const before = Y.encodeStateAsUpdate(doc);
  const result = migrateLegacyCellKeys(doc, { dryRun: true });
  const after = Y.encodeStateAsUpdate(doc);

  assert.deepEqual(result, { migrated: 1, removed: 1, collisions: 0 });
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  // No mutation: legacy key remains; canonical key not created.
  assert.equal(cells.has("Sheet1:1,2"), true);
  assert.equal(cells.has("Sheet1:1:2"), false);
});

test("migrateLegacyCellKeys does not create the cells root when absent", () => {
  const doc = new Y.Doc();
  assert.equal(doc.share.has("cells"), false);

  const before = Y.encodeStateAsUpdate(doc);
  const result = migrateLegacyCellKeys(doc);
  const after = Y.encodeStateAsUpdate(doc);

  assert.deepEqual(result, { migrated: 0, removed: 0, collisions: 0 });
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);
  assert.equal(doc.share.has("cells"), false);
});

test("migrateLegacyCellKeys migrates null values (does not drop data)", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  cells.set("Sheet1:9,9", null);

  const result = migrateLegacyCellKeys(doc);
  assert.deepEqual(result, { migrated: 1, removed: 1, collisions: 0 });

  assert.equal(cells.has("Sheet1:9,9"), false);
  assert.equal(cells.has("Sheet1:9:9"), true);
  assert.equal(cells.get("Sheet1:9:9"), null);
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

test("migrateLegacyCellKeys clones foreign (CJS) nested Yjs types to local constructors", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const remoteCells = remote.getMap("cells");
  remote.transact(() => {
    const cell = new Ycjs.Map();
    cell.set("value", 1);
    const format = new Ycjs.Map();
    format.set("bold", true);
    cell.set("format", format);
    remoteCells.set("Sheet1:0,0", cell);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  doc.getMap("cells"); // ensure root exists in this module
  Ycjs.applyUpdate(doc, update);

  const result = migrateLegacyCellKeys(doc);
  assert.deepEqual(result, { migrated: 1, removed: 1, collisions: 0 });

  const cells = doc.getMap("cells");
  const migrated = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.ok(migrated instanceof Y.Map);
  const migratedFormat = migrated.get("format");
  assert.ok(migratedFormat instanceof Y.Map);
  assert.equal(migratedFormat.get("bold"), true);
});

test("migrateLegacyCellKeys conflict=prefer-legacy overwrites canonical plaintext with legacy plaintext", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  cells.set("Sheet1:0:0", makePlainCell("canonical"));
  cells.set("Sheet1:0,0", makePlainCell("legacy"));

  const result = migrateLegacyCellKeys(doc, { conflict: "prefer-legacy" });
  assert.deepEqual(result, { migrated: 1, removed: 1, collisions: 1 });

  assert.equal(cells.has("Sheet1:0,0"), false);
  const cell = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.equal(cell.get("value"), "legacy");
});

test("migrateLegacyCellKeys conflict=merge preserves canonical values and fills missing fields from legacy", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  const canonical = new Y.Map();
  canonical.set("value", "canonical");
  canonical.set("formula", null);
  cells.set("Sheet1:0:0", canonical);

  const legacy = new Y.Map();
  legacy.set("value", "legacy");
  legacy.set("formula", null);
  legacy.set("modifiedBy", "u-legacy");
  cells.set("Sheet1:0,0", legacy);

  const result = migrateLegacyCellKeys(doc, { conflict: "merge" });
  assert.deepEqual(result, { migrated: 1, removed: 1, collisions: 1 });

  assert.equal(cells.has("Sheet1:0,0"), false);
  const merged = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.ok(merged);
  assert.equal(merged.get("value"), "canonical");
  assert.equal(merged.get("modifiedBy"), "u-legacy");
});

test("migrateLegacyCellKeys conflict=merge supports record-style (plain object) cell payloads", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  cells.set("Sheet1:0:0", { value: "canonical", formula: null });
  cells.set("Sheet1:0,0", { value: "legacy", formula: null, modifiedBy: "u-legacy" });

  const result = migrateLegacyCellKeys(doc, { conflict: "merge" });
  assert.deepEqual(result, { migrated: 1, removed: 1, collisions: 1 });

  assert.equal(cells.has("Sheet1:0,0"), false);
  const merged = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.ok(merged instanceof Y.Map);
  assert.equal(merged.get("value"), "canonical");
  assert.equal(merged.get("modifiedBy"), "u-legacy");
});

test("migrateLegacyCellKeys resolves multiple legacy encodings for the same coordinate", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  cells.set("Sheet1:0,0", makePlainCell("from-colon-comma"));
  cells.set("r0c0", makePlainCell("from-r0c0"));

  const result = migrateLegacyCellKeys(doc, { defaultSheetId: "Sheet1" });
  assert.deepEqual(result, { migrated: 1, removed: 2, collisions: 1 });

  assert.equal(cells.has("Sheet1:0,0"), false);
  assert.equal(cells.has("r0c0"), false);
  const migrated = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.ok(migrated);
  // Deterministic winner: lexicographic sort picks `Sheet1:0,0` over `r0c0`.
  assert.equal(migrated.get("value"), "from-colon-comma");
});
