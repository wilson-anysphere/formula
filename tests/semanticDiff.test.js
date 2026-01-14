import test from "node:test";
import assert from "node:assert/strict";

import { cellKey, semanticDiff } from "../packages/versioning/src/diff/semanticDiff.js";

function sheetFromObject(obj) {
  const cells = new Map();
  for (const [k, v] of Object.entries(obj)) {
    cells.set(k, v);
  }
  return { cells };
}

test("semanticDiff: added cell", () => {
  const before = sheetFromObject({});
  const after = sheetFromObject({
    [cellKey(0, 0)]: { value: 123 },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.added.length, 1);
  assert.deepEqual(diff.added[0].cell, { row: 0, col: 0 });
  assert.equal(diff.removed.length, 0);
  assert.equal(diff.modified.length, 0);
  assert.equal(diff.moved.length, 0);
});

test("semanticDiff: removed cell", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: "x" },
  });
  const after = sheetFromObject({});
  const diff = semanticDiff(before, after);
  assert.equal(diff.removed.length, 1);
  assert.deepEqual(diff.removed[0].cell, { row: 0, col: 0 });
  assert.equal(diff.added.length, 0);
  assert.equal(diff.modified.length, 0);
  assert.equal(diff.moved.length, 0);
});

test("semanticDiff: modified cell (value)", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: 1 },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { value: 2 },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.modified.length, 1);
  assert.deepEqual(diff.modified[0].cell, { row: 0, col: 0 });
  assert.equal(diff.modified[0].oldValue, 1);
  assert.equal(diff.modified[0].newValue, 2);
  // Backwards-compat: non-encrypted cells should not include encryption metadata fields.
  assert.equal("oldEncrypted" in diff.modified[0], false);
  assert.equal("newEncrypted" in diff.modified[0], false);
  assert.equal("oldKeyId" in diff.modified[0], false);
  assert.equal("newKeyId" in diff.modified[0], false);
});

test("semanticDiff: format-only change", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: 1, format: { bold: false } },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { value: 1, format: { bold: true } },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.formatOnly.length, 1);
  assert.deepEqual(diff.formatOnly[0].cell, { row: 0, col: 0 });
  assert.equal(diff.modified.length, 0);
});

test("semanticDiff: moved cell detection", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: "move-me", formula: "=A1+B1" },
  });
  const after = sheetFromObject({
    [cellKey(2, 3)]: { value: "move-me", formula: "=B1 + A1" }, // commutative equiv
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.moved.length, 1);
  assert.deepEqual(diff.moved[0].oldLocation, { row: 0, col: 0 });
  assert.deepEqual(diff.moved[0].newLocation, { row: 2, col: 3 });
  // Backwards-compat: non-encrypted moves should not include encryption metadata fields.
  assert.equal("encrypted" in diff.moved[0], false);
  assert.equal("keyId" in diff.moved[0], false);
  assert.equal(diff.added.length, 0);
  assert.equal(diff.removed.length, 0);
});

test("semanticDiff: semantic-equivalent formulas are not modified", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: null, formula: "=A1 + B1" },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { value: null, formula: "=B1+A1" },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.modified.length, 0);
  assert.equal(diff.added.length, 0);
  assert.equal(diff.removed.length, 0);
  assert.equal(diff.moved.length, 0);
  assert.equal(diff.formatOnly.length, 0);
});

test("semanticDiff: encrypted cell modified is detected and includes key metadata", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { enc: { keyId: "k1", ciphertextBase64: "ct1" } },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { enc: { keyId: "k1", ciphertextBase64: "ct2" } },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.modified.length, 1);
  assert.deepEqual(diff.modified[0].cell, { row: 0, col: 0 });
  assert.equal(diff.modified[0].oldEncrypted, true);
  assert.equal(diff.modified[0].newEncrypted, true);
  assert.equal(diff.modified[0].oldKeyId, "k1");
  assert.equal(diff.modified[0].newKeyId, "k1");
});

test("semanticDiff: encrypted cell moved is detected via enc signature", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { enc: { keyId: "k1", ciphertextBase64: "ct" } },
  });
  const after = sheetFromObject({
    [cellKey(0, 1)]: { enc: { keyId: "k1", ciphertextBase64: "ct" } },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.moved.length, 1);
  assert.deepEqual(diff.moved[0].oldLocation, { row: 0, col: 0 });
  assert.deepEqual(diff.moved[0].newLocation, { row: 0, col: 1 });
  assert.equal(diff.moved[0].encrypted, true);
  assert.equal(diff.moved[0].keyId, "k1");
  assert.equal(diff.added.length, 0);
  assert.equal(diff.removed.length, 0);
});

test("semanticDiff: encrypted cell format-only changes are detected", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { enc: { keyId: "k1", ciphertextBase64: "ct" }, format: { bold: true } },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { enc: { keyId: "k1", ciphertextBase64: "ct" }, format: { bold: false } },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.formatOnly.length, 1);
  assert.deepEqual(diff.formatOnly[0].cell, { row: 0, col: 0 });
  assert.equal(diff.formatOnly[0].oldEncrypted, true);
  assert.equal(diff.formatOnly[0].newEncrypted, true);
  assert.equal(diff.formatOnly[0].oldKeyId, "k1");
  assert.equal(diff.formatOnly[0].newKeyId, "k1");
  assert.equal(diff.modified.length, 0);
});

test("semanticDiff: enc=null is treated as encrypted (fail closed)", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { enc: null, value: 1 },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { value: 2 },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.modified.length, 1);
  assert.deepEqual(diff.modified[0].cell, { row: 0, col: 0 });
  // When an `enc` marker is present, semanticDiff should not fall back to plaintext
  // fields, even if they exist in the payload.
  assert.equal(diff.modified[0].oldValue, null);
  assert.equal(diff.modified[0].newValue, 2);
  assert.equal(diff.modified[0].oldEncrypted, true);
  assert.equal(diff.modified[0].newEncrypted, false);
  assert.equal(diff.modified[0].oldKeyId, null);
  assert.equal(diff.modified[0].newKeyId, null);
 });

test("semanticDiff: encrypted cells do not leak value/formula fields even if provided", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { enc: { keyId: "k1", ciphertextBase64: "ct1" }, value: "leak", formula: "=LEAK()" },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { enc: { keyId: "k1", ciphertextBase64: "ct2" }, value: "leak2", formula: "=LEAK2()" },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.modified.length, 1);
  assert.deepEqual(diff.modified[0].cell, { row: 0, col: 0 });
  assert.equal(diff.modified[0].oldEncrypted, true);
  assert.equal(diff.modified[0].newEncrypted, true);
  assert.equal(diff.modified[0].oldValue, null);
  assert.equal(diff.modified[0].newValue, null);
  assert.equal(diff.modified[0].oldFormula, null);
  assert.equal(diff.modified[0].newFormula, null);
});

test("semanticDiff: NaN is treated as equal (no diff)", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: Number.NaN },
  });
  const after = sheetFromObject({
    [cellKey(0, 0)]: { value: Number.NaN },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.modified.length, 0);
  assert.equal(diff.added.length, 0);
  assert.equal(diff.removed.length, 0);
  assert.equal(diff.moved.length, 0);
  assert.equal(diff.formatOnly.length, 0);
});

test("semanticDiff: NaN does not collide with null for move signatures", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: Number.NaN },
  });
  const after = sheetFromObject({
    [cellKey(0, 1)]: { value: null },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.moved.length, 0);
  assert.equal(diff.removed.length, 1);
  assert.equal(diff.added.length, 1);
  assert.equal(Number.isNaN(diff.removed[0].oldValue), true);
  assert.equal(diff.added[0].newValue, null);
});

test("semanticDiff: BigInt does not collide with string for move signatures", () => {
  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: 10n },
  });
  const after = sheetFromObject({
    [cellKey(0, 1)]: { value: "10" },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.moved.length, 0);
  assert.equal(diff.removed.length, 1);
  assert.equal(diff.added.length, 1);
  assert.equal(diff.removed[0].oldValue, 10n);
  assert.equal(diff.added[0].newValue, "10");
});

test("semanticDiff: does not throw on cyclic cell values (move detection)", () => {
  /** @type {any} */
  const beforeValue = { kind: "cycle" };
  beforeValue.self = beforeValue;
  /** @type {any} */
  const afterValue = { kind: "cycle" };
  afterValue.self = afterValue;

  const before = sheetFromObject({
    [cellKey(0, 0)]: { value: beforeValue },
  });
  const after = sheetFromObject({
    [cellKey(0, 1)]: { value: afterValue },
  });
  const diff = semanticDiff(before, after);
  assert.equal(diff.moved.length, 1);
  assert.deepEqual(diff.moved[0].oldLocation, { row: 0, col: 0 });
  assert.deepEqual(diff.moved[0].newLocation, { row: 0, col: 1 });
});
