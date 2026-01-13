import assert from "node:assert/strict";
import test from "node:test";

import { deepEqual } from "../src/diff/deepEqual.js";

test("deepEqual: primitives", () => {
  assert.equal(deepEqual(1, 1), true);
  assert.equal(deepEqual(1, 2), false);
  assert.equal(deepEqual("a", "a"), true);
  assert.equal(deepEqual("a", "b"), false);
  assert.equal(deepEqual(null, null), true);
  assert.equal(deepEqual(null, undefined), false);
  assert.equal(deepEqual(undefined, undefined), true);
});

test("deepEqual: NaN is equal to NaN", () => {
  assert.equal(deepEqual(Number.NaN, Number.NaN), true);
});

test("deepEqual: arrays + sparse arrays", () => {
  assert.equal(deepEqual([1, 2, 3], [1, 2, 3]), true);
  assert.equal(deepEqual([1, 2, 3], [1, 2, 4]), false);

  const sparse = [];
  sparse.length = 1;
  assert.equal(deepEqual(sparse, [undefined]), false, "array hole should not equal explicit undefined");
});

test("deepEqual: plain objects", () => {
  assert.equal(deepEqual({ a: 1, b: 2 }, { b: 2, a: 1 }), true);
  assert.equal(deepEqual({ a: 1 }, { a: 1, b: undefined }), false);
});

test("deepEqual: repeated references do not have to match aliasing", () => {
  const shared = { a: 1 };
  const a = { x: shared, y: shared };
  const b = { x: { a: 1 }, y: { a: 1 } };
  assert.equal(deepEqual(a, b), true);
});

test("deepEqual: cycle structure is respected", () => {
  const a = { value: 1 };
  // @ts-ignore - create a cycle for testing
  a.self = a;
  const b = { value: 1 };
  // @ts-ignore - create a different cycle structure
  const inner = { value: 1 };
  // @ts-ignore - create a cycle for testing
  inner.self = inner;
  // @ts-ignore
  b.self = inner;

  assert.equal(deepEqual(a, b), false);

  const c = { value: 1 };
  // @ts-ignore - create a cycle for testing
  c.self = c;
  assert.equal(deepEqual(a, c), true);
});
