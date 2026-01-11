import test from "node:test";
import assert from "node:assert/strict";

import { CpuBackend } from "../src/index.js";

test("cpu: sum", async () => {
  const cpu = new CpuBackend();
  const result = await cpu.sum(new Float64Array([1, 2, 3, 4]));
  assert.equal(result, 10);
});

test("cpu: sumproduct", async () => {
  const cpu = new CpuBackend();
  const result = await cpu.sumproduct(new Float64Array([1, 2, 3]), new Float64Array([2, 3, 4]));
  assert.equal(result, 1 * 2 + 2 * 3 + 3 * 4);
});

test("cpu: min/max/average/count", async () => {
  const cpu = new CpuBackend();
  const values = new Float64Array([3, 1, 2, -1]);
  assert.equal(await cpu.min(values), -1);
  assert.equal(await cpu.max(values), 3);
  assert.equal(await cpu.count(values), 4);
  assert.equal(await cpu.average(values), (3 + 1 + 2 + -1) / 4);
});

test("cpu: min/max/average propagate NaN and handle Infinity", async () => {
  const cpu = new CpuBackend();
  const values = new Float64Array([3, Number.POSITIVE_INFINITY, -1, Number.NEGATIVE_INFINITY, Number.NaN]);
  assert.ok(Number.isNaN(await cpu.min(values)));
  assert.ok(Number.isNaN(await cpu.max(values)));
  assert.ok(Number.isNaN(await cpu.average(values)));
});

test("cpu: min/max preserve signed zero like JS Math.min/Math.max", async () => {
  const cpu = new CpuBackend();
  const values = new Float64Array([0, -0]);
  const min = await cpu.min(values);
  const max = await cpu.max(values);
  assert.ok(Object.is(min, -0), `expected -0, got ${min}`);
  assert.ok(Object.is(max, 0) && !Object.is(max, -0), `expected +0, got ${max}`);
});

test("cpu: min/max on empty arrays", async () => {
  const cpu = new CpuBackend();
  assert.equal(await cpu.min(new Float64Array()), Number.POSITIVE_INFINITY);
  assert.equal(await cpu.max(new Float64Array()), Number.NEGATIVE_INFINITY);
  assert.ok(Number.isNaN(await cpu.average(new Float64Array())));
  assert.equal(await cpu.count(new Float64Array()), 0);
});

test("cpu: mmult", async () => {
  const cpu = new CpuBackend();
  // 2x3 * 3x2 => 2x2
  const a = new Float64Array([1, 2, 3, 4, 5, 6]); // rows: [1 2 3], [4 5 6]
  const b = new Float64Array([7, 8, 9, 10, 11, 12]); // rows: [7 8], [9 10], [11 12]
  const out = await cpu.mmult(a, b, 2, 3, 2);
  assert.deepEqual(Array.from(out), [58, 64, 139, 154]);
});

test("cpu: sort", async () => {
  const cpu = new CpuBackend();
  const out = await cpu.sort(new Float64Array([3, 1, 2, -1]));
  assert.deepEqual(Array.from(out), [-1, 1, 2, 3]);
});

test("cpu: sort handles NaN and Infinity like TypedArray#sort", async () => {
  const cpu = new CpuBackend();
  const out = await cpu.sort(new Float64Array([3, Number.NaN, 1, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]));
  assert.deepEqual(Array.from(out), [Number.NEGATIVE_INFINITY, 1, 3, Number.POSITIVE_INFINITY, Number.NaN]);
});

test("cpu: histogram", async () => {
  const cpu = new CpuBackend();
  const values = new Float64Array([0, 0.49, 0.5, 0.99, 1.0, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]);
  const bins = await cpu.histogram(values, { min: 0, max: 1, bins: 2 });
  // With clamping, 1.0 falls in last bin.
  assert.deepEqual(Array.from(bins), [3, 4]);
});
