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

test("cpu: histogram", async () => {
  const cpu = new CpuBackend();
  const values = new Float64Array([0, 0.49, 0.5, 0.99, 1.0]);
  const bins = await cpu.histogram(values, { min: 0, max: 1, bins: 2 });
  // With clamping, 1.0 falls in last bin.
  assert.deepEqual(Array.from(bins), [2, 3]);
});

