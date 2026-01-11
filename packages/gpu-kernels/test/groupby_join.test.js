import test from "node:test";
import assert from "node:assert/strict";

import { KernelEngine } from "../src/index.js";
import { CpuBackend } from "../src/index.js";

test("group-by: empty inputs", async () => {
  const engine = new KernelEngine({ precision: "excel", gpu: { enabled: false } });

  {
    const out = await engine.groupByCount(new Uint32Array());
    assert.equal(out.uniqueKeys.length, 0);
    assert.equal(out.counts.length, 0);
  }

  {
    const out = await engine.groupBySum(new Uint32Array(), new Float64Array());
    assert.equal(out.uniqueKeys.length, 0);
    assert.equal(out.sums.length, 0);
    assert.equal(out.counts.length, 0);
  }

  {
    const out = await engine.hashJoin(new Uint32Array(), new Uint32Array([1, 2, 3]));
    assert.equal(out.leftIndex.length, 0);
    assert.equal(out.rightIndex.length, 0);
  }
});

test("group-by: skewed distribution (many identical keys)", async () => {
  const cpu = new CpuBackend();
  const keys = new Uint32Array([3, 3, 3, 1, 1, 2, 3]);
  const values = new Float64Array([10, 1, 2, 5, 6, 7, 8]);

  const out = await cpu.groupBySum(keys, values);
  assert.deepEqual(Array.from(out.uniqueKeys), [1, 2, 3]);
  assert.deepEqual(Array.from(out.counts), [2, 1, 4]);
  assert.deepEqual(Array.from(out.sums), [11, 7, 21]);
});

test("group-by: high-cardinality keys", async () => {
  const cpu = new CpuBackend();
  const n = 256;
  const keys = new Uint32Array(n);
  const values = new Float64Array(n);
  for (let i = 0; i < n; i++) {
    keys[i] = i;
    values[i] = i * 0.5;
  }

  const out = await cpu.groupBySum(keys, values);
  assert.equal(out.uniqueKeys.length, n);
  for (let i = 0; i < n; i++) {
    assert.equal(out.uniqueKeys[i], i);
    assert.equal(out.counts[i], 1);
    assert.equal(out.sums[i], i * 0.5);
  }
});

test("group-by: NaN/Infinity handling matches JS numeric semantics", async () => {
  const cpu = new CpuBackend();
  const keys = new Uint32Array([1, 1, 1, 2, 2, 2]);
  const values = new Float64Array([1, Number.NaN, 2, Number.POSITIVE_INFINITY, -5, Number.NEGATIVE_INFINITY]);

  const sumOut = await cpu.groupBySum(keys, values);
  assert.deepEqual(Array.from(sumOut.uniqueKeys), [1, 2]);
  assert.deepEqual(Array.from(sumOut.counts), [3, 3]);
  assert.ok(Number.isNaN(sumOut.sums[0]));
  // Infinity + (-Infinity) => NaN (JS/IEEE-754 semantics).
  assert.ok(Number.isNaN(sumOut.sums[1]));

  const minOut = await cpu.groupByMin(keys, values);
  assert.deepEqual(Array.from(minOut.uniqueKeys), [1, 2]);
  assert.ok(Number.isNaN(minOut.mins[0]));
  assert.equal(minOut.mins[1], Number.NEGATIVE_INFINITY);

  const maxOut = await cpu.groupByMax(keys, values);
  assert.deepEqual(Array.from(maxOut.uniqueKeys), [1, 2]);
  assert.ok(Number.isNaN(maxOut.maxs[0]));
  assert.equal(maxOut.maxs[1], Number.POSITIVE_INFINITY);
});

test("group-by: min/max preserve signed zero like Math.min/Math.max", async () => {
  const cpu = new CpuBackend();
  const keys = new Uint32Array([7, 7]);
  const values = new Float64Array([0, -0]);

  const minOut = await cpu.groupByMin(keys, values);
  assert.ok(Object.is(minOut.mins[0], -0), `expected -0, got ${minOut.mins[0]}`);

  const maxOut = await cpu.groupByMax(keys, values);
  assert.ok(Object.is(maxOut.maxs[0], 0) && !Object.is(maxOut.maxs[0], -0), `expected +0, got ${maxOut.maxs[0]}`);
});

test("group-by: supports signed i32 keys (including -1)", async () => {
  const cpu = new CpuBackend();
  const keys = new Int32Array([-1, 0, -1, 2]);
  const values = new Float64Array([1, 2, 3, 4]);

  const out = await cpu.groupBySum(keys, values);
  assert.deepEqual(Array.from(out.uniqueKeys), [-1, 0, 2]);
  assert.deepEqual(Array.from(out.counts), [2, 1, 1]);
  assert.deepEqual(Array.from(out.sums), [4, 2, 4]);
});

test("hashJoin: inner join correctness with duplicates", async () => {
  const cpu = new CpuBackend();
  const leftKeys = new Uint32Array([1, 2, 2, 3]);
  const rightKeys = new Uint32Array([2, 2, 3, 4]);

  const out = await cpu.hashJoin(leftKeys, rightKeys);
  assert.deepEqual(Array.from(out.leftIndex), [1, 1, 2, 2, 3]);
  assert.deepEqual(Array.from(out.rightIndex), [0, 1, 0, 1, 2]);
});

test("hashJoin: supports signed i32 keys", async () => {
  const cpu = new CpuBackend();
  const leftKeys = new Int32Array([-1, 0, -1]);
  const rightKeys = new Int32Array([-1, -1, 2]);

  const out = await cpu.hashJoin(leftKeys, rightKeys);
  assert.deepEqual(Array.from(out.leftIndex), [0, 0, 2, 2]);
  assert.deepEqual(Array.from(out.rightIndex), [0, 1, 0, 1]);
});
