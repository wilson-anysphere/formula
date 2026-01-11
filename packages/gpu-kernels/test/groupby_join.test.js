import test from "node:test";
import assert from "node:assert/strict";

import { CpuBackend } from "../src/index.js";
import { KernelEngine } from "../src/index.js";

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

test("group-by: two-key groupBySum2 is lexicographically sorted by (keyA, keyB)", async () => {
  const cpu = new CpuBackend();
  const keysA = new Uint32Array([1, 1, 2, 2, 2]);
  const keysB = new Uint32Array([10, 10, 10, 11, 11]);
  const values = new Float64Array([1, 2, 3, 4, 5]);

  const out = await cpu.groupBySum2(keysA, keysB, values);
  assert.deepEqual(Array.from(out.uniqueKeysA), [1, 2, 2]);
  assert.deepEqual(Array.from(out.uniqueKeysB), [10, 10, 11]);
  assert.deepEqual(Array.from(out.counts), [2, 1, 2]);
  assert.deepEqual(Array.from(out.sums), [3, 3, 9]);
});

test("group-by: two-key groupByCount2 supports signed keys and ordering", async () => {
  const cpu = new CpuBackend();
  const keysA = new Int32Array([-1, 0, -1, 0]);
  const keysB = new Int32Array([5, 5, 4, 5]);

  const out = await cpu.groupByCount2(keysA, keysB);
  assert.deepEqual(Array.from(out.uniqueKeysA), [-1, -1, 0]);
  assert.deepEqual(Array.from(out.uniqueKeysB), [4, 5, 5]);
  assert.deepEqual(Array.from(out.counts), [1, 1, 2]);
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

test("hashJoin: rejects mixed Int32Array/Uint32Array inputs (non-empty)", async () => {
  const cpu = new CpuBackend();
  await assert.rejects(
    () => cpu.hashJoin(new Int32Array([1]), new Uint32Array([1])),
    /hashJoin key type mismatch/
  );

  // Empty inputs short-circuit to empty output (no need to match dtypes).
  const out = await cpu.hashJoin(new Int32Array(), new Uint32Array([1]));
  assert.equal(out.leftIndex.length, 0);
  assert.equal(out.rightIndex.length, 0);
});

test("hashJoin: left join includes unmatched rows with rightIndex=0xFFFF_FFFF", async () => {
  const cpu = new CpuBackend();
  const leftKeys = new Uint32Array([1, 2, 3]);
  const rightKeys = new Uint32Array([2]);

  const out = await cpu.hashJoin(leftKeys, rightKeys, { joinType: "left" });
  assert.deepEqual(Array.from(out.leftIndex), [0, 1, 2]);
  assert.deepEqual(Array.from(out.rightIndex), [0xffff_ffff, 0, 0xffff_ffff]);
});

test("hashJoin: left join with empty right side returns one unmatched row per left key", async () => {
  const cpu = new CpuBackend();
  const leftKeys = new Uint32Array([10, 20]);
  const out = await cpu.hashJoin(leftKeys, new Uint32Array(), { joinType: "left" });
  assert.deepEqual(Array.from(out.leftIndex), [0, 1]);
  assert.deepEqual(Array.from(out.rightIndex), [0xffff_ffff, 0xffff_ffff]);

  const inner = await cpu.hashJoin(leftKeys, new Uint32Array(), { joinType: "inner" });
  assert.equal(inner.leftIndex.length, 0);
  assert.equal(inner.rightIndex.length, 0);
});

test("KernelEngine: hashJoin left join passes options through (CPU backend)", async () => {
  const engine = new KernelEngine({ precision: "excel", gpu: { enabled: false } });
  const leftKeys = new Uint32Array([1, 2]);
  const rightKeys = new Uint32Array([2]);
  const out = await engine.hashJoin(leftKeys, rightKeys, { joinType: "left" });
  assert.deepEqual(Array.from(out.leftIndex), [0, 1]);
  assert.deepEqual(Array.from(out.rightIndex), [0xffff_ffff, 0]);
});

test("KernelEngine: hashJoin rejects invalid joinType", async () => {
  const engine = new KernelEngine({ precision: "excel", gpu: { enabled: false } });
  await assert.rejects(
    // @ts-ignore
    () => engine.hashJoin(new Uint32Array([1]), new Uint32Array(), { joinType: "outer" }),
    /hashJoin joinType must be/
  );
});

test("group-by/hashJoin: supports key 0xFFFF_FFFF (u32 max) in unsigned mode", async () => {
  const cpu = new CpuBackend();
  const k = 0xffff_ffff;

  const keys = new Uint32Array([k, 1, k, 1]);
  const values = new Float64Array([1, 2, 3, 4]);
  const grouped = await cpu.groupBySum(keys, values);
  assert.deepEqual(Array.from(grouped.uniqueKeys), [1, k]);
  assert.deepEqual(Array.from(grouped.counts), [2, 2]);
  assert.deepEqual(Array.from(grouped.sums), [6, 4]);

  const leftKeys = new Uint32Array([k, 1, k]);
  const rightKeys = new Uint32Array([0, k, k]);
  const joined = await cpu.hashJoin(leftKeys, rightKeys);
  assert.deepEqual(Array.from(joined.leftIndex), [0, 0, 2, 2]);
  assert.deepEqual(Array.from(joined.rightIndex), [1, 2, 1, 2]);
});
