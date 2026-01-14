import assert from "node:assert/strict";
import test from "node:test";

import {
  headSampleRows,
  randomSampleIndices,
  randomSampleRows,
  stratifiedSampleRows,
  systematicSampleRows,
  tailSampleRows,
} from "../src/sampling.js";

test("sampling: randomSampleRows is deterministic with a seed and returns unique rows", () => {
  const rows = Array.from({ length: 10 }, (_v, i) => i + 1);
  const sampleA = randomSampleRows(rows, 4, { seed: 123 });
  const sampleB = randomSampleRows(rows, 4, { seed: 123 });
  assert.deepStrictEqual(sampleA, sampleB);
  assert.equal(new Set(sampleA).size, 4);
  assert.ok(sampleA.every((v) => rows.includes(v)));
});

test("sampling: stratifiedSampleRows includes at least one row per stratum when possible", () => {
  const rows = [
    ["A", 1],
    ["A", 2],
    ["B", 3],
    ["B", 4],
    ["B", 5],
    ["C", 6],
  ];

  const sampled = stratifiedSampleRows(rows, 3, { getStratum: (r) => r[0], seed: 42 });
  assert.equal(sampled.length, 3);
  assert.deepStrictEqual(new Set(sampled.map((r) => r[0])), new Set(["A", "B", "C"]));
});

test("sampling: stratifiedSampleRows allocates extra samples to larger strata", () => {
  const rows = [
    ["A", 1],
    ["A", 2],
    ["B", 3],
    ["B", 4],
    ["B", 5],
    ["C", 6],
  ];

  const sampled = stratifiedSampleRows(rows, 4, { getStratum: (r) => r[0], seed: 7 });
  const counts = sampled.reduce((acc, r) => {
    // @ts-ignore - this is JS-only test code
    acc[r[0]] = (acc[r[0]] ?? 0) + 1;
    return acc;
  }, {});

  assert.equal(sampled.length, 4);
  assert.ok((counts.B ?? 0) >= 2);
  assert.ok((counts.A ?? 0) >= 1);
  assert.ok((counts.C ?? 0) >= 1);
});

test("sampling: stratifiedSampleRows handles large row sets without O(N) weighted helper arrays", () => {
  const total = 200_000;
  const strataCount = 100;
  const keys = Array.from({ length: strataCount }, (_v, i) => `S${i}`);
  const rows = Array.from({ length: total }, (_v, i) => keys[i % strataCount]);

  const sampledA = stratifiedSampleRows(rows, 10, { getStratum: (r) => r, seed: 123 });
  const sampledB = stratifiedSampleRows(rows, 10, { getStratum: (r) => r, seed: 123 });

  assert.deepStrictEqual(sampledA, sampledB);
  assert.equal(sampledA.length, 10);
  assert.equal(new Set(sampledA).size, 10);
});

test("sampling: headSampleRows returns the first N rows", () => {
  const rows = Array.from({ length: 10 }, (_v, i) => i);
  assert.deepStrictEqual(headSampleRows(rows, 0), []);
  assert.deepStrictEqual(headSampleRows(rows, 3), [0, 1, 2]);
  assert.deepStrictEqual(headSampleRows(rows, 20), rows);
});

test("sampling: tailSampleRows returns the last N rows", () => {
  const rows = Array.from({ length: 10 }, (_v, i) => i);
  assert.deepStrictEqual(tailSampleRows(rows, 0), []);
  assert.deepStrictEqual(tailSampleRows(rows, 3), [7, 8, 9]);
  assert.deepStrictEqual(tailSampleRows(rows, 20), rows);
});

test("sampling: systematicSampleRows is deterministic and evenly spaced", () => {
  const rows = Array.from({ length: 10 }, (_v, i) => i);
  const sampleA = systematicSampleRows(rows, 4, { seed: 123 });
  const sampleB = systematicSampleRows(rows, 4, { seed: 123 });

  assert.deepStrictEqual(sampleA, sampleB);
  assert.deepStrictEqual(sampleA, [1, 4, 6, 9]);
});

test("sampling: validate sampleSize / options", () => {
  const rows = Array.from({ length: 10 }, (_v, i) => i);

  assert.throws(() => headSampleRows(rows, -1), /sampleSize must be a non-negative integer/);
  assert.throws(() => tailSampleRows(rows, 1.5), /sampleSize must be a non-negative integer/);
  assert.throws(() => systematicSampleRows(rows, -1), /sampleSize must be a non-negative integer/);
  assert.throws(() => systematicSampleRows(rows, 3, { offset: Number.NaN }), /offset must be a finite number/);

  const stratifiedRows = rows.map((v) => ["S", v]);
  assert.throws(
    () => stratifiedSampleRows(stratifiedRows, 1.1, { getStratum: (r) => r[0], seed: 1 }),
    /sampleSize must be a non-negative integer/,
  );
});

test("sampling: randomSampleIndices uses O(sampleSize) RNG calls for large totals", () => {
  let calls = 0;
  const rng = () => {
    calls += 1;
    if (calls > 100) throw new Error("rng called too many times");
    return 0.123456;
  };

  const indices = randomSampleIndices(1_000_000_000, 10, rng);
  assert.equal(indices.length, 10);
  assert.equal(new Set(indices).size, 10);
  assert.ok(calls <= 100);
});

test("sampling: stratifiedSampleRows does not call rng per row", () => {
  const total = 200_000;
  const strataCount = 100;
  const keys = Array.from({ length: strataCount }, (_v, i) => `S${i}`);
  const rows = Array.from({ length: total }, (_v, i) => keys[i % strataCount]);

  let calls = 0;
  const rng = () => {
    calls += 1;
    if (calls > 10_000) throw new Error("rng called too many times");
    return 0.5;
  };

  const sampled = stratifiedSampleRows(rows, 10, { getStratum: (r) => r, rng });
  assert.equal(sampled.length, 10);
  assert.equal(new Set(sampled).size, 10);
  assert.ok(calls <= 10_000);
});

