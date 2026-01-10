import test from "node:test";
import assert from "node:assert/strict";
import { randomSampleRows, stratifiedSampleRows } from "./sampling.js";

test("randomSampleRows is deterministic with a seed and returns unique rows", () => {
  const rows = Array.from({ length: 10 }, (_, i) => i + 1);
  const sampleA = randomSampleRows(rows, 4, { seed: 123 });
  const sampleB = randomSampleRows(rows, 4, { seed: 123 });
  assert.deepEqual(sampleA, sampleB);
  assert.equal(new Set(sampleA).size, 4);
  assert.ok(sampleA.every((v) => rows.includes(v)));
});

test("stratifiedSampleRows includes at least one row per stratum when possible", () => {
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
  assert.deepEqual(new Set(sampled.map((r) => r[0])), new Set(["A", "B", "C"]));
});

test("stratifiedSampleRows allocates extra samples to larger strata", () => {
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
    acc[r[0]] = (acc[r[0]] ?? 0) + 1;
    return acc;
  }, /** @type {Record<string, number>} */ ({}));

  assert.equal(sampled.length, 4);
  assert.ok(counts.B >= 2);
  assert.ok(counts.A >= 1);
  assert.ok(counts.C >= 1);
});
