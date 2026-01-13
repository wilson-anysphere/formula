import { describe, expect, it } from "vitest";

import {
  headSampleRows,
  randomSampleIndices,
  randomSampleRows,
  stratifiedSampleRows,
  systematicSampleRows,
  tailSampleRows,
} from "./sampling.js";

describe("sampling", () => {
  it("randomSampleRows is deterministic with a seed and returns unique rows", () => {
    const rows = Array.from({ length: 10 }, (_, i) => i + 1);
    const sampleA = randomSampleRows(rows, 4, { seed: 123 });
    const sampleB = randomSampleRows(rows, 4, { seed: 123 });
    expect(sampleA).toEqual(sampleB);
    expect(new Set(sampleA).size).toBe(4);
    expect(sampleA.every((v: any) => rows.includes(v))).toBe(true);
  });

  it("stratifiedSampleRows includes at least one row per stratum when possible", () => {
    const rows = [
      ["A", 1],
      ["A", 2],
      ["B", 3],
      ["B", 4],
      ["B", 5],
      ["C", 6]
    ];

    const sampled = stratifiedSampleRows(rows, 3, { getStratum: (r: any) => r[0], seed: 42 });
    expect(sampled).toHaveLength(3);
    expect(new Set(sampled.map((r: any) => r[0]))).toEqual(new Set(["A", "B", "C"]));
  });

  it("stratifiedSampleRows allocates extra samples to larger strata", () => {
    const rows = [
      ["A", 1],
      ["A", 2],
      ["B", 3],
      ["B", 4],
      ["B", 5],
      ["C", 6]
    ];

    const sampled = stratifiedSampleRows(rows, 4, { getStratum: (r: any) => r[0], seed: 7 });
    const counts = sampled.reduce((acc: Record<string, number>, r: any) => {
      acc[r[0]] = (acc[r[0]] ?? 0) + 1;
      return acc;
    }, {});

    expect(sampled).toHaveLength(4);
    expect(counts.B).toBeGreaterThanOrEqual(2);
    expect(counts.A).toBeGreaterThanOrEqual(1);
    expect(counts.C).toBeGreaterThanOrEqual(1);
  });

  it("stratifiedSampleRows handles large row sets without O(N) weighted helper arrays", () => {
    const total = 200_000;
    const strataCount = 100;
    const keys = Array.from({ length: strataCount }, (_v, i) => `S${i}`);
    // Keep the synthetic rows extremely small so this test stays fast and memory-light.
    // Each row is just a reference to one of `keys`.
    const rows = Array.from({ length: total }, (_v, i) => keys[i % strataCount]);

    const sampledA = stratifiedSampleRows(rows, 10, { getStratum: (r: any) => r, seed: 123 });
    const sampledB = stratifiedSampleRows(rows, 10, { getStratum: (r: any) => r, seed: 123 });

    expect(sampledA).toEqual(sampledB);
    expect(sampledA).toHaveLength(10);
    expect(new Set(sampledA).size).toBe(10);
  });

  it("headSampleRows returns the first N rows", () => {
    const rows = Array.from({ length: 10 }, (_v, i) => i);
    expect(headSampleRows(rows, 0)).toEqual([]);
    expect(headSampleRows(rows, 3)).toEqual([0, 1, 2]);
    expect(headSampleRows(rows, 20)).toEqual(rows);
  });

  it("tailSampleRows returns the last N rows", () => {
    const rows = Array.from({ length: 10 }, (_v, i) => i);
    expect(tailSampleRows(rows, 0)).toEqual([]);
    expect(tailSampleRows(rows, 3)).toEqual([7, 8, 9]);
    expect(tailSampleRows(rows, 20)).toEqual(rows);
  });

  it("systematicSampleRows is deterministic and evenly spaced", () => {
    const rows = Array.from({ length: 10 }, (_v, i) => i);
    const sampleA = systematicSampleRows(rows, 4, { seed: 123 });
    const sampleB = systematicSampleRows(rows, 4, { seed: 123 });

    expect(sampleA).toEqual(sampleB);
    expect(sampleA).toEqual([1, 4, 6, 9]);
  });

  it("sampling helpers validate sampleSize / options", () => {
    const rows = Array.from({ length: 10 }, (_v, i) => i);

    expect(() => headSampleRows(rows, -1 as any)).toThrow(/sampleSize must be a non-negative integer/);
    expect(() => tailSampleRows(rows, 1.5 as any)).toThrow(/sampleSize must be a non-negative integer/);
    expect(() => systematicSampleRows(rows, -1 as any)).toThrow(/sampleSize must be a non-negative integer/);
    expect(() => systematicSampleRows(rows, 3, { offset: Number.NaN as any })).toThrow(/offset must be a finite number/);

    const stratifiedRows = rows.map((v) => ["S", v]);
    expect(() =>
      stratifiedSampleRows(stratifiedRows, 1.1 as any, { getStratum: (r: any) => r[0], seed: 1 }),
    ).toThrow(/sampleSize must be a non-negative integer/);
  });

  it("randomSampleIndices uses O(sampleSize) RNG calls for large totals", () => {
    let calls = 0;
    const rng = () => {
      calls += 1;
      if (calls > 100) throw new Error("rng called too many times");
      return 0.123456;
    };

    const indices = randomSampleIndices(1_000_000_000, 10, rng);
    expect(indices).toHaveLength(10);
    expect(new Set(indices).size).toBe(10);
    expect(calls).toBeLessThanOrEqual(100);
  });

  it("stratifiedSampleRows does not call rng per row", () => {
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

    const sampled = stratifiedSampleRows(rows, 10, { getStratum: (r: any) => r, rng });
    expect(sampled).toHaveLength(10);
    expect(new Set(sampled).size).toBe(10);
    expect(calls).toBeLessThanOrEqual(10_000);
  });
});
