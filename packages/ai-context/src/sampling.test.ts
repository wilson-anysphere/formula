import { describe, expect, it } from "vitest";

import { headSampleRows, randomSampleRows, stratifiedSampleRows, systematicSampleRows, tailSampleRows } from "./sampling.js";

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
    const rows = Array.from({ length: total }, (_v, i) => [keys[i % strataCount], i]);

    const sampledA = stratifiedSampleRows(rows, 10, { getStratum: (r: any) => r[0], seed: 123 });
    const sampledB = stratifiedSampleRows(rows, 10, { getStratum: (r: any) => r[0], seed: 123 });

    expect(sampledA).toEqual(sampledB);
    expect(sampledA).toHaveLength(10);
    expect(new Set(sampledA.map((r: any) => r[0])).size).toBe(10);
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
});
