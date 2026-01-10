import { describe, expect, it } from "vitest";

import { randomSampleRows, stratifiedSampleRows } from "./sampling.js";

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
});

