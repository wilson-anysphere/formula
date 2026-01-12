/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

describe("SpreadsheetApp.reindexCommentCells", () => {
  it("populates a coordinate-keyed comment meta index", () => {
    const comments = [
      // Two threads on B2; resolved should aggregate to false.
      { cellRef: "B2", resolved: false },
      { cellRef: "B2", resolved: true },
      // Single resolved thread on A1.
      { cellRef: "A1", resolved: true },
    ];

    const app = Object.create(SpreadsheetApp.prototype) as SpreadsheetApp;
    (app as any).commentCells = new Set<string>();
    (app as any).commentMeta = new Map<string, { resolved: boolean }>();
    (app as any).commentMetaByCoord = new Map<number, { resolved: boolean }>();
    (app as any).commentManager = { listAll: () => comments };

    const invalidateAll = vi.fn();
    (app as any).sharedProvider = { invalidateAll };

    (app as any).reindexCommentCells();

    expect((app as any).commentCells.has("A1")).toBe(true);
    expect((app as any).commentCells.has("B2")).toBe(true);

    expect((app as any).commentMeta.get("A1")).toEqual({ resolved: true });
    expect((app as any).commentMeta.get("B2")).toEqual({ resolved: false });

    // coordKey = row * 16_384 + col (Excel max cols).
    expect((app as any).commentMetaByCoord.get(0)).toEqual({ resolved: true }); // A1 => (0,0)
    expect((app as any).commentMetaByCoord.get(1 * 16_384 + 1)).toEqual({ resolved: false }); // B2 => (1,1)

    expect(invalidateAll).toHaveBeenCalledTimes(1);
  });
});
