/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { SpreadsheetApp } from "../spreadsheetApp";

describe("SpreadsheetApp.reindexCommentCells", () => {
  it("populates a coordinate-keyed comment meta index", () => {
    const comments = [
      // Two threads on B2; resolved should aggregate to false.
      { cellRef: "B2", resolved: false, content: "First B2" },
      { cellRef: "B2", resolved: true, content: "Second B2" },
      // Single resolved thread on A1.
      { cellRef: "A1", resolved: true, content: "A1 note" },
    ];

    const app = Object.create(SpreadsheetApp.prototype) as SpreadsheetApp;
    (app as any).commentCells = new Set<string>();
    (app as any).commentMeta = new Map<string, { resolved: boolean }>();
    (app as any).commentMetaByCoord = new Map<number, { resolved: boolean }>();
    (app as any).commentPreviewByCoord = new Map<number, string>();
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
    expect((app as any).commentPreviewByCoord.get(0)).toBe("A1 note");
    expect((app as any).commentPreviewByCoord.get(1 * 16_384 + 1)).toBe("First B2");

    expect(invalidateAll).toHaveBeenCalledTimes(1);
  });

  it("does not populate coord indexes for non-A1 cellRefs", () => {
    const comments = [{ cellRef: "Sheet1!A1", resolved: false, content: "Bad ref" }];

    const app = Object.create(SpreadsheetApp.prototype) as SpreadsheetApp;
    (app as any).commentCells = new Set<string>();
    (app as any).commentMeta = new Map<string, { resolved: boolean }>();
    (app as any).commentMetaByCoord = new Map<number, { resolved: boolean }>();
    (app as any).commentPreviewByCoord = new Map<number, string>();
    (app as any).commentManager = { listAll: () => comments };

    const invalidateAll = vi.fn();
    (app as any).sharedProvider = { invalidateAll };

    (app as any).reindexCommentCells();

    expect((app as any).commentCells.has("Sheet1!A1")).toBe(true);
    expect((app as any).commentMeta.get("Sheet1!A1")).toEqual({ resolved: false });

    expect((app as any).commentMetaByCoord.size).toBe(0);
    expect((app as any).commentPreviewByCoord.size).toBe(0);

    expect(invalidateAll).toHaveBeenCalledTimes(1);
  });

  it("supports absolute A1 refs (with $ markers) for coord indexing", () => {
    const comments = [{ cellRef: "$A$1", resolved: true, content: "Absolute" }];

    const app = Object.create(SpreadsheetApp.prototype) as SpreadsheetApp;
    (app as any).commentCells = new Set<string>();
    (app as any).commentMeta = new Map<string, { resolved: boolean }>();
    (app as any).commentMetaByCoord = new Map<number, { resolved: boolean }>();
    (app as any).commentPreviewByCoord = new Map<number, string>();
    (app as any).commentManager = { listAll: () => comments };

    const invalidateAll = vi.fn();
    (app as any).sharedProvider = { invalidateAll };

    (app as any).reindexCommentCells();

    // $A$1 should map to (0,0)
    expect((app as any).commentMetaByCoord.get(0)).toEqual({ resolved: true });
    expect((app as any).commentPreviewByCoord.get(0)).toBe("Absolute");
  });
});
