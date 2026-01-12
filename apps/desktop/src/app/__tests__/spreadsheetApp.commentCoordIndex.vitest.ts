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
    (app as any).commentMetaByCoord = new Map<number, { resolved: boolean }>();
    (app as any).commentPreviewByCoord = new Map<number, string>();
    (app as any).commentThreadsByCellRef = new Map<string, any[]>();
    (app as any).commentIndexVersion = 0;
    (app as any).lastHoveredCommentIndexVersion = -1;
    (app as any).commentManager = { listAll: () => comments };

    const invalidateAll = vi.fn();
    (app as any).sharedProvider = { invalidateAll };

    (app as any).reindexCommentCells();

    expect((app as any).commentThreadsByCellRef.get("A1")).toHaveLength(1);
    expect((app as any).commentThreadsByCellRef.get("B2")).toHaveLength(2);

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
    (app as any).commentMetaByCoord = new Map<number, { resolved: boolean }>();
    (app as any).commentPreviewByCoord = new Map<number, string>();
    (app as any).commentThreadsByCellRef = new Map<string, any[]>();
    (app as any).commentIndexVersion = 0;
    (app as any).lastHoveredCommentIndexVersion = -1;
    (app as any).commentManager = { listAll: () => comments };

    const invalidateAll = vi.fn();
    (app as any).sharedProvider = { invalidateAll };

    (app as any).reindexCommentCells();

    expect((app as any).commentMetaByCoord.size).toBe(0);
    expect((app as any).commentPreviewByCoord.size).toBe(0);
    expect((app as any).commentThreadsByCellRef.get("Sheet1!A1")).toHaveLength(1);

    expect(invalidateAll).toHaveBeenCalledTimes(1);
  });

  it("supports absolute A1 refs (with $ markers) for coord indexing", () => {
    const comments = [{ cellRef: "$A$1", resolved: true, content: "Absolute" }];

    const app = Object.create(SpreadsheetApp.prototype) as SpreadsheetApp;
    (app as any).commentMetaByCoord = new Map<number, { resolved: boolean }>();
    (app as any).commentPreviewByCoord = new Map<number, string>();
    (app as any).commentThreadsByCellRef = new Map<string, any[]>();
    (app as any).commentIndexVersion = 0;
    (app as any).lastHoveredCommentIndexVersion = -1;
    (app as any).commentManager = { listAll: () => comments };

    const invalidateAll = vi.fn();
    (app as any).sharedProvider = { invalidateAll };

    (app as any).reindexCommentCells();

    // $A$1 should map to (0,0)
    expect((app as any).commentMetaByCoord.get(0)).toEqual({ resolved: true });
    expect((app as any).commentPreviewByCoord.get(0)).toBe("Absolute");
    // A1-keyed maps should normalize away $ markers.
    expect((app as any).commentThreadsByCellRef.get("A1")).toHaveLength(1);
  });

  it("indexes only the active sheet in collab mode (sheet-qualified refs)", () => {
    const comments = [
      { cellRef: "Sheet1!A1", resolved: false, content: "S1" },
      { cellRef: "Sheet2!A1", resolved: true, content: "S2" },
    ];

    const app = Object.create(SpreadsheetApp.prototype) as SpreadsheetApp;
    (app as any).sheetId = "Sheet1";
    (app as any).collabMode = true;
    (app as any).commentMetaByCoord = new Map<number, { resolved: boolean }>();
    (app as any).commentPreviewByCoord = new Map<number, string>();
    (app as any).commentThreadsByCellRef = new Map<string, any[]>();
    (app as any).commentIndexVersion = 0;
    (app as any).lastHoveredCommentIndexVersion = -1;
    (app as any).commentManager = { listAll: () => comments };

    const invalidateAll = vi.fn();
    (app as any).sharedProvider = { invalidateAll };

    (app as any).reindexCommentCells();

    // Both sheets appear in the A1-keyed thread index for comments panel lookups.
    expect((app as any).commentThreadsByCellRef.get("Sheet1!A1")).toHaveLength(1);
    expect((app as any).commentThreadsByCellRef.get("Sheet2!A1")).toHaveLength(1);

    // Coord indexes are used for fast rendering/tooltips and are scoped to the active sheet.
    expect((app as any).commentMetaByCoord.size).toBe(1);
    expect((app as any).commentMetaByCoord.get(0)).toEqual({ resolved: false });
    expect((app as any).commentPreviewByCoord.get(0)).toBe("S1");
  });

  it("does not call listAll (and avoids instantiating comments root) before collab provider sync when comments root is absent", () => {
    const listAll = vi.fn(() => []);

    const app = Object.create(SpreadsheetApp.prototype) as SpreadsheetApp;
    (app as any).collabMode = true;
    // Emulate a collab session that has not synced yet and has no `comments` root in the doc.
    (app as any).collabSession = { provider: { synced: false, on: () => {} }, doc: { share: new Map() } };
    (app as any).commentMetaByCoord = new Map<number, { resolved: boolean }>();
    (app as any).commentPreviewByCoord = new Map<number, string>();
    (app as any).commentThreadsByCellRef = new Map<string, any[]>();
    (app as any).commentIndexVersion = 0;
    (app as any).lastHoveredCommentIndexVersion = -1;
    (app as any).commentManager = { listAll };

    const invalidateAll = vi.fn();
    (app as any).sharedProvider = { invalidateAll };

    (app as any).reindexCommentCells();

    expect(listAll).toHaveBeenCalledTimes(0);
    expect((app as any).commentThreadsByCellRef.size).toBe(0);
    expect((app as any).commentMetaByCoord.size).toBe(0);
    expect((app as any).commentPreviewByCoord.size).toBe(0);
    expect((app as any).commentIndexVersion).toBe(1);
    expect(invalidateAll).toHaveBeenCalledTimes(1);
  });
});
