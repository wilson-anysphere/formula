import assert from "node:assert/strict";
import test from "node:test";

import { dedupeOverlappingResults } from "../src/retrieval/ranking.js";

test("dedupeOverlappingResults drops highly-overlapping rects within a workbook/sheet", () => {
  const results = [
    {
      id: "a",
      score: 1,
      metadata: { workbookId: "wb", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 9, c1: 0 } },
    },
    {
      id: "b",
      score: 0.9,
      // Fully contained in `a` => overlap ratio is 1.0 relative to smaller rect.
      metadata: { workbookId: "wb", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 7, c1: 0 } },
    },
  ];

  const out = dedupeOverlappingResults({ results });
  assert.deepEqual(out.map((r) => r.id), ["a"]);
});

test("dedupeOverlappingResults does not dedupe across sheets or workbooks", () => {
  const results = [
    {
      id: "a",
      score: 1,
      metadata: { workbookId: "wb1", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 7, c1: 0 } },
    },
    {
      id: "b",
      score: 0.9,
      metadata: { workbookId: "wb1", sheetName: "Sheet2", rect: { r0: 0, c0: 0, r1: 7, c1: 0 } },
    },
    {
      id: "c",
      score: 0.8,
      metadata: { workbookId: "wb2", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 7, c1: 0 } },
    },
  ];

  const out = dedupeOverlappingResults({ results });
  assert.deepEqual(out.map((r) => r.id), ["a", "b", "c"]);
});

test("dedupeOverlappingResults preserves distinct rects below the overlap threshold", () => {
  const results = [
    {
      id: "a",
      score: 1,
      metadata: { workbookId: "wb", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 7, c1: 0 } }, // 8 cells
    },
    {
      id: "b",
      score: 0.9,
      metadata: { workbookId: "wb", sheetName: "Sheet1", rect: { r0: 4, c0: 0, r1: 11, c1: 0 } }, // 8 cells, overlap=4
    },
  ];

  const out = dedupeOverlappingResults({ results, overlapRatio: 0.8 });
  // overlap ratio = 4 / min(8,8) = 0.5 => keep both
  assert.deepEqual(out.map((r) => r.id), ["a", "b"]);
});

test("dedupeOverlappingResults drops duplicate ids", () => {
  const results = [
    { id: "a", score: 1, metadata: { workbookId: "wb" } },
    { id: "a", score: 0.5, metadata: { workbookId: "wb" } },
  ];

  const out = dedupeOverlappingResults({ results });
  assert.deepEqual(out.map((r) => r.id), ["a"]);
});

