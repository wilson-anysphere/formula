import assert from "node:assert/strict";
import test from "node:test";

import { dedupeOverlappingResults, rerankWorkbookResults } from "../src/retrieval/rankResults.js";

test("rerankWorkbookResults boosts table + namedRange above dataRegion when scores are similar", () => {
  const query = "unrelated query";
  const results = [
    {
      id: "data",
      score: 0.5,
      metadata: {
        kind: "dataRegion",
        title: "Data region A1:C3",
        sheetName: "Sheet1",
        rect: { r0: 0, c0: 0, r1: 2, c1: 2 },
        tokenCount: 120,
      },
    },
    {
      id: "table",
      score: 0.495,
      metadata: {
        kind: "table",
        title: "SalesByRegion",
        sheetName: "Sheet1",
        rect: { r0: 0, c0: 0, r1: 2, c1: 2 },
        tokenCount: 120,
      },
    },
    {
      id: "named",
      score: 0.494,
      metadata: {
        kind: "namedRange",
        title: "SummaryMetrics",
        sheetName: "Sheet1",
        rect: { r0: 5, c0: 0, r1: 6, c1: 1 },
        tokenCount: 120,
      },
    },
  ];

  const reranked = rerankWorkbookResults(query, results);
  assert.deepEqual(
    reranked.map((r) => r.metadata.kind),
    ["table", "namedRange", "dataRegion"]
  );
});

test("rerankWorkbookResults uses deterministic ordering when adjusted scores tie", () => {
  const query = "foo bar";
  const results = [
    {
      id: "b",
      score: 0.5,
      metadata: { kind: "dataRegion", title: "X", sheetName: "Sheet1", tokenCount: 10 },
    },
    {
      id: "a",
      score: 0.5,
      metadata: { kind: "dataRegion", title: "X", sheetName: "Sheet1", tokenCount: 10 },
    },
  ];

  const reranked = rerankWorkbookResults(query, results);
  assert.deepEqual(reranked.map((r) => r.id), ["a", "b"]);
});

test("rerankWorkbookResults boosts results whose title/sheetName match query tokens", () => {
  const query = "revenue";
  const results = [
    {
      id: "no-match",
      score: 0.5,
      metadata: { kind: "dataRegion", title: "Costs", sheetName: "Sheet1", tokenCount: 10 },
    },
    {
      id: "match-title",
      score: 0.5,
      metadata: { kind: "dataRegion", title: "RevenueSummary", sheetName: "Sheet1", tokenCount: 10 },
    },
  ];

  const reranked = rerankWorkbookResults(query, results);
  assert.deepEqual(reranked.map((r) => r.id), ["match-title", "no-match"]);
});

test("rerankWorkbookResults tokenizes underscores/camelCase for lexical matching", () => {
  const query = "revenue_by_region";
  const results = [
    {
      id: "no-match",
      score: 0.5,
      metadata: { kind: "dataRegion", title: "Costs", sheetName: "Sheet1", tokenCount: 10 },
    },
    {
      id: "match",
      score: 0.5,
      metadata: { kind: "dataRegion", title: "RevenueByRegion", sheetName: "Sheet1", tokenCount: 10 },
    },
  ];

  const reranked = rerankWorkbookResults(query, results);
  assert.deepEqual(reranked.map((r) => r.id), ["match", "no-match"]);
});

test("rerankWorkbookResults tokenizes identifier-style queries (camelCase/PascalCase)", () => {
  const query = "RevenueByRegion";
  const results = [
    {
      id: "a",
      score: 0.5,
      metadata: { kind: "dataRegion", title: "Salaries", sheetName: "Sheet1", tokenCount: 10 },
    },
    {
      id: "b",
      score: 0.5,
      metadata: { kind: "dataRegion", title: "Revenue By Region", sheetName: "Sheet1", tokenCount: 10 },
    },
  ];

  const reranked = rerankWorkbookResults(query, results);
  // Token matches in the title should boost the relevant result above the
  // unrelated one even when base scores tie.
  assert.deepEqual(reranked.map((r) => r.id), ["b", "a"]);
});

test("rerankWorkbookResults penalizes very large chunks to favor concise context", () => {
  const query = "";
  const results = [
    {
      id: "large",
      score: 0.51,
      metadata: { kind: "dataRegion", title: "A", sheetName: "Sheet1", tokenCount: 10_000 },
    },
    {
      id: "small",
      score: 0.5,
      metadata: { kind: "dataRegion", title: "B", sheetName: "Sheet1", tokenCount: 50 },
    },
  ];

  const reranked = rerankWorkbookResults(query, results);
  assert.deepEqual(reranked.map((r) => r.id), ["small", "large"]);
});

test("dedupeOverlappingResults removes near-duplicate overlapping chunks (keeps highest score)", () => {
  const results = [
    // Lower-score duplicate comes first intentionally.
    {
      id: "low",
      score: 0.8,
      metadata: { workbookId: "wb", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 9, c1: 9 } },
    },
    {
      id: "high",
      score: 0.9,
      metadata: { workbookId: "wb", sheetName: "Sheet1", rect: { r0: 1, c0: 1, r1: 8, c1: 8 } },
    },
    {
      id: "other",
      score: 0.85,
      metadata: { workbookId: "wb", sheetName: "Sheet1", rect: { r0: 20, c0: 0, r1: 21, c1: 2 } },
    },
    // Same rect but different sheet -> should not dedupe with Sheet1 chunks.
    {
      id: "sheet2",
      score: 0.95,
      metadata: { workbookId: "wb", sheetName: "Sheet2", rect: { r0: 1, c0: 1, r1: 8, c1: 8 } },
    },
  ];

  const deduped = dedupeOverlappingResults(results);
  assert.deepEqual(deduped.map((r) => r.id), ["sheet2", "high", "other"]);
});
