import { beforeEach, describe, expect, it, vi } from "vitest";

// Mock the exact module specifier used by AiCellFunctionEngine so we can spy on the slow path.
vi.mock("../../../../packages/security/dlp/src/selectors.js", async () => {
  const actual = await vi.importActual<any>("../../../../packages/security/dlp/src/selectors.js");
  return {
    ...actual,
    effectiveCellClassification: vi.fn(actual.effectiveCellClassification),
  };
});

import { evaluateFormula } from "./evaluateFormula.js";
import { AI_CELL_PLACEHOLDER, AiCellFunctionEngine, __dlpIndexBuilder } from "./AiCellFunctionEngine.js";

import * as selectors from "../../../../packages/security/dlp/src/selectors.js";

describe("AiCellFunctionEngine DLP indexing", () => {
  beforeEach(() => {
    globalThis.localStorage?.clear();
  });

  it("does not scan all classification records per referenced cell", async () => {
    const workbookId = "ai-cell-dlp-index";
    const sheetId = "Sheet1";

    // 100x100 = 10k per-cell classification records.
    const records = [];
    const updatedAt = new Date().toISOString();
    for (let row = 0; row < 100; row++) {
      for (let col = 0; col < 100; col++) {
        records.push({
          selector: { scope: "cell", documentId: workbookId, sheetId, row, col },
          classification: { level: row === 0 && col === 0 ? "Restricted" : "Public", labels: [] },
          updatedAt,
        });
      }
    }

    globalThis.localStorage?.setItem(`dlp:classifications:${workbookId}`, JSON.stringify(records));

    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({ llmClient: llmClient as any, workbookId, model: "test-model" });

    const effectiveCellClassification = selectors.effectiveCellClassification as unknown as ReturnType<typeof vi.fn>;
    effectiveCellClassification.mockClear();

    const getCellValue = (addr: string) => (addr === "A1" ? "top secret" : null);
    const value = evaluateFormula('=AI("summarize", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });

    expect(value).toBe(AI_CELL_PLACEHOLDER);
    // Perf proxy: we should use the precomputed index, not scan all records per cell.
    expect(effectiveCellClassification).toHaveBeenCalledTimes(0);

    await engine.waitForIdle();
    expect(llmClient.chat).toHaveBeenCalledTimes(1);
  });

  it("memoizes DLP cell indices across evaluations and invalidates on record changes", async () => {
    const workbookId = "ai-cell-dlp-index-memo";
    const sheetId = "Sheet1";
    const updatedAt = new Date().toISOString();

    const records = [
      {
        selector: { scope: "cell", documentId: workbookId, sheetId, row: 0, col: 0 },
        classification: { level: "Restricted", labels: [] },
        updatedAt,
      },
      {
        selector: { scope: "cell", documentId: workbookId, sheetId, row: 0, col: 1 },
        classification: { level: "Public", labels: [] },
        updatedAt,
      },
    ];

    globalThis.localStorage?.setItem(`dlp:classifications:${workbookId}`, JSON.stringify(records));

    const llmClient = {
      chat: vi.fn(async () => ({
        message: { role: "assistant", content: "ok" },
        usage: { promptTokens: 1, completionTokens: 1 },
      })),
    };

    const engine = new AiCellFunctionEngine({ llmClient: llmClient as any, workbookId, model: "test-model" });

    const buildSpy = vi.spyOn(__dlpIndexBuilder, "buildDlpCellIndex");

    const getCellValue = (addr: string) => (addr === "A1" ? "top secret" : null);

    const v1 = evaluateFormula('=AI("one", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!B1" });
    expect(v1).toBe(AI_CELL_PLACEHOLDER);
    expect(buildSpy).toHaveBeenCalledTimes(1);

    const v2 = evaluateFormula('=AI("two", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!C1" });
    expect(v2).toBe(AI_CELL_PLACEHOLDER);
    expect(buildSpy).toHaveBeenCalledTimes(1);

    await engine.waitForIdle();
    expect(llmClient.chat).toHaveBeenCalledTimes(2);

    const next = [
      ...records,
      {
        selector: { scope: "cell", documentId: workbookId, sheetId, row: 1, col: 0 },
        classification: { level: "Public", labels: [] },
        updatedAt,
      },
    ];
    globalThis.localStorage?.setItem(`dlp:classifications:${workbookId}`, JSON.stringify(next));

    const v3 = evaluateFormula('=AI("three", A1)', getCellValue, { ai: engine, cellAddress: "Sheet1!D1" });
    expect(v3).toBe(AI_CELL_PLACEHOLDER);
    expect(buildSpy).toHaveBeenCalledTimes(2);

    await engine.waitForIdle();
    expect(llmClient.chat).toHaveBeenCalledTimes(3);

    buildSpy.mockRestore();
  });
});
