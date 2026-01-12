import { describe, expect, it, vi } from "vitest";

// Mock the exact module specifier used by AiCellFunctionEngine so we can spy on the slow path.
vi.mock("../../../../packages/security/dlp/src/selectors.js", async () => {
  const actual = await vi.importActual<any>("../../../../packages/security/dlp/src/selectors.js");
  return {
    ...actual,
    effectiveCellClassification: vi.fn(actual.effectiveCellClassification),
  };
});

import { evaluateFormula } from "./evaluateFormula.js";
import { AI_CELL_PLACEHOLDER, AiCellFunctionEngine } from "./AiCellFunctionEngine.js";

import * as selectors from "../../../../packages/security/dlp/src/selectors.js";

describe("AiCellFunctionEngine DLP indexing", () => {
  it("does not scan all classification records per referenced cell", async () => {
    globalThis.localStorage?.clear();

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
});

