import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";

import { HashEmbedder, InMemoryVectorStore } from "../../../../../packages/ai-rag/src/index.js";
import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";

import { DocumentControllerSpreadsheetApi } from "./documentControllerSpreadsheetApi.js";

describe("ContextManager.buildWorkbookContextFromSpreadsheetApi (DocumentController adapter)", () => {
  it("indexes & retrieves workbook context with correct A1 ranges (1-based SpreadsheetApi)", async () => {
    const controller = new DocumentController();
    controller.setRangeValues("Sheet1", "A1", [
      ["Region", "Revenue"],
      ["North", 1000],
      ["South", 2000],
    ]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(controller);

    const embedder = new HashEmbedder({ dimension: 128 });
    const vectorStore = new InMemoryVectorStore({ dimension: 128 });

    const cm = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 },
    });

    const out = await cm.buildWorkbookContextFromSpreadsheetApi({
      spreadsheet,
      workbookId: "wb-doc",
      query: "revenue by region",
    });

    expect(out.promptContext).toMatch(/Data region A1:B3/i);
    expect(out.promptContext).toMatch(/Region/);
    expect(out.promptContext).toMatch(/Revenue/);
    expect(out.promptContext).toMatch(/North/);
    expect(out.promptContext).toMatch(/South/);
  });
});

