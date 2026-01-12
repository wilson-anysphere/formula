import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { HashEmbedder } from "../../../../../packages/ai-rag/src/index.js";

import { createDesktopRagService } from "./ragService.js";
import { DocumentControllerSpreadsheetApi } from "../tools/documentControllerSpreadsheetApi.js";

describe("createDesktopRagService (embedder config)", () => {
  it("uses HashEmbedder by default", async () => {
    const controller = new DocumentController();
    let observedEmbedder: unknown = null;

    const service = createDesktopRagService({
      documentController: controller,
      workbookId: "wb_embedder_default",
      createRag: async (opts: any) => {
        observedEmbedder = opts?.embedder;
        return {
          vectorStore: { close: async () => {} },
          contextManager: {},
          indexWorkbook: async () => ({ indexed: 0 }),
        } as any;
      },
    });

    try {
      await service.getContextManager();
      expect(observedEmbedder).toBeInstanceOf(HashEmbedder);
    } finally {
      await service.dispose();
    }
  });

  it("rejects non-hash embedder types", () => {
    const controller = new DocumentController();

    expect(() =>
      createDesktopRagService({
        documentController: controller,
        workbookId: "wb_embedder_type_reject",
        embedder: { type: "unsupported" } as any,
      }),
    ).toThrow(/only supports deterministic hash embeddings/i);
  });

  it("accepts hash embedder config", async () => {
    const controller = new DocumentController();
    const service = createDesktopRagService({
      documentController: controller,
      workbookId: "wb_embedder_type_accept",
      embedder: { type: "hash", dimension: 32 },
    });

    await service.dispose();
  });

  it("does not re-index when only sheet view changes (contentVersion)", async () => {
    const controller = new DocumentController();
    controller.setRangeValues("Sheet1", "A1", [["Header"], ["Value"]]);

    const spreadsheet = new DocumentControllerSpreadsheetApi(controller);

    let indexCalls = 0;

    const service = createDesktopRagService({
      documentController: controller,
      workbookId: "wb_rag_version",
      // Test seam: avoid sqlite/localstorage and just count re-indexing.
      createRag: async () =>
        ({
          vectorStore: { close: async () => {} },
          contextManager: {
            buildWorkbookContextFromSpreadsheetApi: async () => ({ promptContext: "", retrieved: [], indexStats: null }),
          },
          indexWorkbook: async () => {
            indexCalls += 1;
            return { indexed: indexCalls };
          },
        }) as any,
    });

    await service.buildWorkbookContextFromSpreadsheetApi({
      spreadsheet,
      workbookId: "wb_rag_version",
      query: "hello",
    });
    expect(indexCalls).toBe(1);

    // Sheet view only: should not bump DocumentController.contentVersion, so the index stays fresh.
    controller.setFrozen("Sheet1", 1, 0);

    await service.buildWorkbookContextFromSpreadsheetApi({
      spreadsheet,
      workbookId: "wb_rag_version",
      query: "hello again",
    });
    expect(indexCalls).toBe(1);

    // Content change should invalidate the index.
    controller.setCellValue("Sheet1", "A2", "changed");

    await service.buildWorkbookContextFromSpreadsheetApi({
      spreadsheet,
      workbookId: "wb_rag_version",
      query: "hello again",
    });
    expect(indexCalls).toBe(2);

    await service.dispose();
  });
});
