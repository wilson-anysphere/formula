import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";

import { createDesktopRagService } from "./ragService.js";

describe("createDesktopRagService (embedder config)", () => {
  it("rejects non-hash embedder types", () => {
    const controller = new DocumentController();

    expect(() =>
      createDesktopRagService({
        documentController: controller,
        workbookId: "wb_embedder_type_reject",
        embedder: { type: "openai" } as any,
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
});

