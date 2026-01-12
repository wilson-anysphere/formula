import { beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../../document/documentController.js";

const { createDesktopRagServiceSpy, ownedRagServiceDisposeSpy } = vi.hoisted(() => {
  const ownedRagServiceDisposeSpy = vi.fn(async () => {});
  const createDesktopRagServiceSpy = vi.fn(() => ({
    dispose: ownedRagServiceDisposeSpy,
  }));

  return { createDesktopRagServiceSpy, ownedRagServiceDisposeSpy };
});

vi.mock("../../rag/ragService.js", () => ({
  createDesktopRagService: createDesktopRagServiceSpy,
}));

import { createAiChatOrchestrator } from "../orchestrator.js";

describe("ai chat orchestrator disposal", () => {
  beforeEach(() => {
    createDesktopRagServiceSpy.mockClear();
    ownedRagServiceDisposeSpy.mockClear();
  });

  it("disposes the internally created DesktopRagService (idempotent)", async () => {
    const controller = new DocumentController();

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_dispose_owned",
      llmClient: { chat: vi.fn(async () => ({ message: { role: "assistant", content: "ok" } })) } as any,
      model: "mock-model",
    });

    await orchestrator.dispose();
    await orchestrator.dispose();

    expect(createDesktopRagServiceSpy).toHaveBeenCalledTimes(1);
    expect(ownedRagServiceDisposeSpy).toHaveBeenCalledTimes(1);
  });

  it("does not dispose a caller-provided ragService", async () => {
    const controller = new DocumentController();

    const ragService = {
      dispose: vi.fn(async () => {}),
      buildWorkbookContextFromSpreadsheetApi: vi.fn(async () => ({})),
      getContextManager: vi.fn(async () => ({})),
    };

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_dispose_external",
      llmClient: { chat: vi.fn(async () => ({ message: { role: "assistant", content: "ok" } })) } as any,
      model: "mock-model",
      ragService: ragService as any,
    });

    await orchestrator.dispose();

    expect(createDesktopRagServiceSpy).not.toHaveBeenCalled();
    expect(ragService.dispose).not.toHaveBeenCalled();
  });
});

