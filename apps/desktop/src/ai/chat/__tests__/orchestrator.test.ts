import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../../document/documentController.js";

import { LocalStorageAIAuditStore } from "../../../../../../packages/ai-audit/src/local-storage-store.js";
import { ContextManager } from "../../../../../../packages/ai-context/src/contextManager.js";
import { HashEmbedder, InMemoryVectorStore } from "../../../../../../packages/ai-rag/src/index.js";

import { createAiChatOrchestrator } from "../orchestrator.js";

function seed2x2(controller: DocumentController) {
  controller.setCellValue("Sheet1", "A1", 1);
  controller.setCellValue("Sheet1", "B1", 2);
  controller.setCellValue("Sheet1", "A2", 3);
  controller.setCellValue("Sheet1", "B2", 4);
}

function createMockLlmClient(params: { cell: string; value: unknown }) {
  const requests: any[] = [];
  let callCount = 0;
  return {
    requests,
    client: {
      async chat(request: any) {
        requests.push(request);
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [
                {
                  id: "call_1",
                  name: "write_cell",
                  arguments: { cell: params.cell, value: params.value },
                },
              ],
            },
            usage: { promptTokens: 10, completionTokens: 5 },
          };
        }
        return {
          message: { role: "assistant", content: "ok" },
          usage: { promptTokens: 5, completionTokens: 3 },
        };
      },
    },
  };
}

describe("ai chat orchestrator", () => {
  it("denies tool calls when preview requires approval and no approval callback is provided", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const embedder = new HashEmbedder({ dimension: 64 });
    const vectorStore = new InMemoryVectorStore({ dimension: 64 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 },
    });

    const buildContextSpy = vi.spyOn(contextManager as any, "buildWorkbookContextFromSpreadsheetApi");

    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_denied" });
    const mock = createMockLlmClient({ cell: "A1", value: 99 });

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_denied",
      llmClient: mock.client as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_denied",
      contextManager,
      previewOptions: { approval_cell_threshold: 0 },
    });

    await expect(
      orchestrator.sendMessage({
        text: "Set A1 to 99",
        history: [],
      }),
    ).rejects.toThrow(/denied/i);

    expect(controller.getCell("Sheet1", "A1").value).toBe(1);

    expect(buildContextSpy).toHaveBeenCalledTimes(1);

    const firstRequest = mock.requests[0];
    expect(firstRequest.messages?.[0]?.role).toBe("system");
    expect(firstRequest.messages?.[0]?.content).toContain("WORKBOOK_CONTEXT");
    expect(firstRequest.messages?.[0]?.content).toContain("Workbook summary");

    const entries = await auditStore.listEntries({ session_id: "session_denied" });
    expect(entries.length).toBe(1);
    expect(entries[0]?.mode).toBe("chat");
    expect(entries[0]?.model).toBe("mock-model");
    expect(entries[0]?.tool_calls?.[0]?.name).toBe("write_cell");
    expect(entries[0]?.tool_calls?.[0]?.approved).toBe(false);
  });

  it("executes tool calls when approval callback approves the preview", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const embedder = new HashEmbedder({ dimension: 64 });
    const vectorStore = new InMemoryVectorStore({ dimension: 64 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 },
    });

    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_approved" });
    const mock = createMockLlmClient({ cell: "A1", value: 99 });

    const onApprovalRequired = vi.fn(async () => true);

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_approved",
      llmClient: mock.client as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_approved",
      contextManager,
      onApprovalRequired,
      previewOptions: { approval_cell_threshold: 0 },
    });

    const result = await orchestrator.sendMessage({
      text: "Set A1 to 99",
      history: [],
    });

    expect(result.finalText).toBe("ok");
    expect(result.toolResults.length).toBe(1);
    expect(result.toolResults[0]?.ok).toBe(true);
    expect(controller.getCell("Sheet1", "A1").value).toBe(99);

    expect(onApprovalRequired).toHaveBeenCalledTimes(1);

    const entries = await auditStore.listEntries({ session_id: "session_approved" });
    expect(entries.length).toBe(1);
    expect(entries[0]?.tool_calls?.[0]?.approved).toBe(true);

    const firstRequest = mock.requests[0];
    expect(firstRequest.messages?.[0]?.content).toContain("Workbook summary");
  });
});

