import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";

import { MemoryAIAuditStore } from "../../../../../packages/ai-audit/src/memory-store.js";
import { ContextManager } from "../../../../../packages/ai-context/src/contextManager.js";
import { HashEmbedder } from "../../../../../packages/ai-rag/src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../../../../../packages/ai-rag/src/store/inMemoryVectorStore.js";

import { createAiChatOrchestrator } from "./orchestrator.js";

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
          message: {
            role: "assistant",
            content: "ok",
          },
          usage: { promptTokens: 2, completionTokens: 3 },
        };
      },
    },
  };
}

describe("ai chat orchestrator (desktop integration)", () => {
  it("builds workbook context + executes approved tool calls against DocumentController", async () => {
    const controller = new DocumentController();
    controller.setRangeValues("Sheet1", "A1", [
      ["Region", "Revenue"],
      ["North", 1000],
      ["South", 2000],
    ]);

    const embedder = new HashEmbedder({ dimension: 128 });
    const vectorStore = new InMemoryVectorStore({ dimension: 128 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 },
    });
 
    const auditStore = new MemoryAIAuditStore();
    const mock = createMockLlmClient({ cell: "C1", value: 99 });
    const onApprovalRequired = vi.fn(async () => true);
    const onToolCall = vi.fn();
    const onToolResult = vi.fn();

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_test",
      llmClient: mock.client as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_test",
      contextManager,
      onApprovalRequired,
      previewOptions: { approval_cell_threshold: 0 },
    });

    const result = await orchestrator.sendMessage({
      text: "Set C1 to 99",
      attachments: [{ type: "range", reference: "Sheet1!A1:B3" }],
      history: [],
      onToolCall: onToolCall as any,
      onToolResult: onToolResult as any,
    });

    expect(result.finalText).toBe("ok");
    // Ensure the sheet still has the seeded data and that the tool call updated the requested cell.
    expect(controller.getCell("Sheet1", "A1").value).toBe("Region");
    expect(controller.getCell("Sheet1", "A2").value).toBe("North");
    expect(controller.getCell("Sheet1", "C1").value).toBe(99);

    expect(onApprovalRequired).toHaveBeenCalledTimes(1);
    expect(onToolCall).toHaveBeenCalledTimes(1);
    expect(onToolResult).toHaveBeenCalledTimes(1);
    expect(onToolCall.mock.calls[0]?.[1]?.requiresApproval).toBe(true);
    expect(onToolResult.mock.calls[0]?.[1]?.ok).toBe(true);

    const firstRequest = mock.requests[0];
    expect(firstRequest.messages?.[0]?.role).toBe("system");
    expect(firstRequest.messages?.[0]?.content).toContain("WORKBOOK_CONTEXT");
    expect(firstRequest.messages?.[0]?.content).toContain("Workbook summary");
    expect(firstRequest.messages?.[0]?.content).toContain('"kind": "selection"');
    expect(firstRequest.messages?.[0]?.content).toContain("Sheet1!A1:B3");

    const entries = await auditStore.listEntries({ session_id: "session_test" });
    expect(entries.length).toBe(1);
    expect(entries[0]?.mode).toBe("chat");
    expect(entries[0]?.tool_calls?.[0]?.name).toBe("write_cell");
    expect(entries[0]?.tool_calls?.[0]?.approved).toBe(true);
  });

  it("can be cancelled while building workbook context (does not call model)", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", 1);

    const auditStore = new MemoryAIAuditStore();
    const chat = vi.fn(async () => {
      throw new Error("should not call model when cancelled before context is ready");
    });

    const ragService = {
      async getContextManager() {
        return new ContextManager({ tokenBudgetTokens: 800 });
      },
      async buildWorkbookContextFromSpreadsheetApi() {
        return new Promise(() => {});
      },
      async dispose() {}
    };

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_abort",
      llmClient: { chat } as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_abort",
      ragService: ragService as any,
    });

    const abortController = new AbortController();
    const promise = orchestrator.sendMessage({
      text: "Hello",
      history: [],
      signal: abortController.signal,
    });

    abortController.abort();

    await expect(promise).rejects.toMatchObject({ name: "AbortError" });
    expect(chat).not.toHaveBeenCalled();
  });
});
