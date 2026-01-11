import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../../document/documentController.js";

import { LocalStorageAIAuditStore } from "../../../../../../packages/ai-audit/src/local-storage-store.js";
import { ContextManager } from "../../../../../../packages/ai-context/src/contextManager.js";
import { createHeuristicTokenEstimator, estimateToolDefinitionTokens } from "../../../../../../packages/ai-context/src/tokenBudget.js";
import { CONTEXT_SUMMARY_MARKER } from "../../../../../../packages/ai-context/src/trimMessagesToBudget.js";
import { HashEmbedder } from "../../../../../../packages/ai-rag/src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../../../../../../packages/ai-rag/src/store/inMemoryVectorStore.js";

import { DLP_ACTION } from "../../../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL } from "../../../../../../packages/security/dlp/src/classification.js";
import { LocalClassificationStore } from "../../../../../../packages/security/dlp/src/classificationStore.js";
import { LocalPolicyStore } from "../../../../../../packages/security/dlp/src/policyStore.js";

import { createAiChatOrchestrator } from "../orchestrator.js";
import { ChartStore } from "../../../charts/chartStore";
import { createDesktopRagService } from "../../rag/ragService.js";

import { getAiDlpAuditLogger, resetAiDlpAuditLoggerForTests } from "../../dlp/aiDlp.js";

function createInMemoryLocalStorage(): Storage {
  const store = new Map<string, string>();
  return {
    getItem: (key: string) => (store.has(key) ? store.get(key)! : null),
    setItem: (key: string, value: string) => {
      store.set(String(key), String(value));
    },
    removeItem: (key: string) => {
      store.delete(String(key));
    },
    clear: () => {
      store.clear();
    },
    key: (index: number) => Array.from(store.keys())[index] ?? null,
    get length() {
      return store.size;
    }
  } as Storage;
}

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

function createMockLlmClientWithToolCall(call: { name: string; arguments: unknown }) {
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
              toolCalls: [{ id: "call_1", name: call.name, arguments: call.arguments }],
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
  it("does not re-index workbook RAG when document has not changed", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const embedder = new HashEmbedder({ dimension: 32 });
    const vectorStore = new InMemoryVectorStore({ dimension: 32 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 }
    });

    const indexWorkbookSpy = vi.fn(async () => ({ totalChunks: 0, upserted: 0, skipped: 0, deleted: 0 }));

    const ragService = createDesktopRagService({
      documentController: controller,
      workbookId: "wb_rag_incremental",
      createRag: async () =>
        ({
          vectorStore,
          embedder,
          contextManager,
          indexWorkbook: indexWorkbookSpy
        }) as any
    });

    const llmRequests: any[] = [];
    const llmClient = {
      async chat(request: any) {
        llmRequests.push(request);
        return { message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } };
      }
    };

    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_rag_incremental" });

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_rag_incremental",
      llmClient: llmClient as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_rag_incremental",
      ragService,
      previewOptions: { approval_cell_threshold: 0 }
    });

    await orchestrator.sendMessage({ text: "Hello", history: [] });
    await orchestrator.sendMessage({ text: "Hello again", history: [] });

    expect(indexWorkbookSpy).toHaveBeenCalledTimes(1);

    // Mutate the workbook; the next message should re-index.
    controller.setCellValue("Sheet1", "A1", 123);
    await orchestrator.sendMessage({ text: "After change", history: [] });

    expect(indexWorkbookSpy).toHaveBeenCalledTimes(2);
    await ragService.dispose();
  });

  it("blocks before calling the LLM when DLP policy forbids cloud AI processing", async () => {
    resetAiDlpAuditLoggerForTests();

    const storage = createInMemoryLocalStorage();
    const original = (globalThis as any).localStorage;
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    try {
      storage.clear();

      const workbookId = "wb_dlp_block";

      const policyStore = new LocalPolicyStore({ storage: storage as any });
      policyStore.setDocumentPolicy(workbookId, {
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
            // Only block when Restricted content is in scope.
            maxAllowed: "Confidential",
            allowRestrictedContent: false,
            redactDisallowed: false
          }
        }
      });

      const classificationStore = new LocalClassificationStore({ storage: storage as any });
      classificationStore.upsert(
        workbookId,
        { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
        { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] }
      );

      const controller = new DocumentController();
      controller.setCellValue("Sheet1", "A1", "TOP SECRET");

      const embedder = new HashEmbedder({ dimension: 64 });
      const vectorStore = new InMemoryVectorStore({ dimension: 64 });
      const contextManager = new ContextManager({
        tokenBudgetTokens: 800,
        workbookRag: { vectorStore, embedder, topK: 3 }
      });

      const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_dlp_block" });
      const llmClient = { chat: vi.fn(async () => ({ message: { role: "assistant", content: "should not be called" } })) };

      const orchestrator = createAiChatOrchestrator({
        documentController: controller,
        workbookId,
        llmClient: llmClient as any,
        model: "mock-model",
        getActiveSheetId: () => "Sheet1",
        auditStore,
        sessionId: "session_dlp_block",
        contextManager
      });

      await expect(orchestrator.sendMessage({ text: "What is in A1?", history: [] })).rejects.toThrow(
        /Sending data to cloud AI is restricted/i
      );
      expect(llmClient.chat).not.toHaveBeenCalled();

      const events = getAiDlpAuditLogger().list();
      expect(events.some((e: any) => e.details?.type === "ai.workbook_context")).toBe(true);
    } finally {
      if (original === undefined) {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      } else {
        Object.defineProperty(globalThis, "localStorage", { configurable: true, value: original });
      }
    }
  });

  it("redacts tool results when DLP policy requires redaction (no restricted cells sent to the LLM)", async () => {
    resetAiDlpAuditLoggerForTests();

    const storage = createInMemoryLocalStorage();
    const original = (globalThis as any).localStorage;
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    try {
      storage.clear();

      const workbookId = "wb_dlp_redact";

      const policyStore = new LocalPolicyStore({ storage: storage as any });
      policyStore.setDocumentPolicy(workbookId, {
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
            maxAllowed: "Confidential",
            allowRestrictedContent: false,
            redactDisallowed: true
          }
        }
      });

      const classificationStore = new LocalClassificationStore({ storage: storage as any });
      classificationStore.upsert(
        workbookId,
        { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
        { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] }
      );

      const controller = new DocumentController();
      controller.setCellValue("Sheet1", "A1", "TOP SECRET");

      const embedder = new HashEmbedder({ dimension: 64 });
      const vectorStore = new InMemoryVectorStore({ dimension: 64 });
      const contextManager = new ContextManager({
        tokenBudgetTokens: 800,
        workbookRag: { vectorStore, embedder, topK: 3 }
      });

      const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_dlp_redact" });

      // LLM asks to read A1; tool result should come back redacted and be sent to the model.
      const llmClient = {
        chat: vi.fn(async (request: any) => {
          const messages = Array.isArray(request.messages) ? request.messages : [];
          const toolMessage = messages.slice().reverse().find((m: any) => m && m.role === "tool");
          if (toolMessage) {
            const payload = JSON.parse(toolMessage.content);
            expect(payload.data?.values?.[0]?.[0]).toBe("[REDACTED]");
            return { message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } };
          }
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call_1", name: "read_range", arguments: { range: "Sheet1!A1:A1" } }]
            },
            usage: { promptTokens: 1, completionTokens: 1 }
          };
        })
      };

      const orchestrator = createAiChatOrchestrator({
        documentController: controller,
        workbookId,
        llmClient: llmClient as any,
        model: "mock-model",
        getActiveSheetId: () => "Sheet1",
        auditStore,
        sessionId: "session_dlp_redact",
        contextManager
      });

      const result = await orchestrator.sendMessage({ text: "Read A1", history: [] });
      expect(result.finalText).toBe("ok");

      const events = getAiDlpAuditLogger().list();
      expect(events.some((e: any) => e.details?.type === "ai.workbook_context")).toBe(true);
      expect(events.some((e: any) => e.details?.type === "ai.tool.dlp")).toBe(true);
    } finally {
      if (original === undefined) {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      } else {
        Object.defineProperty(globalThis, "localStorage", { configurable: true, value: original });
      }
    }
  });

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

    const onToolCall = vi.fn();
    const onToolResult = vi.fn();

    await expect(
      orchestrator.sendMessage({
        text: "Set A1 to 99",
        history: [],
        onToolCall,
        onToolResult,
      }),
    ).rejects.toThrow(/denied/i);

    expect(controller.getCell("Sheet1", "A1").value).toBe(1);
    expect(onToolCall).toHaveBeenCalledTimes(1);
    expect(onToolResult).toHaveBeenCalledTimes(1);
    expect(onToolResult.mock.calls[0]?.[1]).toMatchObject({
      ok: false,
      error: expect.objectContaining({ code: "approval_denied" }),
    });

    expect(buildContextSpy).toHaveBeenCalledTimes(1);

    const firstRequest = mock.requests[0];
    expect(firstRequest.messages?.[0]?.role).toBe("system");
    expect(firstRequest.messages?.[0]?.content).toContain("WORKBOOK_CONTEXT");
    expect(firstRequest.messages?.[0]?.content).toContain("Workbook summary");

    const entries = await auditStore.listEntries({ session_id: "session_denied" });
    expect(entries.length).toBe(1);
    expect(entries[0]?.mode).toBe("chat");
    expect(entries[0]?.model).toBe("mock-model");
    expect((entries[0] as any)?.input?.context?.retrieved_ranges).toContain("Sheet1!A1:B2");
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

    const onToolCall = vi.fn();
    const onToolResult = vi.fn();

    const result = await orchestrator.sendMessage({
      text: "Set A1 to 99",
      history: [],
      onToolCall,
      onToolResult,
    });

    expect(result.finalText).toBe("ok");
    expect(result.messages[0]?.role).toBe("user");
    expect(result.toolResults.length).toBe(1);
    expect(result.toolResults[0]?.ok).toBe(true);
    expect(controller.getCell("Sheet1", "A1").value).toBe(99);
    expect(result.context.retrievedRanges).toContain("Sheet1!A1:B2");

    expect(onApprovalRequired).toHaveBeenCalledTimes(1);
    expect(onToolCall).toHaveBeenCalledTimes(1);
    expect(onToolCall.mock.calls[0]?.[0]).toMatchObject({ name: "write_cell" });
    expect(onToolResult).toHaveBeenCalledTimes(1);
    expect(onToolResult.mock.calls[0]?.[0]).toMatchObject({ name: "write_cell" });
    expect(onToolResult.mock.calls[0]?.[1]).toMatchObject({ ok: true });

    const entries = await auditStore.listEntries({ session_id: "session_approved" });
    expect(entries.length).toBe(1);
    expect(entries[0]?.tool_calls?.[0]?.approved).toBe(true);
    expect((entries[0] as any)?.input?.context?.retrieved_ranges).toContain("Sheet1!A1:B2");

    const firstRequest = mock.requests[0];
    expect(firstRequest.messages?.[0]?.content).toContain("Workbook summary");
  });

  it("creates a default ContextManager when none is provided", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_default_context" });
    const mock = createMockLlmClient({ cell: "A1", value: 99 });

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_default_context",
      llmClient: mock.client as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_default_context",
      onApprovalRequired: async () => true,
      previewOptions: { approval_cell_threshold: 0 },
    });

    const result = await orchestrator.sendMessage({
      text: "Set A1 to 99",
      history: [],
    });

    expect(result.finalText).toBe("ok");
    expect(controller.getCell("Sheet1", "A1").value).toBe(99);
    expect(result.context.promptContext).toContain("Workbook summary");
    expect(result.context.retrievedChunkIds.length).toBeGreaterThan(0);
    expect(result.context.retrievedRanges).toContain("Sheet1!A1:B2");

    const firstRequest = mock.requests[0];
    expect(firstRequest.messages?.[0]?.content).toContain("WORKBOOK_CONTEXT");
  });

  it("executes create_chart when chart host support is provided (no approval needed for zero-cell-change preview)", async () => {
    const controller = new DocumentController();
    controller.setCellValue("Sheet1", "A1", "Category");
    controller.setCellValue("Sheet1", "B1", "Value");
    controller.setCellValue("Sheet1", "A2", "A");
    controller.setCellValue("Sheet1", "B2", 10);
    controller.setCellValue("Sheet1", "A3", "B");
    controller.setCellValue("Sheet1", "B3", 20);

    const embedder = new HashEmbedder({ dimension: 64 });
    const vectorStore = new InMemoryVectorStore({ dimension: 64 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 },
    });

    const chartStore = new ChartStore({
      defaultSheet: "Sheet1",
      getCellValue: (sheetId, row, col) => {
        const cell = controller.getCell(sheetId, { row, col }) as { value: unknown } | null;
        return cell?.value ?? null;
      },
    });

    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_chart" });
    const mock = createMockLlmClientWithToolCall({
      name: "create_chart",
      arguments: { chart_type: "bar", data_range: "A1:B3", title: "Sales" },
    });

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_chart",
      llmClient: mock.client as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_chart",
      contextManager,
      previewOptions: { approval_cell_threshold: 0 },
      createChart: chartStore.createChart.bind(chartStore),
    });

    const onToolCall = vi.fn();
    const onToolResult = vi.fn();

    const result = await orchestrator.sendMessage({
      text: "Create a chart",
      history: [],
      onToolCall,
      onToolResult,
    });

    expect(result.finalText).toBe("ok");
    expect(result.toolResults.length).toBe(1);
    expect(result.toolResults[0]?.ok).toBe(true);

    expect(onToolCall).toHaveBeenCalledTimes(1);
    expect(onToolCall.mock.calls[0]?.[0]).toMatchObject({ name: "create_chart" });
    expect(onToolResult).toHaveBeenCalledTimes(1);
    expect(onToolResult.mock.calls[0]?.[0]).toMatchObject({ name: "create_chart" });
    expect(onToolResult.mock.calls[0]?.[1]).toMatchObject({ ok: true });

    const charts = chartStore.listCharts();
    expect(charts).toHaveLength(1);
    expect(charts[0]?.sheetId).toBe("Sheet1");
    expect(charts[0]?.title).toBe("Sales");
    expect(charts[0]?.series[0]).toMatchObject({
      name: "Value",
      categories: "Sheet1!$A$2:$A$3",
      values: "Sheet1!$B$2:$B$3",
    });
  });

  it("trims long history to the configured context window and injects a summary instead of sending unbounded messages", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const embedder = new HashEmbedder({ dimension: 64 });
    const vectorStore = new InMemoryVectorStore({ dimension: 64 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 },
    });

    const mock = createMockLlmClient({ cell: "A1", value: 123 });

    const estimator = createHeuristicTokenEstimator();
    const contextWindowTokens = 6_000;
    const reserveForOutputTokens = 600;

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_budget",
      llmClient: mock.client as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      contextManager,
      onApprovalRequired: async () => true,
      previewOptions: { approval_cell_threshold: 0 },
      contextWindowTokens,
      reserveForOutputTokens,
      keepLastMessages: 20,
      tokenEstimator: estimator as any,
    });

    const longHistory = Array.from({ length: 200 }, (_v, i) => ({
      role: i % 2 === 0 ? ("user" as const) : ("assistant" as const),
      content: `m${i}: ` + "x".repeat(300),
    }));

    await orchestrator.sendMessage({
      text: "Set A1 to 123",
      history: longHistory as any,
    });

    const firstRequest = mock.requests[0];
    expect(firstRequest).toBeTruthy();
    expect(Array.isArray(firstRequest.messages)).toBe(true);

    // History should be trimmed (avoid sending the full unbounded array).
    expect(firstRequest.messages.length).toBeLessThan(longHistory.length);

    const summary = firstRequest.messages.find(
      (m: any) => m?.role === "system" && typeof m.content === "string" && m.content.startsWith(CONTEXT_SUMMARY_MARKER),
    );
    expect(summary).toBeTruthy();

    const promptTokens =
      estimator.estimateMessagesTokens(firstRequest.messages) + estimateToolDefinitionTokens(firstRequest.tools, estimator);
    expect(promptTokens).toBeLessThanOrEqual(contextWindowTokens - reserveForOutputTokens);
  });

  it("exposes only read tools in chat mode by default for non-edit prompts", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const embedder = new HashEmbedder({ dimension: 32 });
    const vectorStore = new InMemoryVectorStore({ dimension: 32 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 }
    });

    const requests: any[] = [];
    const llmClient = {
      async chat(request: any) {
        requests.push(request);
        return { message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } };
      }
    };

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_tool_policy_default",
      llmClient: llmClient as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      contextManager,
      sessionId: "session_tool_policy_default",
      onApprovalRequired: async () => true
    });

    await orchestrator.sendMessage({ text: "What is the average of A1:B2?", history: [] });

    expect(requests).toHaveLength(1);
    const toolNames = (requests[0]?.tools ?? []).map((t: any) => t.name).sort();
    expect(toolNames).toEqual(["compute_statistics", "detect_anomalies", "filter_range", "read_range"]);
  });

  it("upgrades chat tool policy to include mutation tools when prompt implies edits (still approval-gated)", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const embedder = new HashEmbedder({ dimension: 32 });
    const vectorStore = new InMemoryVectorStore({ dimension: 32 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 }
    });

    const requests: any[] = [];
    const llmClient = {
      async chat(request: any) {
        requests.push(request);
        return { message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } };
      }
    };

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_tool_policy_upgrade",
      llmClient: llmClient as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      contextManager,
      sessionId: "session_tool_policy_upgrade",
      onApprovalRequired: async () => true
    });

    await orchestrator.sendMessage({ text: "Update cells A1:A2 to 123", history: [] });

    expect(requests).toHaveLength(1);
    const toolDefs = requests[0]?.tools ?? [];
    const toolNames = toolDefs.map((t: any) => t.name);
    expect(toolNames).toContain("write_cell");
    expect(toolNames).toContain("set_range");
    expect(toolNames).not.toContain("fetch_external_data");
    // Least privilege: don't expose chart/pivot tools for basic range edits.
    expect(toolNames).not.toContain("create_chart");
    expect(toolNames).not.toContain("create_pivot_table");
    // Least privilege: avoid unrelated mutation helpers unless explicitly requested.
    expect(toolNames).not.toContain("apply_formula_column");
    expect(toolNames).not.toContain("sort_range");

    const writeCell = toolDefs.find((t: any) => t.name === "write_cell");
    expect(writeCell?.requiresApproval).toBe(true);
  });
});
