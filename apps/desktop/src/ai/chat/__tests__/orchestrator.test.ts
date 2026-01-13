import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../../document/documentController.js";

import { LocalStorageAIAuditStore } from "../../../../../../packages/ai-audit/src/local-storage-store.js";
import { ContextManager } from "../../../../../../packages/ai-context/src/contextManager.js";
import { createHeuristicTokenEstimator, estimateToolDefinitionTokens, stableJsonStringify } from "../../../../../../packages/ai-context/src/tokenBudget.js";
import { CONTEXT_SUMMARY_MARKER } from "../../../../../../packages/ai-context/src/trimMessagesToBudget.js";
import { HashEmbedder } from "../../../../../../packages/ai-rag/src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../../../../../../packages/ai-rag/src/store/inMemoryVectorStore.js";

import { DLP_ACTION } from "../../../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL } from "../../../../../../packages/security/dlp/src/classification.js";
import { LocalClassificationStore } from "../../../../../../packages/security/dlp/src/classificationStore.js";
import { LocalPolicyStore } from "../../../../../../packages/security/dlp/src/policyStore.js";

import { AiChatOrchestratorError, createAiChatOrchestrator } from "../orchestrator.js";
import { ChartStore } from "../../../charts/chartStore";
import { createDesktopRagService } from "../../rag/ragService.js";
import { DocumentControllerSpreadsheetApi } from "../../tools/documentControllerSpreadsheetApi.js";
import { createSheetNameResolverFromIdToNameMap } from "../../../sheet/sheetNameResolver.js";

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
  it("invokes onWorkbookContextBuildStats when provided", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const llmClient = {
      chat: vi.fn(async () => ({ message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } })),
    };
    const onWorkbookContextBuildStats = vi.fn();
    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_context_stats_hook" });

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_context_stats_hook",
      llmClient: llmClient as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_context_stats_hook",
      previewOptions: { approval_cell_threshold: 0 },
      onWorkbookContextBuildStats,
    });

    await orchestrator.sendMessage({ text: "Hello", history: [] });

    expect(onWorkbookContextBuildStats).toHaveBeenCalledTimes(1);
    const stats = onWorkbookContextBuildStats.mock.calls[0]![0];
    expect(stats.mode).toBe("chat");
    expect(stats.model).toBe("mock-model");
    expect(stats.durationMs).toBeGreaterThanOrEqual(0);
    expect(stats.promptContextChars).toBeGreaterThan(0);
    expect(stats.promptContextTokens).toBeGreaterThan(0);
    expect(stats.blockCountByKind.sheet_sample).toBeGreaterThanOrEqual(1);
    expect(stats.blockCellCountByKind.sheet_sample).toBeGreaterThan(0);

    await orchestrator.dispose();
  });

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

  it("does not rescan workbook RAG when DLP is enabled and inputs have not changed", async () => {
    const storage = createInMemoryLocalStorage();
    const original = (globalThis as any).localStorage;
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    const listNonEmptyCellsSpy = vi.spyOn(DocumentControllerSpreadsheetApi.prototype as any, "listNonEmptyCells");
    try {
      storage.clear();

      const workbookId = "wb_dlp_rag_incremental";

      const policyStore = new LocalPolicyStore({ storage: storage as any });
      policyStore.setDocumentPolicy(workbookId, {
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
            maxAllowed: "Confidential",
            allowRestrictedContent: false,
            redactDisallowed: true,
          },
        },
      });

      // Ensure the classification store is present (even when empty) so
      // `maybeGetAiCloudDlpOptions` returns DLP options.
      new LocalClassificationStore({ storage: storage as any });

      const controller = new DocumentController();
      seed2x2(controller);

      const embedder = new HashEmbedder({ dimension: 32 });
      const vectorStore = new InMemoryVectorStore({ dimension: 32 });
      const contextManager = new ContextManager({
        tokenBudgetTokens: 800,
        workbookRag: { vectorStore, embedder, topK: 3 },
      });

      const ragService = createDesktopRagService({
        documentController: controller,
        workbookId,
        createRag: async () =>
          ({
            vectorStore,
            embedder,
            contextManager,
            // Not used when DLP is enabled, but required by DesktopRagService's contract.
            indexWorkbook: async () => ({ totalChunks: 0, upserted: 0, skipped: 0, deleted: 0 }),
          }) as any,
      });

      const llmClient = {
        chat: vi.fn(async () => ({ message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } })),
      };

      const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_dlp_rag_incremental" });

      const orchestrator = createAiChatOrchestrator({
        documentController: controller,
        workbookId,
        llmClient: llmClient as any,
        model: "mock-model",
        getActiveSheetId: () => "Sheet1",
        auditStore,
        sessionId: "session_dlp_rag_incremental",
        ragService,
        previewOptions: { approval_cell_threshold: 0 },
      });

      await orchestrator.sendMessage({ text: "Hello", history: [] });
      await orchestrator.sendMessage({ text: "Hello again", history: [] });

      // RAG scanning is driven by `workbookFromSpreadsheetApi` which calls
      // SpreadsheetApi.listNonEmptyCells once per sheet. With DLP enabled we should
      // still hit it only once when nothing changes.
      expect(listNonEmptyCellsSpy).toHaveBeenCalledTimes(1);

      await ragService.dispose();
    } finally {
      listNonEmptyCellsSpy.mockRestore();
      if (original === undefined) {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      } else {
        Object.defineProperty(globalThis, "localStorage", { configurable: true, value: original });
      }
    }
  });

  it("re-scans workbook RAG when DLP inputs change (policy/classifications) even if the workbook is unchanged", async () => {
    const storage = createInMemoryLocalStorage();
    const original = (globalThis as any).localStorage;
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    const listNonEmptyCellsSpy = vi.spyOn(DocumentControllerSpreadsheetApi.prototype as any, "listNonEmptyCells");
    try {
      storage.clear();

      const workbookId = "wb_dlp_rag_reindex_on_dlp_change";

      const policyStore = new LocalPolicyStore({ storage: storage as any });
      policyStore.setDocumentPolicy(workbookId, {
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
            maxAllowed: "Confidential",
            allowRestrictedContent: false,
            redactDisallowed: true,
          },
        },
      });

      const classificationStore = new LocalClassificationStore({ storage: storage as any });

      const controller = new DocumentController();
      seed2x2(controller);

      const embedder = new HashEmbedder({ dimension: 32 });
      const vectorStore = new InMemoryVectorStore({ dimension: 32 });
      const contextManager = new ContextManager({
        tokenBudgetTokens: 800,
        workbookRag: { vectorStore, embedder, topK: 3 },
      });

      const ragService = createDesktopRagService({
        documentController: controller,
        workbookId,
        createRag: async () =>
          ({
            vectorStore,
            embedder,
            contextManager,
            indexWorkbook: async () => ({ totalChunks: 0, upserted: 0, skipped: 0, deleted: 0 }),
          }) as any,
      });

      const llmClient = {
        chat: vi.fn(async () => ({ message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } })),
      };

      const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_dlp_rag_reindex_on_dlp_change" });

      const orchestrator = createAiChatOrchestrator({
        documentController: controller,
        workbookId,
        llmClient: llmClient as any,
        model: "mock-model",
        getActiveSheetId: () => "Sheet1",
        auditStore,
        sessionId: "session_dlp_rag_reindex_on_dlp_change",
        ragService,
        previewOptions: { approval_cell_threshold: 0 },
      });

      await orchestrator.sendMessage({ text: "Hello", history: [] });

      // Mutate DLP inputs: add a (still-allowed) classification record.
      classificationStore.upsert(
        workbookId,
        { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
        { level: CLASSIFICATION_LEVEL.CONFIDENTIAL, labels: ["test"] },
      );

      await orchestrator.sendMessage({ text: "Hello again", history: [] });

      expect(listNonEmptyCellsSpy).toHaveBeenCalledTimes(2);

      await ragService.dispose();
    } finally {
      listNonEmptyCellsSpy.mockRestore();
      if (original === undefined) {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      } else {
        Object.defineProperty(globalThis, "localStorage", { configurable: true, value: original });
      }
    }
  });

  it("reuses WorkbookContextBuilder across chat messages to reduce SpreadsheetApi.readRange calls", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const readRangeSpy = vi.spyOn(DocumentControllerSpreadsheetApi.prototype as any, "readRange");
    try {
      const ragService = {
        async getContextManager() {
          return new ContextManager({ tokenBudgetTokens: 800 });
        },
        async buildWorkbookContextFromSpreadsheetApi() {
          // Keep the test focused on WorkbookContextBuilder's own sheet/block reads.
          return { retrieved: [] };
        },
        async dispose() {}
      };

      const llmClient = {
        chat: vi.fn(async () => ({ message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } })),
      };

      const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_workbook_context_builder_cache" });

      const orchestrator = createAiChatOrchestrator({
        documentController: controller,
        workbookId: "wb_workbook_context_builder_cache",
        llmClient: llmClient as any,
        model: "mock-model",
        getActiveSheetId: () => "Sheet1",
        auditStore,
        sessionId: "session_workbook_context_builder_cache",
        ragService: ragService as any,
        previewOptions: { approval_cell_threshold: 0 }
      });

      const before1 = readRangeSpy.mock.calls.length;
      await orchestrator.sendMessage({ text: "Hello", history: [] });
      const after1 = readRangeSpy.mock.calls.length;
      const calls1 = after1 - before1;

      const before2 = readRangeSpy.mock.calls.length;
      await orchestrator.sendMessage({ text: "Hello again", history: [] });
      const after2 = readRangeSpy.mock.calls.length;
      const calls2 = after2 - before2;

      expect(calls1).toBeGreaterThan(0);
      expect(calls2).toBeLessThan(calls1);
      // Second message should hit WorkbookContextBuilder's sheet/block caches.
      expect(calls2).toBe(0);
    } finally {
      readRangeSpy.mockRestore();
    }
  });

  it("aborts during context building without continuing read_range work in the background (and does not call the LLM)", async () => {
    const controller = new DocumentController();
    controller.setRangeValues("Sheet1", "A1", [
      ["A"],
      ["B"],
    ]);
    controller.setRangeValues("Sheet2", "A1", [
      ["C"],
      ["D"],
    ]);

    const abortController = new AbortController();
    const signal = abortController.signal;

    const originalReadRange = (DocumentControllerSpreadsheetApi.prototype as any).readRange;
    let readRangeCalls = 0;
    const readRangeSpy = vi.spyOn(DocumentControllerSpreadsheetApi.prototype as any, "readRange").mockImplementation(function (
      this: any,
      range: any,
    ) {
      readRangeCalls += 1;
      if (readRangeCalls === 1) abortController.abort();
      return originalReadRange.call(this, range);
    });

    try {
      const ragService = {
        async getContextManager() {
          return new ContextManager({ tokenBudgetTokens: 800 });
        },
        async buildWorkbookContextFromSpreadsheetApi() {
          // Keep the test focused on WorkbookContextBuilder's own sheet/block reads.
          return { retrieved: [] };
        },
        async dispose() {}
      };

      const llmClient = {
        chat: vi.fn(async () => ({ message: { role: "assistant", content: "should not be called" } })),
      };

      const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_abort_during_context_build" });

      const orchestrator = createAiChatOrchestrator({
        documentController: controller,
        workbookId: "wb_abort_during_context_build",
        llmClient: llmClient as any,
        model: "mock-model",
        getActiveSheetId: () => "Sheet1",
        auditStore,
        sessionId: "session_abort_during_context_build",
        ragService: ragService as any,
        previewOptions: { approval_cell_threshold: 0 }
      });

      await expect(orchestrator.sendMessage({ text: "Hello", history: [], signal })).rejects.toMatchObject({ name: "AbortError" });
      expect(llmClient.chat).not.toHaveBeenCalled();

      const callsAfterAbort = readRangeCalls;
      await new Promise((resolve) => setTimeout(resolve, 25));
      expect(readRangeCalls).toBe(callsAfterAbort);
      expect(readRangeCalls).toBe(1);
    } finally {
      readRangeSpy.mockRestore();
    }
  });

  it("surfaces AbortError when aborted during the LLM/tool loop (does not wrap in AiChatOrchestratorError)", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const abortController = new AbortController();
    const signal = abortController.signal;

    const ragService = {
      async getContextManager() {
        return new ContextManager({ tokenBudgetTokens: 800 });
      },
      async buildWorkbookContextFromSpreadsheetApi() {
        // Keep the test focused on abort propagation, not RAG retrieval.
        return { retrieved: [] };
      },
      async dispose() {}
    };

    let resolveRequestSignal: ((signal: AbortSignal | undefined) => void) | null = null;
    const requestSignalPromise = new Promise<AbortSignal | undefined>((resolve) => {
      resolveRequestSignal = resolve;
    });

    const llmClient = {
      chat: vi.fn(async (request: any) => {
        const requestSignal: AbortSignal | undefined = request?.signal;
        resolveRequestSignal?.(requestSignal);

        return await new Promise((_resolve, reject) => {
          if (!requestSignal) {
            reject(new Error("Expected llmClient.chat to receive request.signal"));
            return;
          }

          const onAbort = () => {
            const err = new Error("Aborted");
            err.name = "AbortError";
            reject(err);
          };

          if (requestSignal.aborted) {
            onAbort();
            return;
          }

          requestSignal.addEventListener("abort", onAbort, { once: true });
        });
      })
    };

    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_abort_during_tool_loop" });

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_abort_during_tool_loop",
      llmClient: llmClient as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_abort_during_tool_loop",
      ragService: ragService as any,
      previewOptions: { approval_cell_threshold: 0 }
    });

    const promise = orchestrator.sendMessage({ text: "hi", history: [], signal });

    const requestSignal = await requestSignalPromise;
    expect(requestSignal).toBe(signal);

    abortController.abort();

    let thrown: unknown;
    try {
      await promise;
    } catch (err) {
      thrown = err;
    }

    expect(thrown).toBeTruthy();
    expect(thrown).toMatchObject({ name: "AbortError" });
    expect(thrown).not.toBeInstanceOf(AiChatOrchestratorError);
  });

  it("recreates WorkbookContextBuilder when DLP inputs change (prevents cache reuse across policy/classification changes)", async () => {
    const storage = createInMemoryLocalStorage();
    const original = (globalThis as any).localStorage;
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    const readRangeSpy = vi.spyOn(DocumentControllerSpreadsheetApi.prototype as any, "readRange");
    try {
      storage.clear();

      const workbookId = "wb_dlp_workbook_context_builder_cache";

      const policyStore = new LocalPolicyStore({ storage: storage as any });
      policyStore.setDocumentPolicy(workbookId, {
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
            maxAllowed: "Confidential",
            allowRestrictedContent: false,
            redactDisallowed: true,
          },
        },
      });

      const classificationStore = new LocalClassificationStore({ storage: storage as any });

      const controller = new DocumentController();
      seed2x2(controller);

      const ragService = {
        async getContextManager() {
          return new ContextManager({ tokenBudgetTokens: 800 });
        },
        async buildWorkbookContextFromSpreadsheetApi() {
          return { retrieved: [] };
        },
        async dispose() {}
      };

      const llmClient = {
        chat: vi.fn(async () => ({ message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } })),
      };

      const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_dlp_workbook_context_builder_cache" });

      const orchestrator = createAiChatOrchestrator({
        documentController: controller,
        workbookId,
        llmClient: llmClient as any,
        model: "mock-model",
        getActiveSheetId: () => "Sheet1",
        auditStore,
        sessionId: "session_dlp_workbook_context_builder_cache",
        ragService: ragService as any,
        previewOptions: { approval_cell_threshold: 0 }
      });

      const before1 = readRangeSpy.mock.calls.length;
      await orchestrator.sendMessage({ text: "Hello", history: [] });
      const after1 = readRangeSpy.mock.calls.length;
      const calls1 = after1 - before1;

      const before2 = readRangeSpy.mock.calls.length;
      await orchestrator.sendMessage({ text: "Hello again", history: [] });
      const after2 = readRangeSpy.mock.calls.length;
      const calls2 = after2 - before2;

      expect(calls1).toBeGreaterThan(0);
      expect(calls2).toBe(0);

      // Mutate DLP inputs: add a (still-allowed) classification record.
      classificationStore.upsert(
        workbookId,
        { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
        { level: CLASSIFICATION_LEVEL.CONFIDENTIAL, labels: ["test"] },
      );

      const before3 = readRangeSpy.mock.calls.length;
      await orchestrator.sendMessage({ text: "After DLP change", history: [] });
      const after3 = readRangeSpy.mock.calls.length;
      const calls3 = after3 - before3;

      // DLP inputs changed => cached builder must be dropped, so we re-read sheet context.
      expect(calls3).toBeGreaterThan(0);
      expect(calls3).toBeGreaterThan(calls2);
    } finally {
      readRangeSpy.mockRestore();
      if (original === undefined) {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      } else {
        Object.defineProperty(globalThis, "localStorage", { configurable: true, value: original });
      }
    }
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
      const onWorkbookContextBuildStats = vi.fn();

      const orchestrator = createAiChatOrchestrator({
        documentController: controller,
        workbookId,
        llmClient: llmClient as any,
        model: "mock-model",
        getActiveSheetId: () => "Sheet1",
        auditStore,
        sessionId: "session_dlp_block",
        contextManager,
        onWorkbookContextBuildStats,
      });

      await expect(orchestrator.sendMessage({ text: "What is in A1?", history: [] })).rejects.toThrow(
        /Sending data to cloud AI is restricted/i
      );
      expect(llmClient.chat).not.toHaveBeenCalled();

      expect(onWorkbookContextBuildStats).toHaveBeenCalledTimes(1);
      const stats = onWorkbookContextBuildStats.mock.calls[0]![0];
      expect(stats.ok).toBe(false);
      expect(stats.error?.name).toBe("DlpViolationError");
      expect(stats.error?.message).toMatch(/Sending data to cloud AI is restricted/i);

      const entries = await auditStore.listEntries({ session_id: "session_dlp_block" });
      expect(entries.length).toBe(1);
      expect(entries[0]?.mode).toBe("chat");
      expect((entries[0]?.input as any)?.blocked).toBe(true);
      expect(JSON.stringify(entries[0]?.input)).not.toContain("TOP SECRET");

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

  it("blocks before calling the LLM when DLP policy forbids cloud AI processing on renamed sheets (uses sheetNameResolver)", async () => {
    resetAiDlpAuditLoggerForTests();

    const storage = createInMemoryLocalStorage();
    const original = (globalThis as any).localStorage;
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    try {
      storage.clear();

      const workbookId = "wb_dlp_block_renamed_sheet";

      const policyStore = new LocalPolicyStore({ storage: storage as any });
      policyStore.setDocumentPolicy(workbookId, {
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
            // Only block when Restricted content is in scope.
            maxAllowed: "Confidential",
            allowRestrictedContent: false,
            redactDisallowed: false,
          },
        },
      });

      const controller = new DocumentController();
      const internalSheetId = controller.addSheet({ name: "Budget" });
      controller.setCellValue(internalSheetId, "A1", "TOP SECRET");

      const classificationStore = new LocalClassificationStore({ storage: storage as any });
      classificationStore.upsert(
        workbookId,
        { scope: "cell", documentId: workbookId, sheetId: internalSheetId, row: 0, col: 0 },
        { level: CLASSIFICATION_LEVEL.RESTRICTED, labels: ["test"] },
      );

      const sheetNameResolver = createSheetNameResolverFromIdToNameMap(new Map([[internalSheetId, "Budget"]]));

      const embedder = new HashEmbedder({ dimension: 64 });
      const vectorStore = new InMemoryVectorStore({ dimension: 64 });
      const contextManager = new ContextManager({
        tokenBudgetTokens: 800,
        workbookRag: { vectorStore, embedder, topK: 3 },
      });

      const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_dlp_block_renamed_sheet" });
      const llmClient = { chat: vi.fn(async () => ({ message: { role: "assistant", content: "should not be called" } })) };

      const orchestrator = createAiChatOrchestrator({
        documentController: controller,
        workbookId,
        llmClient: llmClient as any,
        model: "mock-model",
        getActiveSheetId: () => internalSheetId,
        sheetNameResolver,
        auditStore,
        sessionId: "session_dlp_block_renamed_sheet",
        contextManager,
      });

      await expect(orchestrator.sendMessage({ text: "What is in Budget!A1?", history: [] })).rejects.toThrow(
        /Sending data to cloud AI is restricted/i,
      );
      expect(llmClient.chat).not.toHaveBeenCalled();
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

  it("resolves display sheet names in tool calls when sheetNameResolver is provided (prevents phantom sheets)", async () => {
    const controller = new DocumentController();
    controller.setRangeValues("Sheet2", "A1", [
      ["Region", "Revenue"],
      ["North", 1000],
      ["South", 2000],
    ]);

    const sheetNameResolver = createSheetNameResolverFromIdToNameMap(new Map([["Sheet2", "Budget"]]));

    const embedder = new HashEmbedder({ dimension: 64 });
    const vectorStore = new InMemoryVectorStore({ dimension: 64 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 },
    });

    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_sheet_name_resolver_chat" });
    const mock = createMockLlmClient({ cell: "Budget!C1", value: 99 });
    const onApprovalRequired = vi.fn(async () => true);

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_chat_display_names",
      llmClient: mock.client as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet2",
      sheetNameResolver,
      auditStore,
      sessionId: "session_chat_display_names",
      contextManager,
      onApprovalRequired,
      previewOptions: { approval_cell_threshold: 0 },
    });

    const result = await orchestrator.sendMessage({ text: "Set Budget!C1 to 99", history: [] });

    expect(result.finalText).toBe("ok");
    expect(controller.getCell("Sheet2", "C1").value).toBe(99);
    expect(controller.getSheetIds()).toContain("Sheet2");
    expect(controller.getSheetIds()).not.toContain("Budget");
    expect(onApprovalRequired).toHaveBeenCalledTimes(1);
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

    await orchestrator.dispose();
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
      strictToolVerification: false,
      onApprovalRequired: async () => true
    });

    await orchestrator.sendMessage({ text: "What is the average of A1:B2?", history: [] });

    expect(requests).toHaveLength(1);
    const toolNames = (requests[0]?.tools ?? []).map((t: any) => t.name).sort();
    expect(toolNames).toEqual(["compute_statistics", "detect_anomalies", "filter_range", "read_range"]);
  });

  it("retries once with strict tool verification when the model answers a data question without tools", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    let callCount = 0;
    const llmClient = {
      chat: vi.fn(async (_request: any) => {
        callCount += 1;

        // First run: model answers directly (no toolCalls) even though the prompt
        // clearly references spreadsheet data.
        if (callCount === 1) {
          return {
            message: { role: "assistant", content: "The sum is 4." },
            usage: { promptTokens: 1, completionTokens: 1 }
          };
        }

        // Strict retry: model issues a tool call.
        if (callCount === 2) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call_1", name: "read_range", arguments: { range: "Sheet1!A1:A2" } }]
            },
            usage: { promptTokens: 1, completionTokens: 1 }
          };
        }

        // After tool execution: model provides final answer.
        return {
          message: { role: "assistant", content: "The sum is 4." },
          usage: { promptTokens: 1, completionTokens: 1 }
        };
      }),
    };

    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_strict_tool_verification" });
    const embedder = new HashEmbedder({ dimension: 32 });
    const vectorStore = new InMemoryVectorStore({ dimension: 32 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 }
    });

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_strict_tool_verification",
      llmClient: llmClient as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_strict_tool_verification",
      contextManager,
      strictToolVerification: true,
    });

    const onToolCall = vi.fn();
    const onToolResult = vi.fn();

    const result = await orchestrator.sendMessage({
      text: "What is the sum of A1:A2?",
      history: [],
      onToolCall,
      onToolResult,
    });

    // Should have retried (1st call answered without tools; retry performs tool loop).
    expect(llmClient.chat).toHaveBeenCalledTimes(3);
    expect(result.toolResults.length).toBe(1);
    expect(result.toolResults[0]).toMatchObject({ tool: "read_range", ok: true });

    expect(onToolCall).toHaveBeenCalledTimes(1);
    expect(onToolCall.mock.calls[0]?.[0]).toMatchObject({ name: "read_range" });
    expect(onToolResult).toHaveBeenCalledTimes(1);
    expect(onToolResult.mock.calls[0]?.[0]).toMatchObject({ name: "read_range" });
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
      strictToolVerification: false,
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

  it("never forwards small table/range attachment data to the model prompt (prompt + audit)", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const requests: any[] = [];
    const llmClient = {
      async chat(request: any) {
        requests.push(request);
        return { message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } };
      },
    };

    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_no_raw_table_prompt" });

    // Keep the test focused on prompt/audit behavior; no need to exercise workbook RAG.
    // Echo attachments back so they show up in promptContext (system prompt) too.
    const ragService = {
      async getContextManager() {
        return new ContextManager({ tokenBudgetTokens: 800 });
      },
      async buildWorkbookContextFromSpreadsheetApi(params: any) {
        return { retrieved: [], attachments: params.attachments };
      },
      async dispose() {},
    };

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_no_raw_table_prompt",
      llmClient: llmClient as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_no_raw_table_prompt",
      ragService: ragService as any,
      strictToolVerification: false,
    });

    const secret = "TOP SECRET";
    const attachments = [{ type: "table" as const, reference: "Sheet1!A1:B2", data: { snapshot: secret } }];

    await orchestrator.sendMessage({ text: "Hello", history: [], attachments });

    expect(requests).toHaveLength(1);
    const systemMessage = (requests[0]?.messages ?? []).find((m: any) => m?.role === "system");
    const userMessage = (requests[0]?.messages ?? []).find((m: any) => m?.role === "user");
    expect(systemMessage).toBeTruthy();
    expect(userMessage).toBeTruthy();
    expect(systemMessage.content).not.toContain(secret);
    expect(userMessage.content).not.toContain(secret);

    const entries = await auditStore.listEntries({ session_id: "session_no_raw_table_prompt" });
    expect(entries).toHaveLength(1);
    expect(JSON.stringify((entries[0] as any)?.input)).not.toContain(secret);
  });

  it("compacts large attachment data in prompts and audit logs", async () => {
    const controller = new DocumentController();
    seed2x2(controller);

    const requests: any[] = [];
    const llmClient = {
      async chat(request: any) {
        requests.push(request);
        return { message: { role: "assistant", content: "ok" }, usage: { promptTokens: 1, completionTokens: 1 } };
      },
    };

    const auditStore = new LocalStorageAIAuditStore({ key: "test_audit_attachment_compaction" });

    // Keep the test focused on prompt/audit behavior; no need to exercise workbook RAG.
    const ragService = {
      async getContextManager() {
        return new ContextManager({ tokenBudgetTokens: 800 });
      },
      async buildWorkbookContextFromSpreadsheetApi() {
        return { retrieved: [] };
      },
      async dispose() {},
    };

    const orchestrator = createAiChatOrchestrator({
      documentController: controller,
      workbookId: "wb_attachment_compaction",
      llmClient: llmClient as any,
      model: "mock-model",
      getActiveSheetId: () => "Sheet1",
      auditStore,
      sessionId: "session_attachment_compaction",
      ragService: ragService as any,
      strictToolVerification: false,
    });

    const giant = "x".repeat(50_000);
    const attachments = [
      { type: "table" as const, reference: "Sheet1!A1:B2", data: { snapshot: giant } },
    ];

    await orchestrator.sendMessage({ text: "Hello", history: [], attachments });

    expect(requests).toHaveLength(1);
    const userMessage = (requests[0]?.messages ?? []).find((m: any) => m?.role === "user");
    expect(userMessage).toBeTruthy();
    expect(userMessage.content).not.toContain(giant);

    const entries = await auditStore.listEntries({ session_id: "session_attachment_compaction" });
    expect(entries).toHaveLength(1);
    const auditAttachments = (entries[0] as any)?.input?.attachments;
    expect(auditAttachments).toEqual([
      {
        type: "table",
        reference: "Sheet1!A1:B2",
        data: expect.objectContaining({
          truncated: true,
          hash: expect.any(String),
          original_chars: expect.any(Number),
        }),
      },
    ]);
    expect(JSON.stringify(auditAttachments)).not.toContain(giant);

    const expectedJson = stableJsonStringify({ snapshot: giant });
    const expectedHash = fnv1a32(expectedJson).toString(16);
    expect(auditAttachments[0].data.hash).toBe(expectedHash);
    expect(auditAttachments[0].data.original_chars).toBe(expectedJson.length);
  });
});

function fnv1a32(value: string): number {
  // Keep in sync with the chat orchestrator's attachment hashing.
  let hash = 0x811c9dc5;
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}
