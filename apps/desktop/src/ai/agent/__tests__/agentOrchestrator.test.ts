import { describe, expect, it, vi } from "vitest";

import { MemoryAIAuditStore } from "@formula/ai-audit";
import { createHeuristicTokenEstimator, estimateToolDefinitionTokens } from "../../../../../../packages/ai-context/src/tokenBudget.js";

import { DocumentController } from "../../../document/documentController.js";
import { ContextManager } from "../../../../../../packages/ai-context/src/contextManager.js";
import { HashEmbedder } from "../../../../../../packages/ai-rag/src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../../../../../../packages/ai-rag/src/store/inMemoryVectorStore.js";
import { runAgentTask } from "../agentOrchestrator.js";

import { DLP_ACTION } from "../../../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL } from "../../../../../../packages/security/dlp/src/classification.js";
import { LocalClassificationStore } from "../../../../../../packages/security/dlp/src/classificationStore.js";
import { LocalPolicyStore } from "../../../../../../packages/security/dlp/src/policyStore.js";

import { createDesktopRagService } from "../../rag/ragService.js";
import { getAiDlpAuditLogger, resetAiDlpAuditLoggerForTests } from "../../dlp/aiDlp.js";
import { createSheetNameResolverFromIdToNameMap } from "../../../sheet/sheetNameResolver.js";
import { DocumentControllerSpreadsheetApi } from "../../tools/documentControllerSpreadsheetApi.js";

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

describe("runAgentTask (agent mode orchestrator)", () => {
  it("invokes onWorkbookContextBuildStats when provided", async () => {
    const documentController = new DocumentController();
    documentController.setCellValue("Sheet1", { row: 0, col: 0 }, "seed");

    const llmClient = {
      chat: vi.fn(async () => ({ message: { role: "assistant", content: "done" }, usage: { promptTokens: 1, completionTokens: 1 } })),
    };

    const onWorkbookContextBuildStats = vi.fn();
    const auditStore = new MemoryAIAuditStore();
    const result = await runAgentTask({
      goal: "Reply done.",
      workbookId: "wb_agent_stats_hook",
      documentController,
      llmClient: llmClient as any,
      auditStore,
      maxIterations: 2,
      maxDurationMs: 10_000,
      model: "unit-test-model",
      onWorkbookContextBuildStats,
    });

    expect(result.status).toBe("complete");
    expect(result.final).toBe("done");

    expect(onWorkbookContextBuildStats).toHaveBeenCalledTimes(1);
    const stats = onWorkbookContextBuildStats.mock.calls[0]![0];
    expect(stats.mode).toBe("agent");
    expect(stats.model).toBe("unit-test-model");
    expect(stats.durationMs).toBeGreaterThanOrEqual(0);
    expect(stats.blockCountByKind.sheet_sample).toBeGreaterThanOrEqual(1);
    expect(stats.blockCellCountByKind.sheet_sample).toBeGreaterThan(0);
  });

  it("does not re-index workbook RAG when workbook has not changed", async () => {
    const documentController = new DocumentController();
    documentController.setCellValue("Sheet1", { row: 0, col: 0 }, "seed");

    const embedder = new HashEmbedder({ dimension: 32 });
    const vectorStore = new InMemoryVectorStore({ dimension: 32 });
    const contextManager = new ContextManager({
      tokenBudgetTokens: 800,
      workbookRag: { vectorStore, embedder, topK: 3 }
    });

    const indexWorkbookSpy = vi.fn(async () => ({ totalChunks: 0, upserted: 0, skipped: 0, deleted: 0 }));
    const ragService = createDesktopRagService({
      documentController,
      workbookId: "wb_agent_rag_incremental",
      createRag: async () =>
        ({
          vectorStore,
          embedder,
          contextManager,
          indexWorkbook: indexWorkbookSpy
        }) as any
    });

    let callCount = 0;
    const llmClient = {
      chat: vi.fn(async () => {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A1" } }]
            },
            usage: { promptTokens: 1, completionTokens: 1 }
          };
        }
        return {
          message: { role: "assistant", content: "done" },
          usage: { promptTokens: 1, completionTokens: 1 }
        };
      })
    };

    const auditStore = new MemoryAIAuditStore();
    const result = await runAgentTask({
      goal: "Read A1",
      workbookId: "wb_agent_rag_incremental",
      documentController,
      llmClient: llmClient as any,
      auditStore,
      ragService,
      maxIterations: 4,
      maxDurationMs: 10_000,
      model: "unit-test-model"
    });

    expect(result.status).toBe("complete");
    // One index run for the first model call; subsequent iterations should skip re-indexing.
    expect(indexWorkbookSpy).toHaveBeenCalledTimes(1);
    await ragService.dispose();
  });

  it("re-scans workbook RAG when DLP inputs change during an agent run (policy/classifications)", async () => {
    const storage = createInMemoryLocalStorage();
    const original = (globalThis as any).localStorage;
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    const listNonEmptyCellsSpy = vi.spyOn(DocumentControllerSpreadsheetApi.prototype as any, "listNonEmptyCells");
    try {
      storage.clear();

      const workbookId = "wb_agent_dlp_rag_reindex_on_dlp_change";

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

      const documentController = new DocumentController();
      documentController.setCellValue("Sheet1", { row: 0, col: 0 }, "seed");

      const embedder = new HashEmbedder({ dimension: 32 });
      const vectorStore = new InMemoryVectorStore({ dimension: 32 });
      const contextManager = new ContextManager({
        tokenBudgetTokens: 800,
        workbookRag: { vectorStore, embedder, topK: 3 },
      });

      const ragService = createDesktopRagService({
        documentController,
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

      let callCount = 0;
      const llmClient = {
        chat: vi.fn(async () => {
          callCount += 1;
          if (callCount === 1) {
            // Mutate DLP inputs after the first context build; the agent should pick up
            // the updated localStorage state and force a DLP re-index on the next iteration.
            classificationStore.upsert(
              workbookId,
              { scope: "cell", documentId: workbookId, sheetId: "Sheet1", row: 0, col: 0 },
              { level: CLASSIFICATION_LEVEL.CONFIDENTIAL, labels: ["test"] },
            );
            return {
              message: {
                role: "assistant",
                content: "",
                toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A1" } }],
              },
              usage: { promptTokens: 1, completionTokens: 1 },
            };
          }
          return {
            message: { role: "assistant", content: "done" },
            usage: { promptTokens: 1, completionTokens: 1 },
          };
        }),
      };

      const auditStore = new MemoryAIAuditStore();
      const result = await runAgentTask({
        goal: "Read A1",
        workbookId,
        documentController,
        llmClient: llmClient as any,
        auditStore,
        ragService,
        maxIterations: 4,
        maxDurationMs: 10_000,
        model: "unit-test-model",
      });

      expect(result.status).toBe("complete");
      // With DLP enabled, RAG scanning uses workbookFromSpreadsheetApi which calls
      // listNonEmptyCells once per sheet for each indexing run.
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

  it("blocks before calling the LLM when DLP policy forbids cloud AI processing", async () => {
    resetAiDlpAuditLoggerForTests();

    const storage = createInMemoryLocalStorage();
    const original = (globalThis as any).localStorage;
    Object.defineProperty(globalThis, "localStorage", { configurable: true, value: storage });
    try {
      storage.clear();
      const workbookId = "wb_agent_dlp_block";

      const policyStore = new LocalPolicyStore({ storage: storage as any });
      policyStore.setDocumentPolicy(workbookId, {
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
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

      const documentController = new DocumentController();
      documentController.setCellValue("Sheet1", { row: 0, col: 0 }, "TOP SECRET");

      const llmClient = { chat: vi.fn(async () => ({ message: { role: "assistant", content: "should not be called" } })) };
      const onWorkbookContextBuildStats = vi.fn();
      const auditStore = new MemoryAIAuditStore();

      const result = await runAgentTask({
        goal: "Read the secret cell",
        workbookId,
        documentController,
        llmClient: llmClient as any,
        auditStore,
        onWorkbookContextBuildStats,
        maxIterations: 2,
        maxDurationMs: 10_000,
        model: "unit-test-model"
      });

      expect(result.status).toBe("error");
      expect(result.error).toMatch(/Sending data to cloud AI is restricted/i);
      expect(llmClient.chat).not.toHaveBeenCalled();

      expect(onWorkbookContextBuildStats).toHaveBeenCalledTimes(1);
      const stats = onWorkbookContextBuildStats.mock.calls[0]![0];
      expect(stats.ok).toBe(false);
      expect(stats.error?.name).toBe("DlpViolationError");

      const events = getAiDlpAuditLogger().list();
      expect(events.some((e: any) => e.details?.type === "ai.workbook_context")).toBe(true);

      // Even though the run was blocked before the first model call, we should still
      // finalize an agent-mode audit entry for the failed run.
      const auditEntries = await auditStore.listEntries({ session_id: result.session_id });
      expect(auditEntries).toHaveLength(1);
      const entry = auditEntries[0]!;
      expect(entry.mode).toBe("agent");
      expect(entry.tool_calls).toHaveLength(0);
      expect(entry.user_feedback).toBe("rejected");
      expect(entry.input).toMatchObject({ goal: "Read the secret cell" });
      // Audit input should only include metadata about the run, not workbook context payloads.
      expect(entry.input).not.toHaveProperty("workbook_context");
      expect(entry.input).not.toHaveProperty("workbookContext");
      expect(JSON.stringify(entry.input)).not.toContain("TOP SECRET");
    } finally {
      if (original === undefined) {
        // eslint-disable-next-line @typescript-eslint/no-dynamic-delete
        delete (globalThis as any).localStorage;
      } else {
        Object.defineProperty(globalThis, "localStorage", { configurable: true, value: original });
      }
    }
  });

  it("emits progress events in order across multiple tool iterations and records audit", async () => {
    const documentController = new DocumentController();
    documentController.setCellValue("Sheet1", { row: 0, col: 0 }, "seed");

    let callCount = 0;
    const llmClient = {
      async chat() {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "write_cell", arguments: { cell: "Sheet1!A1", value: 5 } }]
            }
          };
        }

        if (callCount === 2) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-2", name: "read_range", arguments: { range: "Sheet1!A1:A1" } }]
            }
          };
        }

        return {
          message: {
            role: "assistant",
            content: "All set."
          }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();
    const events: Array<{ type: string }> = [];
    const onApprovalRequired = vi.fn(async (_prompt: any) => true);

    const result = await runAgentTask({
      goal: "Set A1 to 5 then read it back.",
      workbookId: "wb-1",
      documentController,
      llmClient: llmClient as any,
      auditStore,
      onProgress: (event) => events.push({ type: event.type }),
      onApprovalRequired,
      maxIterations: 8,
      maxDurationMs: 10_000,
      model: "unit-test-model"
    });

    expect(result.status).toBe("complete");
    expect(result.final).toBe("All set.");
    expect(documentController.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(5);

    expect(events.map((e) => e.type)).toEqual([
      "planning",
      "tool_call",
      "tool_result",
      "planning",
      "tool_call",
      "tool_result",
      "planning",
      "assistant_message",
      "complete"
    ]);

    const auditEntries = await auditStore.listEntries({ session_id: result.session_id });
    expect(auditEntries).toHaveLength(1);
    expect(auditEntries[0]!.mode).toBe("agent");
    expect(auditEntries[0]!.input).toMatchObject({ goal: "Set A1 to 5 then read it back." });
    expect(auditEntries[0]!.tool_calls.map((c) => c.name)).toEqual(["write_cell", "read_range"]);
    expect(onApprovalRequired).toHaveBeenCalledTimes(1);
  });

  it("resolves display sheet names in tool calls when sheetNameResolver is provided", async () => {
    const documentController = new DocumentController();
    documentController.setRangeValues("Sheet2", "A1", [
      ["Region", "Revenue"],
      ["North", 1000],
      ["South", 2000],
    ]);

    const sheetNameResolver = createSheetNameResolverFromIdToNameMap(new Map([["Sheet2", "Budget"]]));

    let callCount = 0;
    const llmClient = {
      async chat() {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "write_cell", arguments: { cell: "Budget!C1", value: 99 } }],
            },
            usage: { promptTokens: 1, completionTokens: 1 },
          };
        }
        return {
          message: {
            role: "assistant",
            content: "done",
          },
          usage: { promptTokens: 1, completionTokens: 1 },
        };
      },
    };

    const auditStore = new MemoryAIAuditStore();
    const onApprovalRequired = vi.fn(async () => true);

    const result = await runAgentTask({
      goal: "Set C1 to 99",
      workbookId: "wb_agent_display_names",
      defaultSheetId: "Sheet2",
      documentController,
      sheetNameResolver,
      llmClient: llmClient as any,
      auditStore,
      onApprovalRequired,
      maxIterations: 4,
      maxDurationMs: 10_000,
      model: "unit-test-model",
    });

    expect(result.status).toBe("complete");
    expect(result.final).toBe("done");
    expect(documentController.getCell("Sheet2", "C1").value).toBe(99);
    expect(documentController.getSheetIds()).toContain("Sheet2");
    expect(documentController.getSheetIds()).not.toContain("Budget");
    expect(onApprovalRequired).toHaveBeenCalledTimes(1);
  });

  it("stops safely when approval is required and denied (no mutation applied)", async () => {
    const documentController = new DocumentController();
    documentController.setCellValue("Sheet1", { row: 0, col: 0 }, 1);

    const llmClient = {
      async chat() {
        return {
          message: {
            role: "assistant",
            content: "",
            toolCalls: [{ id: "call-1", name: "write_cell", arguments: { cell: "Sheet1!A1", value: null } }]
          }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();
    const events: string[] = [];

    const onApprovalRequired = vi.fn(async (_prompt: any) => false);

    const result = await runAgentTask({
      goal: "Clear A1",
      workbookId: "wb-2",
      documentController,
      llmClient: llmClient as any,
      auditStore,
      onProgress: (event) => events.push(event.type),
      onApprovalRequired,
      maxIterations: 2,
      maxDurationMs: 10_000,
      model: "unit-test-model"
    });

    expect(result.status).toBe("needs_approval");
    expect(result.denied_call?.name).toBe("write_cell");
    expect(documentController.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(1);
    expect(events).toEqual(["planning", "tool_call", "tool_result", "cancelled"]);

    expect(onApprovalRequired).toHaveBeenCalledTimes(1);
    const preview = onApprovalRequired.mock.calls[0]?.[0]?.preview as any;
    expect(preview.summary.deletes).toBe(1);

    const auditEntries = await auditStore.listEntries({ session_id: result.session_id });
    expect(auditEntries).toHaveLength(1);
    expect(auditEntries[0]!.tool_calls[0]).toMatchObject({
      name: "write_cell",
      requires_approval: true,
      approved: false
    });
  });

  it("honors AbortSignal cancellation", async () => {
    const documentController = new DocumentController();
    documentController.setCellValue("Sheet1", { row: 0, col: 0 }, 42);

    const llmClient = {
      async chat() {
        return {
          message: {
            role: "assistant",
            content: "",
            toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A1" } }]
          }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();
    const controller = new AbortController();
    const events: string[] = [];

    const result = await runAgentTask({
      goal: "Read A1",
      workbookId: "wb-3",
      documentController,
      llmClient: llmClient as any,
      auditStore,
      signal: controller.signal,
      onProgress: (event) => {
        events.push(event.type);
        if (event.type === "tool_call") controller.abort();
      },
      maxIterations: 4,
      maxDurationMs: 10_000,
      model: "unit-test-model"
    });

    expect(result.status).toBe("cancelled");
    expect(events).toEqual(["planning", "tool_call", "cancelled"]);
  });

  it("can continue after approval denial when configured (agent re-plans)", async () => {
    const documentController = new DocumentController();
    documentController.setCellValue("Sheet1", { row: 0, col: 0 }, 1);

    let callCount = 0;
    const llmClient = {
      async chat(request: any) {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "write_cell", arguments: { cell: "Sheet1!A1", value: 99 } }]
            }
          };
        }

        const last = request.messages.at(-1);
        expect(last.role).toBe("tool");
        expect(last.toolCallId).toBe("call-1");
        const payload = JSON.parse(last.content);
        expect(payload.ok).toBe(false);
        expect(payload.error?.code).toBe("approval_denied");

        return {
          message: {
            role: "assistant",
            content: "Okay, I won't make that change. Is there something else you'd like?"
          }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();
    const events: string[] = [];
    const onApprovalRequired = vi.fn(async (_prompt: any) => false);

    const result = await runAgentTask({
      goal: "Set A1 to 99",
      workbookId: "wb-4",
      documentController,
      llmClient: llmClient as any,
      auditStore,
      onProgress: (event) => events.push(event.type),
      onApprovalRequired,
      continueOnApprovalDenied: true,
      maxIterations: 4,
      maxDurationMs: 10_000,
      model: "unit-test-model"
    });

    expect(result.status).toBe("complete");
    expect(result.final).toContain("Okay, I won't make that change");
    expect(documentController.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(1);
    expect(onApprovalRequired).toHaveBeenCalledTimes(1);
    expect(events).toEqual(["planning", "tool_call", "tool_result", "planning", "assistant_message", "complete"]);

    const auditEntries = await auditStore.listEntries({ session_id: result.session_id });
    expect(auditEntries).toHaveLength(1);
    expect(auditEntries[0]!.tool_calls[0]).toMatchObject({
      name: "write_cell",
      requires_approval: true,
      approved: false,
      ok: false
    });
  });

  it("trims large tool results between planning iterations to stay under the context window", async () => {
    const documentController = new DocumentController();
    // Seed a large vertical range so `read_range` returns a big tool payload.
    const values = Array.from({ length: 2000 }, (_, i) => [i]);
    documentController.setRangeValues("Sheet1", "A1", values);

    const estimator = createHeuristicTokenEstimator();
    const contextWindowTokens = 3_000;
    const reserveForOutputTokens = 400;

    let callCount = 0;
    const llmClient = {
      async chat(request: any) {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A2000", include_formulas: false } }]
            }
          };
        }

        const promptTokens =
          estimator.estimateMessagesTokens(request.messages) + estimateToolDefinitionTokens(request.tools, estimator);
        expect(promptTokens).toBeLessThanOrEqual(contextWindowTokens - reserveForOutputTokens);

        const toolMsg = request.messages.find((m: any) => m?.role === "tool" && typeof m.content === "string");
        expect(toolMsg).toBeTruthy();
        // The raw JSON payload for 2000 cells is large; the orchestrator should trim it
        // down rather than passing the full tool output back to the model.
        expect(String(toolMsg.content).length).toBeLessThan(5_000);

        return { message: { role: "assistant", content: "done" } };
      }
    };

    const auditStore = new MemoryAIAuditStore();

    const result = await runAgentTask({
      goal: "Read A1:A2000",
      workbookId: "wb-budget",
      documentController,
      llmClient: llmClient as any,
      auditStore,
      maxIterations: 4,
      maxDurationMs: 10_000,
      model: "unit-test-model",
      contextWindowTokens,
      reserveForOutputTokens,
      tokenEstimator: estimator as any
    });

    expect(result.status).toBe("complete");
    expect(result.final).toBe("done");
  });
});
