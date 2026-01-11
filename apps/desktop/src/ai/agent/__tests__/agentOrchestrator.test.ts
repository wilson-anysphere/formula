import { describe, expect, it, vi } from "vitest";

import { MemoryAIAuditStore } from "@formula/ai-audit";

import { DocumentController } from "../../../document/documentController.js";
import { runAgentTask } from "../agentOrchestrator.js";

import { DLP_ACTION } from "../../../../../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL } from "../../../../../../packages/security/dlp/src/classification.js";
import { LocalClassificationStore } from "../../../../../../packages/security/dlp/src/classificationStore.js";
import { LocalPolicyStore } from "../../../../../../packages/security/dlp/src/policyStore.js";

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

describe("runAgentTask (agent mode orchestrator)", () => {
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
      const auditStore = new MemoryAIAuditStore();

      const result = await runAgentTask({
        goal: "Read the secret cell",
        workbookId,
        documentController,
        llmClient: llmClient as any,
        auditStore,
        maxIterations: 2,
        maxDurationMs: 10_000,
        model: "unit-test-model"
      });

      expect(result.status).toBe("error");
      expect(result.error).toMatch(/Sending data to cloud AI is restricted/i);
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
    const onApprovalRequired = vi.fn(async () => true);

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

    const onApprovalRequired = vi.fn(async () => false);

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
    const preview = onApprovalRequired.mock.calls[0]![0].preview;
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
    const onApprovalRequired = vi.fn(async () => false);

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
});
