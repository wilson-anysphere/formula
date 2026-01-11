import { describe, expect, it, vi } from "vitest";

import { MemoryAIAuditStore } from "@formula/ai-audit";

import { DocumentController } from "../../../document/documentController.js";
import { runAgentTask } from "../agentOrchestrator.js";

describe("runAgentTask (agent mode orchestrator)", () => {
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

    const result = await runAgentTask({
      goal: "Set A1 to 5 then read it back.",
      workbookId: "wb-1",
      documentController,
      llmClient: llmClient as any,
      auditStore,
      onProgress: (event) => events.push({ type: event.type }),
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
    expect(events).toEqual(["planning", "tool_call", "cancelled"]);

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
});

