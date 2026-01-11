import { describe, expect, it } from "vitest";

import { MemoryAIAuditStore } from "@formula/ai-audit";

import { runChatWithToolsAuditedVerified } from "../src/llm/audited-run.js";
import { SpreadsheetLLMToolExecutor } from "../src/llm/integration.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import { parseA1Cell } from "../src/spreadsheet/a1.js";

describe("runChatWithToolsAuditedVerified", () => {
  it("fails verification for data questions when no tools are used", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const toolExecutor = new SpreadsheetLLMToolExecutor(workbook);

    const client = {
      async chat() {
        return {
          message: { role: "assistant", content: "I think the average is 10." },
          usage: { promptTokens: 10, completionTokens: 5 }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();

    const result = await runChatWithToolsAuditedVerified({
      client,
      tool_executor: toolExecutor as any,
      messages: [{ role: "user", content: "What is the average of A1:A3?" }],
      verify_claims: true,
      audit: {
        audit_store: auditStore,
        session_id: "session-verification-1",
        mode: "chat",
        input: { prompt: "What is the average of A1:A3?" },
        model: "unit-test-model"
      }
    });

    expect(result.verification.needs_tools).toBe(true);
    expect(result.verification.used_tools).toBe(false);
    expect(result.verification.verified).toBe(false);
    expect(result.verification.confidence).toBeLessThan(0.5);

    const entries = await auditStore.listEntries({ session_id: "session-verification-1" });
    expect(entries[0]!.verification).toEqual(result.verification);
  });

  it("passes verification when a read-only data tool succeeds", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 3 });

    const toolExecutor = new SpreadsheetLLMToolExecutor(workbook);

    let callCount = 0;
    const client = {
      async chat() {
        callCount++;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A3" } }]
            },
            usage: { promptTokens: 10, completionTokens: 5 }
          };
        }

        return {
          message: { role: "assistant", content: "Average is 2." },
          usage: { promptTokens: 2, completionTokens: 3 }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();

    const result = await runChatWithToolsAuditedVerified({
      client,
      tool_executor: toolExecutor as any,
      messages: [{ role: "user", content: "What is the average of Sheet1!A1:A3?" }],
      verify_claims: true,
      audit: {
        audit_store: auditStore,
        session_id: "session-verification-2",
        mode: "chat",
        input: { prompt: "What is the average of Sheet1!A1:A3?" },
        model: "unit-test-model"
      }
    });

    expect(result.verification.needs_tools).toBe(true);
    expect(result.verification.used_tools).toBe(true);
    expect(result.verification.verified).toBe(true);
    expect(result.verification.confidence).toBeGreaterThanOrEqual(0.9);

    const entries = await auditStore.listEntries({ session_id: "session-verification-2" });
    expect(entries[0]!.tool_calls.some((c) => c.name === "read_range" && c.ok === true)).toBe(true);
    expect(entries[0]!.verification).toEqual(result.verification);
  });

  it("flags incorrect numeric claims with computed actuals + tool evidence", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 3 });

    const toolExecutor = new SpreadsheetLLMToolExecutor(workbook);

    let callCount = 0;
    const client = {
      async chat() {
        callCount++;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A3" } }]
            },
            usage: { promptTokens: 10, completionTokens: 5 }
          };
        }

        return {
          message: { role: "assistant", content: "Average is 10." },
          usage: { promptTokens: 2, completionTokens: 3 }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();

    const result = await runChatWithToolsAuditedVerified({
      client,
      tool_executor: toolExecutor as any,
      messages: [{ role: "user", content: "What is the average of Sheet1!A1:A3?" }],
      verify_claims: true,
      audit: {
        audit_store: auditStore,
        session_id: "session-verification-4",
        mode: "chat",
        input: { prompt: "What is the average of Sheet1!A1:A3?" },
        model: "unit-test-model"
      }
    });

    expect(result.verification.verified).toBe(false);
    expect(result.verification.claims).toHaveLength(1);
    expect(result.verification.claims?.[0]).toMatchObject({
      verified: false,
      expected: 10,
      actual: 2
    });

    const evidence = (result.verification.claims?.[0] as any)?.toolEvidence;
    expect(evidence?.call?.name).toBe("compute_statistics");
    expect(evidence?.result?.data?.statistics?.mean).toBe(2);
  });

  it("retries once with a strict system message when strict_tool_verification is enabled", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 3 });

    const toolExecutor = new SpreadsheetLLMToolExecutor(workbook);

    let callCount = 0;
    const client = {
      async chat(request: any) {
        callCount++;
        if (callCount === 1) {
          return {
            message: { role: "assistant", content: "Probably 2." },
            usage: { promptTokens: 10, completionTokens: 5 }
          };
        }

        if (callCount === 2) {
          const hasStrictSystem = request.messages.some(
            (m: any) => m?.role === "system" && typeof m.content === "string" && m.content.includes("MUST use tools")
          );
          expect(hasStrictSystem).toBe(true);
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A3" } }]
            },
            usage: { promptTokens: 10, completionTokens: 5 }
          };
        }

        return {
          message: { role: "assistant", content: "Average is 2." },
          usage: { promptTokens: 2, completionTokens: 3 }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();

    const result = await runChatWithToolsAuditedVerified({
      client,
      tool_executor: toolExecutor as any,
      messages: [{ role: "user", content: "What is the average of A1:A3?" }],
      strict_tool_verification: true,
      verify_claims: true,
      audit: {
        audit_store: auditStore,
        session_id: "session-verification-3",
        mode: "chat",
        input: { prompt: "What is the average of A1:A3?" },
        model: "unit-test-model"
      }
    });

    expect(callCount).toBe(3);
    expect(result.verification.verified).toBe(true);
  });
});
