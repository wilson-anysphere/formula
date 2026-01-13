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

  it("fails verification for data questions when only mutation tools are used", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
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
              toolCalls: [{ id: "call-1", name: "write_cell", arguments: { cell: "Sheet1!A1", value: 1 } }]
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
      // Intentionally disable claim verification so we test tool-usage verification.
      verify_claims: false,
      audit: {
        audit_store: auditStore,
        session_id: "session-verification-mutation-only-1",
        mode: "chat",
        input: { prompt: "What is the average of Sheet1!A1:A3?" },
        model: "unit-test-model"
      }
    });

    expect(result.verification.needs_tools).toBe(true);
    expect(result.verification.used_tools).toBe(true);
    expect(result.verification.verified).toBe(false);

    const entries = await auditStore.listEntries({ session_id: "session-verification-mutation-only-1" });
    expect(entries[0]!.tool_calls.some((c) => c.name === "write_cell" && c.ok === true)).toBe(true);
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

  it("passes verification for mutation prompts when a mutation tool succeeds", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
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
              toolCalls: [{ id: "call-1", name: "write_cell", arguments: { cell: "Sheet1!A1", value: 1 } }]
            },
            usage: { promptTokens: 10, completionTokens: 5 }
          };
        }

        return {
          message: { role: "assistant", content: "Done." },
          usage: { promptTokens: 2, completionTokens: 3 }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();

    const result = await runChatWithToolsAuditedVerified({
      client,
      tool_executor: toolExecutor as any,
      messages: [{ role: "user", content: "Set Sheet1!A1 to 1" }],
      verify_claims: false,
      audit: {
        audit_store: auditStore,
        session_id: "session-verification-mutation-prompt-1",
        mode: "chat",
        input: { prompt: "Set Sheet1!A1 to 1" },
        model: "unit-test-model"
      }
    });

    expect(result.verification.needs_tools).toBe(true);
    expect(result.verification.used_tools).toBe(true);
    expect(result.verification.verified).toBe(true);

    const entries = await auditStore.listEntries({ session_id: "session-verification-mutation-prompt-1" });
    expect(entries[0]!.tool_calls.some((c) => c.name === "write_cell" && c.ok === true)).toBe(true);
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

  it("flags incorrect median claims with computed actuals + tool evidence", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 3 });
    workbook.setCell(parseA1Cell("Sheet1!A4"), { value: 100 });

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
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A4" } }]
            },
            usage: { promptTokens: 10, completionTokens: 5 }
          };
        }

        return {
          message: { role: "assistant", content: "The median of Sheet1!A1:A4 is 10." },
          usage: { promptTokens: 2, completionTokens: 3 }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();

    const result = await runChatWithToolsAuditedVerified({
      client,
      tool_executor: toolExecutor as any,
      messages: [{ role: "user", content: "What is the median of Sheet1!A1:A4?" }],
      verify_claims: true,
      audit: {
        audit_store: auditStore,
        session_id: "session-verification-median-1",
        mode: "chat",
        input: { prompt: "What is the median of Sheet1!A1:A4?" },
        model: "unit-test-model"
      }
    });

    expect(result.verification.verified).toBe(false);
    expect(result.verification.claims).toHaveLength(1);
    expect(result.verification.claims?.[0]).toMatchObject({
      verified: false,
      expected: 10,
      actual: 2.5
    });

    const evidence = (result.verification.claims?.[0] as any)?.toolEvidence;
    expect(evidence?.call?.name).toBe("compute_statistics");
    expect(evidence?.result?.data?.statistics?.median).toBe(2.5);
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

  it("retries once when tools are attempted but none succeed (strict_tool_verification)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 2 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 3 });

    const toolExecutor = new SpreadsheetLLMToolExecutor(workbook);

    let callCount = 0;
    const client = {
      async chat(request: any) {
        callCount++;

        // First run: model tries a tool call, but it's invalid/denied (too large).
        if (callCount === 1) {
          const hasStrictSystem = request.messages.some(
            (m: any) => m?.role === "system" && typeof m.content === "string" && m.content.includes("MUST use tools")
          );
          expect(hasStrictSystem).toBe(false);

          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:Z1000" } }]
            },
            usage: { promptTokens: 10, completionTokens: 5 }
          };
        }

        // Model answers anyway after the failed tool call (no more tools).
        if (callCount === 2) {
          return {
            message: { role: "assistant", content: "Average is 999." },
            usage: { promptTokens: 2, completionTokens: 3 }
          };
        }

        // Second run: strict system message should be present and the model should use tools successfully.
        if (callCount === 3) {
          const hasStrictSystem = request.messages.some(
            (m: any) => m?.role === "system" && typeof m.content === "string" && m.content.includes("MUST use tools")
          );
          expect(hasStrictSystem).toBe(true);

          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-2", name: "read_range", arguments: { range: "Sheet1!A1:A3" } }]
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
      audit: {
        audit_store: auditStore,
        session_id: "session-verification-strict-tool-fail-1",
        mode: "chat",
        input: { prompt: "What is the average of A1:A3?" },
        model: "unit-test-model"
      }
    });

    expect(callCount).toBe(4);
    expect(result.verification.needs_tools).toBe(true);
    expect(result.verification.verified).toBe(true);

    const entries = await auditStore.listEntries({ session_id: "session-verification-strict-tool-fail-1" });
    expect(entries).toHaveLength(1);
    expect(entries[0]!.tool_calls).toHaveLength(2);
    expect(entries[0]!.tool_calls[0]).toMatchObject({ name: "read_range", ok: false });
    expect(entries[0]!.tool_calls[1]).toMatchObject({ name: "read_range", ok: true });
  });
});
