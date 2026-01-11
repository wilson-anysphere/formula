import { afterEach, describe, expect, it, vi } from "vitest";

import { MemoryAIAuditStore } from "@formula/ai-audit";

import { runChatWithToolsAudited } from "../src/llm/audited-run.js";
import { SpreadsheetLLMToolExecutor } from "../src/llm/integration.js";
import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import { parseA1Cell } from "../src/spreadsheet/a1.js";

describe("runChatWithToolsAudited", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("writes an audit entry including approvals + token usage", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const toolExecutor = new SpreadsheetLLMToolExecutor(workbook, { require_approval_for_mutations: true });

    let callCount = 0;
    const requests: any[] = [];
    const client = {
      async chat(request: any) {
        requests.push(request);
        callCount++;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [
                {
                  id: "call-1",
                  name: "write_cell",
                  arguments: { cell: "Sheet1!A1", value: 1 }
                }
              ]
            },
            usage: { promptTokens: 10, completionTokens: 5 }
          };
        }

        return {
          message: {
            role: "assistant",
            content: "done"
          },
          usage: { promptTokens: 2, completionTokens: 3 }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();

    const result = await runChatWithToolsAudited({
      client,
      tool_executor: toolExecutor as any,
      messages: [{ role: "user", content: "Set A1 to 1" }],
      audit: {
        audit_store: auditStore,
        session_id: "session-1",
        mode: "chat",
        input: { prompt: "Set A1 to 1" },
        model: "unit-test-model"
      },
      require_approval: async () => true
    });

    expect(result.final).toBe("done");
    expect(requests[0]?.model).toBe("unit-test-model");
    expect(workbook.getCell(parseA1Cell("Sheet1!A1")).value).toBe(1);

    const entries = await auditStore.listEntries({ session_id: "session-1" });
    expect(entries.length).toBe(1);
    expect(entries[0]!.token_usage).toEqual({ prompt_tokens: 12, completion_tokens: 8, total_tokens: 20 });
    expect(entries[0]!.tool_calls[0]).toMatchObject({
      name: "write_cell",
      requires_approval: true,
      approved: true,
      ok: true
    });
    expect(entries[0]!.user_feedback).toBe("accepted");
  });

  it("redacts fetch_external_data secrets in audit tool call parameters", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    const toolExecutor = new SpreadsheetLLMToolExecutor(workbook, {
      allow_external_data: true,
      allowed_external_hosts: ["api.example.com"]
    });

    vi.stubGlobal(
      "fetch",
      vi.fn(async () => {
        return new Response(JSON.stringify([{ a: 1 }]), {
          status: 200,
          headers: { "content-type": "application/json" }
        });
      }) as any
    );

    let callCount = 0;
    const client = {
      async chat() {
        callCount++;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [
                {
                  id: "call-1",
                  name: "fetch_external_data",
                  arguments: {
                    source_type: "api",
                    url: "https://api.example.com/data?api_key=SECRET#frag",
                    destination: "Sheet1!A1",
                    headers: { Authorization: "Bearer SECRET", Accept: "application/json" }
                  }
                }
              ]
            },
            usage: { promptTokens: 10, completionTokens: 5 }
          };
        }

        return {
          message: {
            role: "assistant",
            content: "done"
          },
          usage: { promptTokens: 2, completionTokens: 3 }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();

    const result = await runChatWithToolsAudited({
      client,
      tool_executor: toolExecutor as any,
      messages: [{ role: "user", content: "Fetch data" }],
      audit: {
        audit_store: auditStore,
        session_id: "session-redaction-1",
        mode: "chat",
        input: { prompt: "Fetch data" },
        model: "unit-test-model"
      },
      require_approval: async () => true
    });

    expect(result.final).toBe("done");

    const entries = await auditStore.listEntries({ session_id: "session-redaction-1" });
    expect(entries.length).toBe(1);

    const toolCall = entries[0]!.tool_calls[0]!;
    expect(toolCall.name).toBe("fetch_external_data");
    const params = toolCall.parameters as any;
    expect(params.url).toContain("api_key=REDACTED");
    expect(params.url).not.toContain("SECRET");
    expect(params.url).not.toContain("frag");
    expect(params.headers).toEqual({ Authorization: "REDACTED", Accept: "application/json" });
  });
});
