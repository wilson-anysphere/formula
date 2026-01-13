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

  it("records token usage from streaming done events", async () => {
    const auditStore = new MemoryAIAuditStore();

    const client = {
      async chat() {
        throw new Error("chat() should not be used when streamChat is available");
      },
      async *streamChat() {
        yield { type: "text", delta: "hello" };
        yield { type: "done", usage: { promptTokens: 7, completionTokens: 4, totalTokens: 11 } };
      }
    };

    const result = await runChatWithToolsAudited({
      client: client as any,
      tool_executor: { tools: [], execute: async () => ({ ok: true }) },
      messages: [{ role: "user", content: "hi" }],
      audit: {
        audit_store: auditStore,
        session_id: "session-stream-usage-1",
        mode: "chat",
        input: { prompt: "hi" },
        model: "unit-test-model"
      }
    });

    expect(result.final).toBe("hello");

    const entries = await auditStore.listEntries({ session_id: "session-stream-usage-1" });
    expect(entries.length).toBe(1);
    expect(entries[0]!.token_usage).toEqual({ prompt_tokens: 7, completion_tokens: 4, total_tokens: 11 });
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
                      url: "https://api.example.com/data?api_key=SECRET&ACCESS_TOKEN=SECRET2&client_secret=SECRET3#frag",
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
    expect(params.url).toContain("ACCESS_TOKEN=REDACTED");
    expect(params.url).toContain("client_secret=REDACTED");
    expect(params.url).not.toContain("SECRET");
    expect(params.url).not.toContain("frag");
    expect(params.headers).toEqual({ Authorization: "REDACTED", Accept: "application/json" });
  });

  it("uses identical URL redaction for tool results and audit parameters", async () => {
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

    const rawUrl =
      "https://api.example.com/data?api_key=SECRET&ACCESS_TOKEN=SECRET2&client_secret=SECRET3&city=berlin#frag";

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
                    url: rawUrl,
                    destination: "Sheet1!A1"
                  }
                }
              ]
            }
          };
        }

        return {
          message: {
            role: "assistant",
            content: "done"
          }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();

    await runChatWithToolsAudited({
      client,
      tool_executor: toolExecutor as any,
      messages: [{ role: "user", content: "Fetch data" }],
      audit: {
        audit_store: auditStore,
        session_id: "session-redaction-match-1",
        mode: "chat",
        input: { prompt: "Fetch data" },
        model: "unit-test-model",
        store_full_tool_results: true
      },
      require_approval: async () => true
    });

    const entries = await auditStore.listEntries({ session_id: "session-redaction-match-1" });
    expect(entries.length).toBe(1);

    const toolCall = entries[0]!.tool_calls[0]!;
    const paramsUrl = (toolCall.parameters as any).url;
    const resultUrl = (toolCall.result as any)?.data?.url;

    expect(paramsUrl).toBe(resultUrl);
    expect(paramsUrl).toContain("api_key=REDACTED");
    expect(paramsUrl).toContain("ACCESS_TOKEN=REDACTED");
    expect(paramsUrl).toContain("client_secret=REDACTED");
    expect(paramsUrl).toContain("city=berlin");
    expect(paramsUrl).not.toContain("SECRET");
    expect(paramsUrl).not.toContain("frag");
  });

  it("stores a bounded tool result summary in audit logs by default", async () => {
    const bigValues = Array.from({ length: 100 }, (_, r) => Array.from({ length: 100 }, (_, c) => r * 100 + c));

    let callCount = 0;
    const client = {
      async chat() {
        callCount++;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:CV100" } }]
            },
            usage: { promptTokens: 10, completionTokens: 5 }
          };
        }
        return {
          message: { role: "assistant", content: "done" },
          usage: { promptTokens: 2, completionTokens: 3 }
        };
      }
    };

    const tool_executor = {
      tools: [{ name: "read_range", description: "read", parameters: {} }],
      async execute(call: any) {
        return {
          tool: "read_range",
          ok: true,
          timing: { started_at_ms: 0, duration_ms: 0 },
          data: { range: call.arguments.range, values: bigValues }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();
    await runChatWithToolsAudited({
      client,
      tool_executor,
      messages: [{ role: "user", content: "Read Sheet1!A1:CV100" }],
      audit: {
        audit_store: auditStore,
        session_id: "session-compaction-1",
        mode: "chat",
        input: { prompt: "Read Sheet1!A1:CV100" },
        model: "unit-test-model"
      }
    });

    const entries = await auditStore.listEntries({ session_id: "session-compaction-1" });
    expect(entries.length).toBe(1);

    const toolCall = entries[0]!.tool_calls[0]!;
    expect(toolCall.result).toBeUndefined();
    expect(toolCall.result_truncated).toBe(true);
    expect(toolCall.audit_result_summary).toBeTypeOf("string");
    expect((toolCall.audit_result_summary as string).length).toBeLessThanOrEqual(20_000);

    const summary = JSON.parse(toolCall.audit_result_summary as string);
    expect(summary.tool).toBe("read_range");
    expect(summary.data?.truncated).toBe(true);
  });

  it("caps oversized tool call parameters in audit logs", async () => {
    const huge = "x".repeat(5000);

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
                  name: "set_range",
                  arguments: { range: "Sheet1!A1", values: [[huge]] }
                }
              ]
            }
          };
        }
        return { message: { role: "assistant", content: "done" } };
      }
    };

    const tool_executor = {
      tools: [{ name: "set_range", description: "set", parameters: {} }],
      async execute(call: any) {
        return {
          tool: "set_range",
          ok: true,
          timing: { started_at_ms: 0, duration_ms: 0 },
          data: { range: call.arguments.range, updated_cells: 1 }
        };
      }
    };

    const auditStore = new MemoryAIAuditStore();
    await runChatWithToolsAudited({
      client,
      tool_executor,
      messages: [{ role: "user", content: "Set A1" }],
      audit: {
        audit_store: auditStore,
        session_id: "session-param-cap-1",
        mode: "chat",
        input: { prompt: "Set A1" },
        model: "unit-test-model",
        max_audit_parameter_chars: 1000
      }
    });

    const entries = await auditStore.listEntries({ session_id: "session-param-cap-1" });
    expect(entries.length).toBe(1);
    const toolCall = entries[0]!.tool_calls[0]!;
    const params = toolCall.parameters as any;
    expect(params?.audit_truncated).toBe(true);
    expect(typeof params?.audit_original_chars).toBe("number");
    expect(params?.range).toBe("Sheet1!A1");
    expect(JSON.stringify(params).length).toBeLessThanOrEqual(1000);
    expect(JSON.stringify(params)).not.toContain(huge);
  });
});
