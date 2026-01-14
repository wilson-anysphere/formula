import { describe, expect, it, vi } from "vitest";

import { runChatWithTools } from "./toolCalling.js";

import { TOOL_REGISTRY, ToolExecutor as SpreadsheetToolExecutor, InMemoryWorkbook, parseA1Cell } from "../../ai-tools/src/index.js";

describe("runChatWithTools", () => {
  it("executes tool calls and returns the final assistant response", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 42 });
    const executor = new SpreadsheetToolExecutor(workbook);

    const toolExecutor = {
      tools: [
        {
          name: "read_range",
          description: TOOL_REGISTRY.read_range.description,
          parameters: TOOL_REGISTRY.read_range.jsonSchema
        }
      ],
      async execute(call: any) {
        const result = await executor.execute({ name: call.name, parameters: call.arguments });
        if (!result.ok) throw new Error(result.error?.message ?? "tool failed");
        return result;
      }
    };

    let callCount = 0;
    const client = {
      async chat(request: any) {
        callCount++;

        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [
                {
                  id: "call-1",
                  name: "read_range",
                  arguments: { range: "Sheet1!A1:A1" }
                }
              ]
            }
          };
        }

        const lastMessage = request.messages.at(-1);
        expect(lastMessage.role).toBe("tool");
        expect(lastMessage.toolCallId).toBe("call-1");

        return {
          message: {
            role: "assistant",
            content: "A1 is 42."
          }
        };
      }
    };

    const result = await runChatWithTools({
      client: client as any,
      toolExecutor: toolExecutor as any,
      messages: [{ role: "user", content: "What's in A1?" }]
    });

    expect(result.final).toBe("A1 is 42.");
  });

  it("trims tool call names before dispatching to the tool executor", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 42 });
    const executor = new SpreadsheetToolExecutor(workbook);

    const toolExecutor = {
      tools: [
        {
          name: "read_range",
          description: TOOL_REGISTRY.read_range.description,
          parameters: TOOL_REGISTRY.read_range.jsonSchema,
        },
      ],
      async execute(call: any) {
        // The LLM may include leading/trailing whitespace around tool names; the loop
        // should normalize it so tool dispatch still works.
        expect(call.name).toBe("read_range");
        const result = await executor.execute({ name: call.name, parameters: call.arguments });
        if (!result.ok) throw new Error(result.error?.message ?? "tool failed");
        return result;
      },
    };

    let callCount = 0;
    const client = {
      async chat(request: any) {
        callCount++;

        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [
                {
                  id: "call-1",
                  name: "  read_range  ",
                  arguments: { range: "Sheet1!A1:A1" },
                },
              ],
            },
          };
        }

        const lastMessage = request.messages.at(-1);
        expect(lastMessage.role).toBe("tool");
        expect(lastMessage.toolCallId).toBe("call-1");

        return {
          message: {
            role: "assistant",
            content: "A1 is 42.",
          },
        };
      },
    };

    const result = await runChatWithTools({
      client: client as any,
      toolExecutor: toolExecutor as any,
      messages: [{ role: "user", content: "What's in A1?" }],
    });

    expect(result.final).toBe("A1 is 42.");
  });

  it("enforces approval for tools that require it", async () => {
    const client = {
      async chat() {
        return {
          message: {
            role: "assistant",
            content: "",
            toolCalls: [{ id: "call-1", name: "write_cell", arguments: { cell: "A1", value: 1 } }]
          }
        };
      }
    };

    const toolExecutor = {
      tools: [
        {
          name: "write_cell",
          description: TOOL_REGISTRY.write_cell.description,
          parameters: TOOL_REGISTRY.write_cell.jsonSchema,
          requiresApproval: true
        }
      ],
      async execute() {
        throw new Error("should not execute when denied");
      }
    };

    await expect(
      runChatWithTools({
        client: client as any,
        toolExecutor: toolExecutor as any,
        messages: [{ role: "user", content: "Set A1 to 1" }],
        requireApproval: async () => false
      })
    ).rejects.toThrow(/requires approval and was denied/);
  });

  it("can continue when approval is denied (surfaces denial as tool result)", async () => {
    let callCount = 0;
    const client = {
      async chat(request: any) {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "write_cell", arguments: { cell: "A1", value: 1 } }]
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
            content: "Okay, I won't make that change."
          }
        };
      }
    };

    const toolExecutor = {
      tools: [
        {
          name: "write_cell",
          description: TOOL_REGISTRY.write_cell.description,
          parameters: TOOL_REGISTRY.write_cell.jsonSchema,
          requiresApproval: true
        }
      ],
      async execute() {
        throw new Error("should not execute when denied");
      }
    };

    const result = await runChatWithTools({
      client: client as any,
      toolExecutor: toolExecutor as any,
      messages: [{ role: "user", content: "Set A1 to 1" }],
      requireApproval: async () => false,
      continueOnApprovalDenied: true
    });

    expect(result.final).toBe("Okay, I won't make that change.");
    expect(callCount).toBe(2);
  });

  it("aborts while waiting for approval", async () => {
    const abortController = new AbortController();

    const client = {
      async chat() {
        return {
          message: {
            role: "assistant",
            content: "",
            toolCalls: [{ id: "call-1", name: "write_cell", arguments: { cell: "A1", value: 1 } }]
          }
        };
      }
    };

    const toolExecutor = {
      tools: [
        {
          name: "write_cell",
          description: TOOL_REGISTRY.write_cell.description,
          parameters: TOOL_REGISTRY.write_cell.jsonSchema,
          requiresApproval: true
        }
      ],
      execute: vi.fn(async () => ({ ok: true }))
    };

    const requireApproval = vi.fn(async () => {
      queueMicrotask(() => abortController.abort());
      return new Promise<boolean>(() => {});
    });

    await expect(
      runChatWithTools({
        client: client as any,
        toolExecutor: toolExecutor as any,
        messages: [{ role: "user", content: "Set A1 to 1" }],
        requireApproval,
        signal: abortController.signal
      })
    ).rejects.toMatchObject({ name: "AbortError" });

    expect(requireApproval).toHaveBeenCalledTimes(1);
    expect(toolExecutor.execute).not.toHaveBeenCalled();
  });

  it("surfaces unknown tool calls as tool results and continues the loop", async () => {
    const toolExecutor = {
      tools: [
        {
          name: "read_range",
          description: "read",
          parameters: {},
        },
      ],
      execute: vi.fn(async () => {
        throw new Error("should not execute unknown tool");
      }),
    };

    let callCount = 0;
    const client = {
      async chat(request: any) {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "nonexistent_tool", arguments: {} }],
            },
          };
        }

        const last = request.messages.at(-1);
        expect(last.role).toBe("tool");
        expect(last.toolCallId).toBe("call-1");
        const payload = JSON.parse(last.content);
        expect(payload.tool).toBe("nonexistent_tool");
        expect(payload.ok).toBe(false);
        expect(payload.error?.code).toBe("unknown_tool");
        expect(payload.available_tools).toEqual(["read_range"]);

        return { message: { role: "assistant", content: "Recovered." } };
      },
    };

    const onToolResult = vi.fn();
    const result = await runChatWithTools({
      client: client as any,
      toolExecutor: toolExecutor as any,
      messages: [{ role: "user", content: "call an unknown tool" }],
      onToolResult,
    });

    expect(result.final).toBe("Recovered.");
    expect(callCount).toBe(2);
    expect(toolExecutor.execute).not.toHaveBeenCalled();

    expect(onToolResult).toHaveBeenCalledTimes(1);
    expect(onToolResult.mock.calls[0]?.[1]).toMatchObject({ ok: false, error: { code: "unknown_tool" } });
  });

  it("wraps tool execution errors when continueOnToolError=true (continues loop)", async () => {
    const toolExecutor = {
      tools: [
        {
          name: "read_range",
          description: "read",
          parameters: {},
        },
      ],
      execute: vi.fn(async () => {
        throw new Error("boom");
      }),
    };

    let callCount = 0;
    const client = {
      async chat(request: any) {
        callCount += 1;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A1" } }],
            },
          };
        }

        const last = request.messages.at(-1);
        expect(last.role).toBe("tool");
        expect(last.toolCallId).toBe("call-1");
        const payload = JSON.parse(last.content);
        expect(payload.ok).toBe(false);
        expect(payload.error?.code).toBe("tool_execution_error");
        expect(String(payload.error?.message ?? "")).toMatch(/boom/);

        return { message: { role: "assistant", content: "Recovered from tool error." } };
      },
    };

    const result = await runChatWithTools({
      client: client as any,
      toolExecutor: toolExecutor as any,
      messages: [{ role: "user", content: "trigger tool error" }],
      continueOnToolError: true,
    });

    expect(result.final).toBe("Recovered from tool error.");
    expect(callCount).toBe(2);
    expect(toolExecutor.execute).toHaveBeenCalledTimes(1);
  });

  it("rethrows tool execution errors by default (continueOnToolError=false)", async () => {
    const toolExecutor = {
      tools: [
        {
          name: "read_range",
          description: "read",
          parameters: {},
        },
      ],
      execute: vi.fn(async () => {
        throw new Error("boom");
      }),
    };

    let callCount = 0;
    const client = {
      async chat() {
        callCount += 1;
        return {
          message: {
            role: "assistant",
            content: "",
            toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:A1" } }],
          },
        };
      },
    };

    const onToolResult = vi.fn();
    await expect(
      runChatWithTools({
        client: client as any,
        toolExecutor: toolExecutor as any,
        messages: [{ role: "user", content: "trigger tool error" }],
        onToolResult,
      }),
    ).rejects.toThrow(/boom/);

    expect(callCount).toBe(1);
    expect(toolExecutor.execute).toHaveBeenCalledTimes(1);
    expect(onToolResult).toHaveBeenCalledTimes(1);
    expect(onToolResult.mock.calls[0]?.[1]).toMatchObject({ ok: false, error: { code: "tool_execution_error" } });
  });

  it("summarizes large tool results before appending them to the model context", async () => {
    const bigValues = Array.from({ length: 100 }, (_, r) => Array.from({ length: 100 }, (_, c) => r * 100 + c));
    const maxToolResultChars = 1_000;

    const toolExecutor = {
      tools: [
        {
          name: "read_range",
          description: "Read a range",
          parameters: {}
        }
      ],
      async execute(call: any) {
        return {
          tool: "read_range",
          ok: true,
          timing: { started_at_ms: 0, duration_ms: 0 },
          data: { range: call.arguments.range, values: bigValues }
        };
      }
    };

    let callCount = 0;
    const client = {
      async chat(request: any) {
        callCount++;
        if (callCount === 1) {
          return {
            message: {
              role: "assistant",
              content: "",
              toolCalls: [{ id: "call-1", name: "read_range", arguments: { range: "Sheet1!A1:CV100" } }]
            }
          };
        }

        const last = request.messages.at(-1);
        expect(last.role).toBe("tool");
        expect(last.toolCallId).toBe("call-1");
        expect(typeof last.content).toBe("string");
        expect(last.content.length).toBeLessThanOrEqual(maxToolResultChars);

        const payload = JSON.parse(last.content);
        expect(payload.tool).toBe("read_range");
        expect(payload.ok).toBe(true);
        expect(payload.data?.truncated).toBe(true);
        expect(payload.data?.shape).toEqual({ rows: 100, cols: 100 });
        expect(payload.data?.values?.length).toBeLessThanOrEqual(20);
        expect(payload.data?.values?.[0]?.length).toBeLessThanOrEqual(10);

        return {
          message: {
            role: "assistant",
            content: "done"
          }
        };
      }
    };

    const result = await runChatWithTools({
      client: client as any,
      toolExecutor: toolExecutor as any,
      messages: [{ role: "user", content: "Read the data" }],
      maxToolResultChars
    });

    expect(result.final).toBe("done");
  });
});
