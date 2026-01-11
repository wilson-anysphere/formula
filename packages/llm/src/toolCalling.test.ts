import { describe, expect, it } from "vitest";

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
});
