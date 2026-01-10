import test from "node:test";
import assert from "node:assert/strict";
import { runChatWithTools } from "./toolCalling.js";

test("runChatWithTools executes tool calls and returns the final assistant response", async () => {
  /** @type {import("./types.js").LLMClient} */
  const client = {
    callCount: 0,
    async chat(request) {
      this.callCount++;

      if (this.callCount === 1) {
        return {
          message: {
            role: "assistant",
            content: "",
            toolCalls: [
              {
                id: "call-1",
                name: "read_range",
                arguments: { range: "Sheet1!A1:A1" },
              },
            ],
          },
        };
      }

      // Ensure the follow-up contains the tool response.
      const lastMessage = request.messages.at(-1);
      assert.equal(lastMessage.role, "tool");
      assert.equal(lastMessage.toolCallId, "call-1");

      return {
        message: {
          role: "assistant",
          content: "A1 is 42.",
        },
      };
    },
  };

  /** @type {import("./types.js").ToolExecutor} */
  const toolExecutor = {
    tools: [
      {
        name: "read_range",
        description: "Read a range of cells",
        parameters: {
          type: "object",
          properties: {
            range: { type: "string" },
          },
          required: ["range"],
        },
      },
    ],
    async execute(call) {
      assert.equal(call.name, "read_range");
      return { range: call.arguments.range, values: [[42]] };
    },
  };

  /** @type {import("./types.js").ToolCall[]} */
  const seenToolCalls = [];

  const result = await runChatWithTools({
    client,
    toolExecutor,
    messages: [{ role: "user", content: "What's in A1?" }],
    onToolCall: (call) => seenToolCalls.push(call),
  });

  assert.equal(result.final, "A1 is 42.");
  assert.equal(seenToolCalls.length, 1);
  assert.equal(seenToolCalls[0].name, "read_range");
});

test("runChatWithTools enforces approval for tools that require it", async () => {
  /** @type {import("./types.js").LLMClient} */
  const client = {
    async chat() {
      return {
        message: {
          role: "assistant",
          content: "",
          toolCalls: [
            { id: "call-1", name: "write_cell", arguments: { cell: "A1", value: 1 } },
          ],
        },
      };
    },
  };

  /** @type {import("./types.js").ToolExecutor} */
  const toolExecutor = {
    tools: [
      {
        name: "write_cell",
        description: "Write to a cell",
        parameters: { type: "object", properties: { cell: { type: "string" }, value: {} }, required: ["cell", "value"] },
        requiresApproval: true,
      },
    ],
    async execute() {
      throw new Error("should not execute when denied");
    },
  };

  await assert.rejects(
    () =>
      runChatWithTools({
        client,
        toolExecutor,
        messages: [{ role: "user", content: "Set A1 to 1" }],
        requireApproval: async () => false,
      }),
    /requires approval and was denied/,
  );
});
