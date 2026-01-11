/**
 * Provider-agnostic tool calling loop (streaming).
 *
 * This mirrors `runChatWithTools` but consumes `client.streamChat()` when available
 * so UIs can render assistant output incrementally while still supporting tool calls.
 */

import { runChatWithTools } from "./toolCalling.js";

/**
 * @param {unknown} value
 */
function safeStringify(value) {
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}

/**
 * @param {unknown} input
 */
function tryParseJson(input) {
  if (typeof input !== "string") return input;
  try {
    return JSON.parse(input);
  } catch {
    return input;
  }
}

/**
 * @param {import("./types.js").ToolDefinition[]} tools
 * @param {string} name
 */
function toolRequiresApproval(tools, name) {
  const tool = tools.find((t) => t.name === name);
  return Boolean(tool?.requiresApproval);
}

/**
 * @typedef {import("./types.js").LLMClient} LLMClient
 * @typedef {import("./types.js").ToolExecutor} ToolExecutor
 * @typedef {import("./types.js").LLMMessage} LLMMessage
 * @typedef {import("./types.js").ToolCall} ToolCall
 * @typedef {import("./types.js").ChatStreamEvent} ChatStreamEvent
 */

/**
 * @param {AsyncIterable<ChatStreamEvent>} stream
 * @param {(event: ChatStreamEvent) => void} [onStreamEvent]
 * @returns {Promise<Extract<LLMMessage, { role: "assistant" }>>}
 */
async function collectAssistantMessageFromStream(stream, onStreamEvent) {
  /** @type {string} */
  let content = "";

  /** @type {Map<string, { id: string, name?: string, args: string }>} */
  const toolCallsById = new Map();
  /** @type {string[]} */
  const toolCallOrder = [];

  /**
   * @param {string} id
   */
  function getOrCreateToolCall(id) {
    const existing = toolCallsById.get(id);
    if (existing) return existing;
    const created = { id, args: "" };
    toolCallsById.set(id, created);
    toolCallOrder.push(id);
    return created;
  }

  for await (const event of stream) {
    onStreamEvent?.(event);
    if (!event || typeof event !== "object") continue;

    if (event.type === "text") {
      content += event.delta ?? "";
      continue;
    }

    if (event.type === "tool_call_start") {
      const call = getOrCreateToolCall(event.id);
      call.name = event.name;
      continue;
    }

    if (event.type === "tool_call_delta") {
      const call = getOrCreateToolCall(event.id);
      call.args += event.delta ?? "";
      continue;
    }

    if (event.type === "tool_call_end") {
      // The loop doesn't require explicit end markers, but providers may emit them.
      // We keep consuming until the stream ends / emits `done`.
      continue;
    }

    if (event.type === "done") {
      break;
    }
  }

  /** @type {ToolCall[]} */
  const toolCalls = [];
  for (const id of toolCallOrder) {
    const call = toolCallsById.get(id);
    if (!call) continue;
    if (!call.name) {
      throw new Error(`Streaming tool call is missing a name (id=${id}).`);
    }
    const args = call.args.trim() ? tryParseJson(call.args) : {};
    toolCalls.push({
      id,
      name: call.name,
      arguments: args,
    });
  }

  return {
    role: "assistant",
    content,
    toolCalls: toolCalls.length ? toolCalls : undefined,
  };
}

/**
 * @param {{
 *   client: LLMClient,
 *   toolExecutor: ToolExecutor,
 *   messages: LLMMessage[],
 *   maxIterations?: number,
 *   onStreamEvent?: (event: ChatStreamEvent) => void,
 *   onToolCall?: (call: ToolCall, meta: { requiresApproval: boolean }) => void,
 *   onToolResult?: (call: ToolCall, result: unknown) => void,
 *   requireApproval?: (call: ToolCall) => Promise<boolean>,
 *   continueOnApprovalDenied?: boolean,
 *   model?: string,
 *   temperature?: number,
 *   maxTokens?: number,
 *   signal?: AbortSignal
 * }} params
 */
export async function runChatWithToolsStreaming(params) {
  if (typeof params.client.streamChat !== "function") {
    return runChatWithTools(params);
  }

  const maxIterations = params.maxIterations ?? 8;
  const requireApproval = params.requireApproval ?? (async () => true);
  const continueOnApprovalDenied = params.continueOnApprovalDenied ?? false;
  const toolDefs = params.toolExecutor.tools ?? [];

  /** @type {LLMMessage[]} */
  const messages = params.messages.slice();

  for (let i = 0; i < maxIterations; i++) {
    const assistant = await collectAssistantMessageFromStream(
      params.client.streamChat({
        messages,
        tools: toolDefs,
        toolChoice: toolDefs.length ? "auto" : "none",
        model: params.model,
        temperature: params.temperature,
        maxTokens: params.maxTokens,
        signal: params.signal,
      }),
      params.onStreamEvent,
    );

    messages.push(assistant);

    const toolCalls = assistant.toolCalls ?? [];
    if (!toolCalls.length) {
      return { messages, final: assistant.content };
    }

    let denied = false;
    let deniedToolName = null;
    for (const call of toolCalls) {
      const requiresApproval = toolRequiresApproval(toolDefs, call.name);
      params.onToolCall?.(call, { requiresApproval });

      if (denied) {
        const skippedResult = {
          tool: call.name,
          ok: false,
          error: {
            code: "skipped_due_to_approval_denied",
            message: `Skipped tool call (${call.name}) because a prior tool call was denied (${deniedToolName ?? "unknown"}).`,
          },
        };
        params.onToolResult?.(call, skippedResult);
        messages.push({
          role: "tool",
          toolCallId: call.id,
          content: safeStringify(skippedResult),
        });
        continue;
      }

      if (requiresApproval) {
        const approved = await requireApproval(call);
        if (!approved) {
          const deniedResult = {
            tool: call.name,
            ok: false,
            error: {
              code: "approval_denied",
              message: `Tool call requires approval and was denied: ${call.name}`,
            },
          };
          params.onToolResult?.(call, deniedResult);
          messages.push({
            role: "tool",
            toolCallId: call.id,
            content: safeStringify(deniedResult),
          });
          if (!continueOnApprovalDenied) {
            throw new Error(deniedResult.error.message);
          }
          denied = true;
          deniedToolName = call.name;
          continue;
        }
      }

      const result = await params.toolExecutor.execute(call);
      params.onToolResult?.(call, result);

      messages.push({
        role: "tool",
        toolCallId: call.id,
        content: safeStringify(result),
      });
    }
  }

  throw new Error(`Exceeded max tool-calling iterations (${maxIterations}).`);
}
