/**
 * Provider-agnostic tool calling loop (streaming).
 *
 * This mirrors `runChatWithTools` but consumes `client.streamChat()` when available
 * so UIs can render assistant output incrementally while still supporting tool calls.
 */

import { runChatWithTools } from "./toolCalling.js";
import { serializeToolResultForModel } from "./toolResultSerialization.js";

/**
 * Tool names are identifiers and should not include leading/trailing whitespace.
 *
 * @param {unknown} value
 * @returns {string}
 */
function normalizeToolCallName(value) {
  if (typeof value !== "string") return "";
  const trimmed = value.trim();
  return trimmed || value;
}

/**
 * @param {string} [message]
 */
function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

/**
 * @param {AbortSignal | undefined} signal
 */
function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

/**
 * @template T
 * @param {AbortSignal | undefined} signal
 * @param {Promise<T>} promise
 * @returns {Promise<T>}
 */
async function withAbort(signal, promise) {
  if (!signal) return promise;
  throwIfAborted(signal);

  /** @type {(() => void) | null} */
  let removeListener = null;
  try {
    return await Promise.race([
      promise,
      new Promise((_, reject) => {
        const onAbort = () => reject(createAbortError());
        signal.addEventListener("abort", onAbort, { once: true });
        removeListener = () => signal.removeEventListener("abort", onAbort);
      }),
    ]);
  } finally {
    removeListener?.();
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
 * @param {AbortSignal | undefined} signal
 * @returns {Promise<Extract<LLMMessage, { role: "assistant" }>>}
 */
async function collectAssistantMessageFromStream(stream, onStreamEvent, signal) {
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
    throwIfAborted(signal);
    onStreamEvent?.(event);
    if (!event || typeof event !== "object") continue;

    if (event.type === "text") {
      content += event.delta ?? "";
      continue;
    }

    if (event.type === "tool_call_start") {
      const call = getOrCreateToolCall(event.id);
      call.name = normalizeToolCallName(event.name);
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
      name: normalizeToolCallName(call.name),
      arguments: args,
    });
  }

  /** @type {any} */
  const message = { role: "assistant", content };
  if (toolCalls.length) message.toolCalls = toolCalls;
  return message;
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
 *   // When true, tool execution failures are surfaced to the model as tool results
 *   // (`ok:false`) and the loop continues, allowing the model to re-plan.
 *   // Default is false (rethrow tool execution errors).
 *   continueOnToolError?: boolean,
 *   // Max size of a serialized tool result appended to the model context (in characters).
 *   // Default is 20_000.
 *   maxToolResultChars?: number,
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
  const continueOnToolError = params.continueOnToolError ?? false;
  const toolDefs = params.toolExecutor.tools ?? [];
  const maxToolResultChars =
    typeof params.maxToolResultChars === "number" && Number.isFinite(params.maxToolResultChars) && params.maxToolResultChars > 0
      ? Math.floor(params.maxToolResultChars)
      : 20_000;
  const toolNameSet = new Set(toolDefs.map((t) => t.name));

  /** @type {LLMMessage[]} */
  const messages = params.messages.slice();

  for (let i = 0; i < maxIterations; i++) {
    throwIfAborted(params.signal);
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
      params.signal,
    );

    messages.push(assistant);

    const toolCalls = assistant.toolCalls ?? [];
    if (!toolCalls.length) {
      return { messages, final: assistant.content };
    }

    let denied = false;
    let deniedToolName = null;
    for (const call of toolCalls) {
      throwIfAborted(params.signal);
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
          content: serializeToolResultForModel({ toolCall: call, result: skippedResult, maxChars: maxToolResultChars }),
        });
        continue;
      }

      if (!toolNameSet.has(call.name)) {
        const availableTools = toolDefs.map((t) => t.name);
        const baseResult = {
          tool: call.name,
          ok: false,
          error: {
            code: "unknown_tool",
            message: `Unknown tool: ${call.name}`,
          },
        };
        /** @type {any} */
        let unknownToolResult = baseResult;
        try {
          const candidate = { ...baseResult, available_tools: availableTools };
          if (JSON.stringify(candidate).length <= maxToolResultChars) {
            unknownToolResult = candidate;
          }
        } catch {
          // fall back to minimal envelope
        }
        params.onToolResult?.(call, unknownToolResult);
        messages.push({
          role: "tool",
          toolCallId: call.id,
          content: serializeToolResultForModel({ toolCall: call, result: unknownToolResult, maxChars: maxToolResultChars }),
        });
        continue;
      }

      if (requiresApproval) {
        throwIfAborted(params.signal);
        const approved = await withAbort(params.signal, requireApproval(call));
        if (!approved) {
          if (params.signal?.aborted) throw createAbortError();
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
            content: serializeToolResultForModel({ toolCall: call, result: deniedResult, maxChars: maxToolResultChars }),
          });
          if (!continueOnApprovalDenied) {
            throw new Error(deniedResult.error.message);
          }
          denied = true;
          deniedToolName = call.name;
          continue;
        }
      }

      throwIfAborted(params.signal);
      /** @type {unknown} */
      let result;
      try {
        result = await withAbort(params.signal, params.toolExecutor.execute(call));
      } catch (error) {
        if (error && typeof error === "object" && /** @type {any} */ (error).name === "AbortError") {
          throw error;
        }
        const message =
          error instanceof Error
            ? error.message
            : typeof error === "string"
              ? error
              : (() => {
                  try {
                    return JSON.stringify(error);
                  } catch {
                    return String(error);
                  }
                })();
        const errorResult = {
          tool: call.name,
          ok: false,
          error: {
            code: "tool_execution_error",
            message,
          },
        };
        params.onToolResult?.(call, errorResult);
        messages.push({
          role: "tool",
          toolCallId: call.id,
          content: serializeToolResultForModel({ toolCall: call, result: errorResult, maxChars: maxToolResultChars }),
        });
        if (!continueOnToolError) {
          throw error;
        }
        continue;
      }
      params.onToolResult?.(call, result);

      messages.push({
        role: "tool",
        toolCallId: call.id,
        content: serializeToolResultForModel({ toolCall: call, result, maxChars: maxToolResultChars }),
      });
    }
  }

  throw new Error(`Exceeded max tool-calling iterations (${maxIterations}).`);
}
