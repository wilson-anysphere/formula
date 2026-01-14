/**
 * Provider-agnostic tool calling loop.
 *
 * The LLM client is responsible for translating `ToolDefinition`s + messages
 * into the underlying inference API. This module implements the generic loop:
 * call LLM -> execute tools -> call LLM -> …
 *
 * SECURITY NOTE: Tool results are appended to the conversation as `role: "tool"`
 * messages and are sent back to the model on the next iteration. Any sensitive
 * data controls (permissions/DLP/redaction) must be enforced by the tool executor
 * before returning a result, not just when constructing the initial prompt context.
 */

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
 */

/**
 * @param {{
 *   client: LLMClient,
 *   toolExecutor: ToolExecutor,
 *   messages: LLMMessage[],
 *   maxIterations?: number,
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
export async function runChatWithTools(params) {
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
    const response = await withAbort(
      params.signal,
      params.client.chat({
        messages,
        tools: toolDefs,
        toolChoice: toolDefs.length ? "auto" : "none",
        model: params.model,
        temperature: params.temperature,
        maxTokens: params.maxTokens,
        signal: params.signal,
      }),
    );

    // Normalize tool call names defensively (some providers/models may include
    // leading/trailing whitespace, which would break tool dispatch).
    if (Array.isArray(response.message?.toolCalls)) {
      for (const call of response.message.toolCalls) {
        const normalized = normalizeToolCallName(call?.name);
        if (normalized && normalized !== call.name) call.name = normalized;
      }
    }

    messages.push(response.message);

    const toolCalls = response.message.toolCalls ?? [];
    if (!toolCalls.length) {
      return {
        messages,
        final: response.message.content,
      };
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
          // Only include `available_tools` if the full list fits within the budget.
          if (JSON.stringify(candidate).length <= maxToolResultChars) {
            unknownToolResult = candidate;
          }
        } catch {
          // If stringification fails for any reason, fall back to the minimal envelope.
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
        // Preserve AbortSignal semantics.
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
