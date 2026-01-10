/**
 * Provider-agnostic tool calling loop.
 *
 * The LLM client is responsible for translating `ToolDefinition`s + messages
 * into provider-specific APIs (OpenAI, Anthropic, etc). This module implements
 * the generic loop: call LLM -> execute tools -> call LLM -> …
 */

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
 *   model?: string,
 *   temperature?: number,
 *   maxTokens?: number
 * }} params
 */
export async function runChatWithTools(params) {
  const maxIterations = params.maxIterations ?? 8;
  const requireApproval = params.requireApproval ?? (async () => true);
  const toolDefs = params.toolExecutor.tools ?? [];

  /** @type {LLMMessage[]} */
  const messages = params.messages.slice();

  for (let i = 0; i < maxIterations; i++) {
    const response = await params.client.chat({
      messages,
      tools: toolDefs,
      toolChoice: toolDefs.length ? "auto" : "none",
      model: params.model,
      temperature: params.temperature,
      maxTokens: params.maxTokens,
    });

    messages.push(response.message);

    const toolCalls = response.message.toolCalls ?? [];
    if (!toolCalls.length) {
      return {
        messages,
        final: response.message.content,
      };
    }

    for (const call of toolCalls) {
      const requiresApproval = toolRequiresApproval(toolDefs, call.name);
      params.onToolCall?.(call, { requiresApproval });

      if (requiresApproval) {
        const approved = await requireApproval(call);
        if (!approved) {
          throw new Error(`Tool call requires approval and was denied: ${call.name}`);
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
