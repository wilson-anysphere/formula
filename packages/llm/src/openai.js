/**
 * Minimal OpenAI Chat Completions client with tool calling support.
 *
 * Note: this package keeps the surface small and dependency-free. It uses the
 * global `fetch` available in modern Node and browsers.
 */

/**
 * @param {unknown} value
 */
function toJsonString(value) {
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

/**
 * @param {string} input
 */
function tryParseJson(input) {
  try {
    return JSON.parse(input);
  } catch {
    return input;
  }
}

/**
 * @param {import("./types.js").LLMMessage[]} messages
 */
function toOpenAIMessages(messages) {
  return messages.map((m) => {
    if (m.role === "tool") {
      return {
        role: "tool",
        tool_call_id: m.toolCallId,
        content: m.content,
      };
    }

    if (m.role === "assistant" && m.toolCalls?.length) {
      return {
        role: "assistant",
        content: m.content ?? "",
        tool_calls: m.toolCalls.map((c) => ({
          id: c.id,
          type: "function",
          function: {
            name: c.name,
            arguments: toJsonString(c.arguments),
          },
        })),
      };
    }

    return { role: m.role, content: m.content };
  });
}

/**
 * @param {import("./types.js").ToolDefinition[]} tools
 */
function toOpenAITools(tools) {
  return tools.map((t) => ({
    type: "function",
    function: {
      name: t.name,
      description: t.description,
      parameters: t.parameters,
    },
  }));
}

export class OpenAIClient {
  /**
   * @param {{
   *   apiKey?: string,
   *   model?: string,
   *   baseUrl?: string,
   *   timeoutMs?: number
   * }} [options]
   */
  constructor(options = {}) {
    const apiKey = options.apiKey ?? process.env.OPENAI_API_KEY;
    if (!apiKey) throw new Error("OPENAI_API_KEY is required to use OpenAIClient");

    this.apiKey = apiKey;
    this.model = options.model ?? "gpt-4o-mini";
    this.baseUrl = options.baseUrl ?? "https://api.openai.com/v1";
    this.timeoutMs = options.timeoutMs ?? 30_000;
  }

  /**
   * @param {import("./types.js").ChatRequest} request
   * @returns {Promise<import("./types.js").ChatResponse>}
   */
  async chat(request) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const response = await fetch(`${this.baseUrl}/chat/completions`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${this.apiKey}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model: request.model ?? this.model,
          messages: toOpenAIMessages(request.messages),
          tools: request.tools?.length ? toOpenAITools(request.tools) : undefined,
          tool_choice: request.tools?.length ? request.toolChoice ?? "auto" : undefined,
          temperature: request.temperature,
          max_tokens: request.maxTokens,
          stream: false,
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(`OpenAI chat error ${response.status}: ${text}`);
      }

      const json = await response.json();
      const message = json.choices?.[0]?.message;
      const toolCalls = (message?.tool_calls ?? []).map((c) => ({
        id: c.id,
        name: c.function?.name,
        arguments: tryParseJson(c.function?.arguments ?? "{}"),
      }));

      return {
        message: {
          role: "assistant",
          content: message?.content ?? "",
          toolCalls: toolCalls.length ? toolCalls : undefined,
        },
        usage: json.usage
          ? {
              promptTokens: json.usage.prompt_tokens,
              completionTokens: json.usage.completion_tokens,
            }
          : undefined,
        raw: json,
      };
    } finally {
      clearTimeout(timeout);
    }
  }

  /**
   * Stream text deltas (tool call deltas are ignored for now).
   *
   * @param {import("./types.js").ChatRequest} request
   * @returns {AsyncIterable<import("./types.js").ChatStreamEvent>}
   */
  async *streamChat(request) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const response = await fetch(`${this.baseUrl}/chat/completions`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${this.apiKey}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model: request.model ?? this.model,
          messages: toOpenAIMessages(request.messages),
          tools: request.tools?.length ? toOpenAITools(request.tools) : undefined,
          tool_choice: request.tools?.length ? request.toolChoice ?? "auto" : undefined,
          temperature: request.temperature,
          max_tokens: request.maxTokens,
          stream: true,
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(`OpenAI streamChat error ${response.status}: ${text}`);
      }

      const reader = response.body?.getReader();
      if (!reader) return;

      const decoder = new TextDecoder();
      let buffer = "";

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });

        const parts = buffer.split("\n\n");
        buffer = parts.pop() ?? "";

        for (const part of parts) {
          const line = part.trim();
          if (!line.startsWith("data:")) continue;
          const data = line.slice("data:".length).trim();
          if (data === "[DONE]") return;
          const json = JSON.parse(data);
          const delta = json.choices?.[0]?.delta?.content;
          if (delta) yield { type: "text", delta };
        }
      }
    } finally {
      clearTimeout(timeout);
    }
  }
}
