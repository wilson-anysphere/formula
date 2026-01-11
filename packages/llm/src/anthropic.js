/**
 * Minimal Anthropic Messages API client with tool calling + streaming support.
 *
 * Note: this package keeps the surface small and dependency-free. It uses the
 * global `fetch` available in modern Node and browsers.
 */

/**
 * @param {import("./types.js").LLMMessage[]} messages
 */
function extractSystemPrompt(messages) {
  const systemMessages = messages.filter((m) => m.role === "system").map((m) => m.content);
  return systemMessages.join("\n\n").trim();
}

/**
 * @param {unknown} value
 */
function tryParseJson(value) {
  if (typeof value !== "string") return value;
  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}

/**
 * Anthropic `tool_use.input` must be an object; if callers provide a JSON string
 * (common for OpenAI-compatible tool calls) decode it.
 *
 * @param {unknown} value
 */
function toToolInput(value) {
  if (value == null) return {};
  if (typeof value === "object") return value;
  const parsed = tryParseJson(value);
  if (parsed && typeof parsed === "object") return parsed;
  return { value: parsed };
}

/**
 * @param {import("./types.js").LLMMessage[]} messages
 */
function toAnthropicMessages(messages) {
  /** @type {any[]} */
  const out = [];

  for (const m of messages) {
    if (m.role === "system") continue;

    if (m.role === "tool") {
      out.push({
        role: "user",
        content: [
          {
            type: "tool_result",
            tool_use_id: m.toolCallId,
            content: m.content ?? "",
          },
        ],
      });
      continue;
    }

    if (m.role === "assistant" && m.toolCalls?.length) {
      /** @type {any[]} */
      const content = [];
      if (m.content) content.push({ type: "text", text: m.content });
      for (const c of m.toolCalls) {
        content.push({ type: "tool_use", id: c.id, name: c.name, input: toToolInput(c.arguments) });
      }
      out.push({ role: "assistant", content });
      continue;
    }

    out.push({ role: m.role, content: m.content ?? "" });
  }

  return out;
}

/**
 * @param {import("./types.js").ToolDefinition[]} tools
 */
function toAnthropicTools(tools) {
  return tools.map((t) => ({
    name: t.name,
    description: t.description,
    input_schema: t.parameters,
  }));
}

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
 * Parse SSE events from a text buffer. Returns the complete events and the remaining buffer.
 *
 * @param {string} buffer
 */
function splitSSEEvents(buffer) {
  const parts = buffer.split(/\r?\n\r?\n/);
  return { events: parts.slice(0, -1), rest: parts.at(-1) ?? "" };
}

/**
 * @param {string} sseEvent
 */
function extractSSEData(sseEvent) {
  const lines = sseEvent.split(/\r?\n/);
  const dataLines = [];
  for (const line of lines) {
    if (!line.startsWith("data:")) continue;
    dataLines.push(line.slice("data:".length).trimStart());
  }
  if (!dataLines.length) return null;
  return dataLines.join("\n").trim();
}

export class AnthropicClient {
  /**
  * @param {{
  *   apiKey?: string,
  *   model?: string,
  *   baseUrl?: string,
  *   timeoutMs?: number,
  *   maxTokens?: number
  * }} [options]
  */
  constructor(options = {}) {
    const envKey = globalThis.process?.env?.ANTHROPIC_API_KEY;
    const apiKey = options.apiKey ?? envKey;
    if (!apiKey) {
      throw new Error(
        "Anthropic API key is required. Pass `apiKey` to `new AnthropicClient({ apiKey })` or set the `ANTHROPIC_API_KEY` environment variable in Node.js.",
      );
    }

    this.apiKey = apiKey;
    this.model = options.model ?? "claude-3-5-sonnet-latest";
    this.baseUrl = (options.baseUrl ?? "https://api.anthropic.com/v1").replace(/\/$/, "");
    this.timeoutMs = options.timeoutMs ?? 30_000;
    this.maxTokens = options.maxTokens ?? 1024;
  }

  /**
   * @param {import("./types.js").ChatRequest} request
   * @returns {Promise<import("./types.js").ChatResponse>}
   */
  async chat(request) {
    const controller = new AbortController();
    const requestSignal = request?.signal;
    /** @type {(() => void) | null} */
    let removeRequestAbortListener = null;
    if (requestSignal) {
      if (requestSignal.aborted) {
        controller.abort();
      } else if (typeof requestSignal.addEventListener === "function") {
        const onAbort = () => controller.abort();
        requestSignal.addEventListener("abort", onAbort, { once: true });
        removeRequestAbortListener = () => requestSignal.removeEventListener("abort", onAbort);
      }
    }
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);

    try {
      const system = extractSystemPrompt(request.messages);
      const response = await fetch(`${this.baseUrl}/messages`, {
        method: "POST",
        headers: {
          "x-api-key": this.apiKey,
          "anthropic-version": "2023-06-01",
          "content-type": "application/json",
        },
        body: JSON.stringify({
          model: request.model ?? this.model,
          system: system || undefined,
          messages: toAnthropicMessages(request.messages),
          tools: request.tools?.length ? toAnthropicTools(request.tools) : undefined,
          tool_choice: request.tools?.length
            ? request.toolChoice === "none"
              ? { type: "none" }
              : { type: "auto" }
            : undefined,
          max_tokens: request.maxTokens ?? this.maxTokens,
          temperature: request.temperature,
          stream: false,
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(`Anthropic chat error ${response.status}: ${text}`);
      }

      const json = await response.json();
      const blocks = Array.isArray(json.content) ? json.content : [];
      const text = blocks
        .filter((b) => b?.type === "text" && typeof b.text === "string")
        .map((b) => b.text)
        .join("");

      const toolCalls = blocks
        .filter((b) => b?.type === "tool_use" && typeof b.id === "string" && typeof b.name === "string")
        .map((b) => ({ id: b.id, name: b.name, arguments: b.input ?? {} }));

      return {
        message: {
          role: "assistant",
          content: text,
          toolCalls: toolCalls.length ? toolCalls : undefined,
        },
        usage: json.usage
          ? {
              promptTokens: json.usage.input_tokens,
              completionTokens: json.usage.output_tokens,
              totalTokens:
                typeof json.usage.input_tokens === "number" && typeof json.usage.output_tokens === "number"
                  ? json.usage.input_tokens + json.usage.output_tokens
                  : undefined,
            }
          : undefined,
        raw: json,
      };
    } finally {
      clearTimeout(timeout);
      removeRequestAbortListener?.();
    }
  }

  /**
   * @param {import("./types.js").ChatRequest} request
   * @returns {AsyncIterable<import("./types.js").ChatStreamEvent>}
   */
  async *streamChat(request) {
    const controller = new AbortController();
    const requestSignal = request?.signal;
    /** @type {(() => void) | null} */
    let removeRequestAbortListener = null;
    if (requestSignal) {
      if (requestSignal.aborted) {
        controller.abort();
      } else if (typeof requestSignal.addEventListener === "function") {
        const onAbort = () => controller.abort();
        requestSignal.addEventListener("abort", onAbort, { once: true });
        removeRequestAbortListener = () => requestSignal.removeEventListener("abort", onAbort);
      }
    }
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);

    try {
      const system = extractSystemPrompt(request.messages);
      const response = await fetch(`${this.baseUrl}/messages`, {
        method: "POST",
        headers: {
          "x-api-key": this.apiKey,
          "anthropic-version": "2023-06-01",
          "content-type": "application/json",
        },
        body: JSON.stringify({
          model: request.model ?? this.model,
          system: system || undefined,
          messages: toAnthropicMessages(request.messages),
          tools: request.tools?.length ? toAnthropicTools(request.tools) : undefined,
          tool_choice: request.tools?.length
            ? request.toolChoice === "none"
              ? { type: "none" }
              : { type: "auto" }
            : undefined,
          max_tokens: request.maxTokens ?? this.maxTokens,
          temperature: request.temperature,
          stream: true,
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(`Anthropic streamChat error ${response.status}: ${text}`);
      }

      const reader = response.body?.getReader();
      if (!reader) {
        const full = await this.chat({ ...request, signal: controller.signal });
        const text = full.message.content ?? "";
        if (text) yield { type: "text", delta: text };
        for (const call of full.message.toolCalls ?? []) {
          yield { type: "tool_call_start", id: call.id, name: call.name };
          const args = toJsonString(call.arguments ?? {});
          if (args) yield { type: "tool_call_delta", id: call.id, delta: args };
          yield { type: "tool_call_end", id: call.id };
        }
        const usage = full.usage
          ? {
              promptTokens: full.usage.promptTokens,
              completionTokens: full.usage.completionTokens,
              totalTokens:
                full.usage.promptTokens != null && full.usage.completionTokens != null
                  ? full.usage.promptTokens + full.usage.completionTokens
                  : undefined,
            }
          : undefined;
        yield usage ? { type: "done", usage } : { type: "done" };
        return;
      }

      const decoder = new TextDecoder();
      let buffer = "";
      /** @type {{ promptTokens?: number, completionTokens?: number, totalTokens?: number } | null} */
      let usage = null;

      /** @type {Map<number, { id: string, type: "text" | "tool_use" }>} */
      const blocksByIndex = new Map();
      /** @type {Set<string>} */
      const openToolCalls = new Set();

      function closeOpenToolCalls() {
        const ids = Array.from(openToolCalls);
        openToolCalls.clear();
        return ids;
      }

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });

        const { events, rest } = splitSSEEvents(buffer);
        buffer = rest;

        for (const ev of events) {
          const data = extractSSEData(ev);
          if (!data) continue;
          const json = JSON.parse(data);
 
          const maybeUsage = json?.message?.usage ?? json?.usage;
          if (maybeUsage && typeof maybeUsage === "object") {
            const next = usage ?? {};
            if (maybeUsage.input_tokens != null) next.promptTokens = maybeUsage.input_tokens;
            if (maybeUsage.output_tokens != null) next.completionTokens = maybeUsage.output_tokens;
            if (next.promptTokens != null && next.completionTokens != null) {
              next.totalTokens = next.promptTokens + next.completionTokens;
            }
            usage = next;
          }

          if (json.type === "content_block_start") {
            const index = json.index;
            const block = json.content_block;
            if (typeof index === "number" && block?.type === "tool_use" && typeof block.id === "string") {
              blocksByIndex.set(index, { id: block.id, type: "tool_use" });
              if (typeof block.name === "string") {
                openToolCalls.add(block.id);
                yield { type: "tool_call_start", id: block.id, name: block.name };
              }
              if (block.input && typeof block.input === "object") {
                const encoded = toJsonString(block.input);
                if (encoded && encoded !== "{}") {
                  yield { type: "tool_call_delta", id: block.id, delta: encoded };
                }
              }
            } else if (typeof index === "number" && block?.type === "text") {
              blocksByIndex.set(index, { id: `text-${index}`, type: "text" });
            }
            continue;
          }

          if (json.type === "content_block_delta") {
            const index = json.index;
            const delta = json.delta;
            if (delta?.type === "text_delta" && typeof delta.text === "string") {
              yield { type: "text", delta: delta.text };
              continue;
            }

            if (delta?.type === "input_json_delta" && typeof delta.partial_json === "string") {
              const block = typeof index === "number" ? blocksByIndex.get(index) : null;
              if (block?.type === "tool_use") {
                yield { type: "tool_call_delta", id: block.id, delta: delta.partial_json };
              }
              continue;
            }
          }

          if (json.type === "content_block_stop") {
            const index = json.index;
            const block = typeof index === "number" ? blocksByIndex.get(index) : null;
            if (block?.type === "tool_use") {
              openToolCalls.delete(block.id);
              yield { type: "tool_call_end", id: block.id };
            }
            continue;
          }

          if (json.type === "message_stop") {
            for (const id of closeOpenToolCalls()) yield { type: "tool_call_end", id };
            yield usage ? { type: "done", usage } : { type: "done" };
            return;
          }
        }
      }

      for (const id of closeOpenToolCalls()) yield { type: "tool_call_end", id };
      yield usage ? { type: "done", usage } : { type: "done" };
    } finally {
      clearTimeout(timeout);
      removeRequestAbortListener?.();
    }
  }
}
