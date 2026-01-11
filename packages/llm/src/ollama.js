/**
 * Minimal Ollama chat client with tool calling + streaming support.
 *
 * Ollama's `/api/chat` endpoint is OpenAI-ish (role/content messages and optional
 * tool calling). Streaming responses are newline-delimited JSON objects.
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
function toOllamaMessages(messages) {
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
function toOllamaTools(tools) {
  return tools.map((t) => ({
    type: "function",
    function: {
      name: t.name,
      description: t.description,
      parameters: t.parameters,
    },
  }));
}

export class OllamaChatClient {
  /**
   * @param {{
   *   baseUrl?: string,
   *   model?: string,
   *   timeoutMs?: number
   * }} [options]
   */
  constructor(options = {}) {
    this.model = options.model ?? "llama3.1";
    this.baseUrl = (options.baseUrl ?? "http://127.0.0.1:11434").replace(/\/$/, "");
    this.timeoutMs = options.timeoutMs ?? 30_000;
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
      /** @type {Record<string, any>} */
      const options = {};
      if (typeof request.temperature === "number") options.temperature = request.temperature;
      if (typeof request.maxTokens === "number") options.num_predict = request.maxTokens;

      const tools =
        request.tools?.length && request.toolChoice !== "none" ? toOllamaTools(request.tools) : undefined;

      const response = await fetch(`${this.baseUrl}/api/chat`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
        },
        body: JSON.stringify({
          model: request.model ?? this.model,
          messages: toOllamaMessages(request.messages),
          tools,
          stream: false,
          options: Object.keys(options).length ? options : undefined,
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(`Ollama chat error ${response.status}: ${text}`);
      }

      const json = await response.json();
      const message = json.message ?? {};
      const rawToolCalls = Array.isArray(message.tool_calls) ? message.tool_calls : [];
      const toolCalls = rawToolCalls
        .map((c, index) => {
          const id = typeof c?.id === "string" ? c.id : `toolcall-${index}`;
          const fn = c?.function;
          const name = fn?.name;
          const args = fn?.arguments;
          return {
            id,
            name,
            arguments: typeof args === "string" ? tryParseJson(args) : args ?? {},
          };
        })
        .filter((c) => typeof c.name === "string" && c.name.length > 0);

      const promptTokens = typeof json.prompt_eval_count === "number" && Number.isFinite(json.prompt_eval_count)
        ? json.prompt_eval_count
        : undefined;
      const completionTokens = typeof json.eval_count === "number" && Number.isFinite(json.eval_count)
        ? json.eval_count
        : undefined;

      return {
        message: {
          role: "assistant",
          content: message.content ?? "",
          toolCalls: toolCalls.length ? toolCalls : undefined,
        },
        usage:
          promptTokens != null || completionTokens != null
            ? {
                promptTokens: promptTokens ?? undefined,
                completionTokens: completionTokens ?? undefined,
                totalTokens:
                  typeof promptTokens === "number" && typeof completionTokens === "number"
                    ? promptTokens + completionTokens
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
      /** @type {Record<string, any>} */
      const options = {};
      if (typeof request.temperature === "number") options.temperature = request.temperature;
      if (typeof request.maxTokens === "number") options.num_predict = request.maxTokens;

      const tools =
        request.tools?.length && request.toolChoice !== "none" ? toOllamaTools(request.tools) : undefined;

      const response = await fetch(`${this.baseUrl}/api/chat`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
        },
        body: JSON.stringify({
          model: request.model ?? this.model,
          messages: toOllamaMessages(request.messages),
          tools,
          stream: true,
          options: Object.keys(options).length ? options : undefined,
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(`Ollama streamChat error ${response.status}: ${text}`);
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
        yield { type: "done" };
        return;
      }

      const decoder = new TextDecoder();
      let buffer = "";
      /** @type {{ promptTokens?: number, completionTokens?: number, totalTokens?: number } | null} */
      let usage = null;

      /** @type {Map<string, { name?: string, args: string, started: boolean }>} */
      const toolCallsById = new Map();
      /** @type {Set<string>} */
      const openToolCalls = new Set();

      function closeOpenToolCalls() {
        const ids = Array.from(openToolCalls);
        openToolCalls.clear();
        return ids;
      }

      function getOrCreateToolCall(id) {
        const existing = toolCallsById.get(id);
        if (existing) return existing;
        const next = { args: "", started: false };
        toolCallsById.set(id, next);
        return next;
      }

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });

        const lines = buffer.split(/\r?\n/);
        buffer = lines.pop() ?? "";

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed) continue;
          const json = JSON.parse(trimmed);
          const msg = json.message;

          const contentDelta = msg?.content;
          if (typeof contentDelta === "string" && contentDelta.length > 0) {
            yield { type: "text", delta: contentDelta };
          }

          const toolCalls = msg?.tool_calls;
          if (Array.isArray(toolCalls)) {
            for (let idx = 0; idx < toolCalls.length; idx++) {
              const c = toolCalls[idx];
              const id = typeof c.id === "string" ? c.id : `toolcall-${idx}`;
              const name = typeof c.function?.name === "string" ? c.function.name : undefined;
              const args = typeof c.function?.arguments === "string" ? c.function.arguments : "";

              const state = getOrCreateToolCall(id);
              if (name) state.name = name;

              if (!state.started && state.name) {
                state.started = true;
                openToolCalls.add(id);
                yield { type: "tool_call_start", id, name: state.name };
              }

              if (typeof args === "string" && args.length > 0) {
                const prev = state.args;
                // Best-effort diffing: Ollama may stream the full argument string repeatedly.
                const delta = args.startsWith(prev) ? args.slice(prev.length) : args;
                state.args = args;
                if (delta) yield { type: "tool_call_delta", id, delta };
              }
            }
          }

          if (json.done) {
            for (const id of closeOpenToolCalls()) yield { type: "tool_call_end", id };
            const promptTokens = json.prompt_eval_count;
            const completionTokens = json.eval_count;
            if (typeof promptTokens === "number" || typeof completionTokens === "number") {
              usage = {
                promptTokens: typeof promptTokens === "number" ? promptTokens : undefined,
                completionTokens: typeof completionTokens === "number" ? completionTokens : undefined,
                totalTokens:
                  typeof promptTokens === "number" && typeof completionTokens === "number"
                    ? promptTokens + completionTokens
                    : undefined,
              };
            }
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
