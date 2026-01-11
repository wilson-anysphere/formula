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
    const envKey = globalThis.process?.env?.OPENAI_API_KEY;
    const apiKey = options.apiKey ?? envKey;
    if (!apiKey) {
      throw new Error(
        "OpenAI API key is required. Pass `apiKey` to `new OpenAIClient({ apiKey })` or set the `OPENAI_API_KEY` environment variable in Node.js."
      );
    }

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
              totalTokens: json.usage.total_tokens,
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
   * Stream text + tool-call deltas.
   *
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
      const requestBody = {
        model: request.model ?? this.model,
        messages: toOpenAIMessages(request.messages),
        tools: request.tools?.length ? toOpenAITools(request.tools) : undefined,
        tool_choice: request.tools?.length ? request.toolChoice ?? "auto" : undefined,
        temperature: request.temperature,
        max_tokens: request.maxTokens,
        stream: true,
      };

      /**
       * @param {boolean} includeUsage
       */
      const doFetch = (includeUsage) =>
        fetch(`${this.baseUrl}/chat/completions`, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${this.apiKey}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(
            includeUsage ? { ...requestBody, stream_options: { include_usage: true } } : requestBody,
          ),
          signal: controller.signal,
        });

      let response = await doFetch(true);

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        // Some OpenAI-compatible backends (older proxies, etc.) don't support
        // `stream_options`. Retry without it so streaming still works.
        if (response.status === 400 && /stream_options/i.test(text)) {
          response = await doFetch(false);
        }
        if (!response.ok) {
          const retryText = await response.text().catch(() => "");
          throw new Error(`OpenAI streamChat error ${response.status}: ${retryText || text}`);
        }
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
      /**
       * OpenAI identifies tool calls by a stable `index` and sometimes omits the
       * `id` field on early chunks. Buffer argument fragments by index until we
       * learn the real `id` so downstream consumers can reliably key by id.
       *
       * @type {Map<number, { id?: string, name?: string, started: boolean, bufferedArgs?: string }>}
       */
      const toolCallsByIndex = new Map();
      /** @type {Set<string>} */
      const openToolCallIds = new Set();

      function closeOpenToolCalls() {
        const ids = Array.from(openToolCallIds);
        openToolCallIds.clear();
        return ids;
      }

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });

        const parts = buffer.split(/\r?\n\r?\n/);
        buffer = parts.pop() ?? "";

        for (const part of parts) {
          const lines = part.split(/\r?\n/);
          const dataLines = [];
          for (const line of lines) {
            if (!line.startsWith("data:")) continue;
            dataLines.push(line.slice("data:".length).trimStart());
          }
          if (!dataLines.length) continue;
          const data = dataLines.join("\n").trim();
          if (!data) continue;

          if (data === "[DONE]") {
            for (const id of closeOpenToolCalls()) yield { type: "tool_call_end", id };
            yield usage ? { type: "done", usage } : { type: "done" };
            return;
          }

          const json = JSON.parse(data);
          if (json.usage && typeof json.usage === "object") {
            usage = {
              promptTokens: json.usage.prompt_tokens,
              completionTokens: json.usage.completion_tokens,
              totalTokens: json.usage.total_tokens,
            };
          }
          const choice = json.choices?.[0];
          const delta = choice?.delta;

          const textDelta = delta?.content;
          if (typeof textDelta === "string" && textDelta.length > 0) {
            yield { type: "text", delta: textDelta };
          }

          const toolCalls = delta?.tool_calls;
            if (Array.isArray(toolCalls)) {
              for (const callDelta of toolCalls) {
                const index = callDelta?.index;
                if (typeof index !== "number") continue;

                const state = toolCallsByIndex.get(index) ?? { started: false };

                const idFromDelta = typeof callDelta?.id === "string" ? callDelta.id : null;
                const nameFromDelta = typeof callDelta?.function?.name === "string" ? callDelta.function.name : null;
                const argsFragment =
                  typeof callDelta?.function?.arguments === "string" ? callDelta.function.arguments : null;

                if (idFromDelta) state.id = idFromDelta;
                if (nameFromDelta) state.name = nameFromDelta;

                // If the backend hasn't provided the stable `id` yet, buffer args
                // fragments by index so we can emit them once the id arrives.
                if (!state.id && argsFragment) {
                  state.bufferedArgs = (state.bufferedArgs ?? "") + argsFragment;
                  toolCallsByIndex.set(index, state);
                  continue;
                }

                // If we just learned the id (or had it already) and there were
                // buffered fragments, flush them now.
                if (state.id && state.bufferedArgs) {
                  if (!state.started && state.name) {
                    state.started = true;
                    openToolCallIds.add(state.id);
                    yield { type: "tool_call_start", id: state.id, name: state.name };
                  }
                  yield { type: "tool_call_delta", id: state.id, delta: state.bufferedArgs };
                  state.bufferedArgs = "";
                }

                if (!state.started && state.id && state.name) {
                  state.started = true;
                  openToolCallIds.add(state.id);
                  yield { type: "tool_call_start", id: state.id, name: state.name };
                }

                toolCallsByIndex.set(index, state);

                if (state.id && argsFragment) {
                  yield { type: "tool_call_delta", id: state.id, delta: argsFragment };
                }
              }
            }

          if (typeof choice?.finish_reason === "string" && choice.finish_reason === "tool_calls") {
            for (const id of closeOpenToolCalls()) yield { type: "tool_call_end", id };
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
