/**
 * Cursor-only LLM client (dependency-free).
 *
 * The Cursor backend speaks an OpenAI-compatible Chat Completions protocol:
 * `POST /chat/completions` (typically under a `/v1` base path).
 *
 * This client is intentionally *not* provider-selectable and does *not* read
 * user API keys from environment variables. Callers must inject auth via
 * `getAuthHeaders` / `authToken` (and/or rely on cookies when same-origin).
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

/**
 * Read a Vite `import.meta.env.*` variable if present.
 *
 * @param {string} key
 * @returns {string | undefined}
 */
function readViteEnv(key) {
  try {
    const env = /** @type {any} */ (import.meta).env;
    const value = env?.[key];
    if (typeof value === "string") {
      const trimmed = value.trim();
      if (trimmed) return trimmed;
    }
  } catch {
    // Not running under Vite (or import.meta.env unavailable).
  }
  return undefined;
}

/**
 * Read a Node `process.env.*` variable if present.
 *
 * NOTE: This is only used for Cursor-specific configuration (base URL / timeouts),
 * not for provider API keys.
 *
 * @param {string} key
 * @returns {string | undefined}
 */
function readNodeEnv(key) {
  const env = globalThis.process?.env;
  const value = env?.[key];
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed) return trimmed;
  }
  return undefined;
}

/**
 * @param {string | undefined} explicitBaseUrl
 */
function resolveBaseUrl(explicitBaseUrl) {
  const raw =
    explicitBaseUrl ?? readNodeEnv("CURSOR_AI_BASE_URL") ?? readViteEnv("VITE_CURSOR_AI_BASE_URL") ?? "";
  return String(raw).replace(/\/$/, "");
}

/**
 * @param {number | undefined} explicitTimeoutMs
 */
function resolveTimeoutMs(explicitTimeoutMs) {
  if (typeof explicitTimeoutMs === "number") return explicitTimeoutMs;
  const raw = readNodeEnv("CURSOR_AI_TIMEOUT_MS") ?? readViteEnv("VITE_CURSOR_AI_TIMEOUT_MS");
  if (typeof raw === "string" && raw.length > 0) {
    const parsed = Number(raw);
    if (Number.isFinite(parsed) && parsed > 0) return parsed;
  }
  return 30_000;
}

/**
 * Compute the OpenAI-compatible chat completions endpoint from a configured base URL.
 *
 * Examples:
 * - `https://cursor.test` => `https://cursor.test/v1/chat/completions`
 * - `https://cursor.test/v1` => `https://cursor.test/v1/chat/completions`
 * - `https://cursor.test/v1/chat` => `https://cursor.test/v1/chat/completions`
 * - `""` (same-origin) => `/v1/chat/completions`
 *
 * @param {string} baseUrl
 */
function chatCompletionsEndpointFromBaseUrl(baseUrl) {
  const trimmed = String(baseUrl ?? "").replace(/\/$/, "");
  if (!trimmed) return "/v1/chat/completions";
  if (trimmed.endsWith("/chat/completions")) return trimmed;
  if (trimmed.endsWith("/v1/chat")) return `${trimmed}/completions`;
  if (trimmed.endsWith("/v1")) return `${trimmed}/chat/completions`;
  return `${trimmed}/v1/chat/completions`;
}

/**
 * @param {AbortSignal | undefined} requestSignal
 * @param {AbortController} controller
 */
function forwardAbortSignal(requestSignal, controller) {
  /** @type {(() => void) | null} */
  let removeListener = null;
  if (!requestSignal) return removeListener;

  if (requestSignal.aborted) {
    controller.abort();
    return removeListener;
  }

  if (typeof requestSignal.addEventListener === "function") {
    const onAbort = () => controller.abort();
    requestSignal.addEventListener("abort", onAbort, { once: true });
    removeListener = () => requestSignal.removeEventListener("abort", onAbort);
  }

  return removeListener;
}

export class CursorLLMClient {
  /**
   * @param {{
   *   baseUrl?: string,
   *   model?: string,
   *   timeoutMs?: number,
   *   authToken?: string,
   *   getAuthHeaders?: () => (Record<string, string> | Promise<Record<string, string>>)
   * }} [options]
   */
  constructor(options = {}) {
    this.baseUrl = resolveBaseUrl(options.baseUrl);
    this.chatCompletionsEndpoint = chatCompletionsEndpointFromBaseUrl(this.baseUrl);
    this.timeoutMs = resolveTimeoutMs(options.timeoutMs);

    // Cursor backend controls model routing. `model` is treated as an optional
    // hint that may be ignored by the backend.
    this.model = options.model;

    this.authToken = options.authToken;
    this.getAuthHeaders = options.getAuthHeaders;
  }

  /**
   * @returns {Promise<Record<string, string>>}
   */
  async _resolveHeaders() {
    /** @type {Record<string, string>} */
    const headers = { "Content-Type": "application/json" };

    if (this.authToken) {
      headers.Authorization = `Bearer ${this.authToken}`;
    }

    // Prefer `getAuthHeaders()` so Cursor can manage authentication. If it
    // returns an Authorization header, it should override `authToken`.
    const extra = this.getAuthHeaders ? await this.getAuthHeaders() : null;
    if (extra && typeof extra === "object") {
      for (const [key, value] of Object.entries(extra)) {
        if (typeof value !== "string") continue;
        headers[key] = value;
      }
    }

    // Always ensure JSON content type.
    headers["Content-Type"] = "application/json";

    return headers;
  }

  /**
   * @param {import("./types.js").ChatRequest} request
   * @returns {Promise<import("./types.js").ChatResponse>}
   */
  async chat(request) {
    const controller = new AbortController();
    const removeRequestAbortListener = forwardAbortSignal(request?.signal, controller);
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const response = await fetch(this.chatCompletionsEndpoint, {
        method: "POST",
        headers: await this._resolveHeaders(),
        // Allow cookie-based auth in browser/desktop runtimes when a proxy is
        // used. Node runtimes typically ignore cookies anyway.
        credentials: "include",
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
        throw new Error(`Cursor chat error ${response.status}: ${text}`);
      }

      const json = await response.json();
      const message = json.choices?.[0]?.message;
      const rawToolCalls = Array.isArray(message?.tool_calls) ? message.tool_calls : [];
      const toolCalls = rawToolCalls
        .map((c, index) => ({
          id: typeof c?.id === "string" ? c.id : `toolcall-${index}`,
          name: c?.function?.name,
          arguments: tryParseJson(c?.function?.arguments ?? "{}"),
        }))
        .filter((c) => typeof c.name === "string" && c.name.length > 0);

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
    const removeRequestAbortListener = forwardAbortSignal(request?.signal, controller);
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);

    try {
      const headers = await this._resolveHeaders();

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
        fetch(this.chatCompletionsEndpoint, {
          method: "POST",
          headers,
          credentials: "include",
          body: JSON.stringify(
            includeUsage ? { ...requestBody, stream_options: { include_usage: true } } : requestBody,
          ),
          signal: controller.signal,
        });

      let response = await doFetch(true);

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        // Some OpenAI-compatible backends don't support `stream_options`. Retry
        // without it so streaming still works.
        if (response.status === 400 && /stream_options/i.test(text)) {
          response = await doFetch(false);
        }
        if (!response.ok) {
          const retryText = await response.text().catch(() => "");
          throw new Error(`Cursor streamChat error ${response.status}: ${retryText || text}`);
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
       * learn the stable `id` + `name`, then emit `tool_call_start` followed by
       * the buffered `tool_call_delta` fragments.
       *
       * Some OpenAI-compatible backends incorrectly stream the full arguments
       * string repeatedly (instead of deltas). Track the reconstructed argument
       * string so we can diff and only emit the incremental suffix.
       *
       * @type {Map<number, { id?: string, name?: string, started: boolean, pendingArgs: string, args: string }>}
       */
      const toolCallsByIndex = new Map();
      /** @type {Set<string>} */
      const openToolCallIds = new Set();
      let nextToolCallIndexToStart = 0;

      function closeOpenToolCalls() {
        const ids = Array.from(openToolCallIds);
        openToolCallIds.clear();
        return ids;
      }

      /**
       * OpenAI tool calls are logically ordered by the numeric `index` field.
       * Some backends can emit tool call chunks out of order; only emit tool call
       * start events once we've seen contiguous indexes starting from 0 so the
       * downstream tool loop executes calls in a deterministic order.
       */
      function* startToolCallsInOrder() {
        while (true) {
          const state = toolCallsByIndex.get(nextToolCallIndexToStart);
          if (!state) break;
          if (state.started) {
            nextToolCallIndexToStart += 1;
            continue;
          }
          if (!state.id || !state.name) break;

          state.started = true;
          openToolCallIds.add(state.id);
          yield { type: "tool_call_start", id: state.id, name: state.name };
          if (state.pendingArgs) {
            yield { type: "tool_call_delta", id: state.id, delta: state.pendingArgs };
            state.pendingArgs = "";
          }
          nextToolCallIndexToStart += 1;
        }
      }

      function* flushPendingToolCalls() {
        const indexes = Array.from(toolCallsByIndex.keys()).sort((a, b) => a - b);
        for (const index of indexes) {
          const state = toolCallsByIndex.get(index);
          if (!state || typeof state !== "object") continue;
          if (state.started) continue;
          if (!state.id && state.name) state.id = `toolcall-${index}`;
          if (!state.started && state.id && state.name) {
            state.started = true;
            openToolCallIds.add(state.id);
            yield { type: "tool_call_start", id: state.id, name: state.name };
            if (state.pendingArgs) {
              yield { type: "tool_call_delta", id: state.id, delta: state.pendingArgs };
              state.pendingArgs = "";
            }
          }
        }
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
            for (const event of flushPendingToolCalls()) yield event;
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

              const state = toolCallsByIndex.get(index) ?? { started: false, pendingArgs: "", args: "" };
              const idFromDelta = typeof callDelta?.id === "string" ? callDelta.id : null;
              const nameFromDelta = typeof callDelta?.function?.name === "string" ? callDelta.function.name : null;
              const argsFragment =
                typeof callDelta?.function?.arguments === "string" ? callDelta.function.arguments : null;

              if (idFromDelta) state.id = idFromDelta;
              if (nameFromDelta) state.name = nameFromDelta;

              toolCallsByIndex.set(index, state);

              if (argsFragment) {
                // Best-effort diffing: tolerate backends that repeatedly send the
                // full argument string instead of deltas.
                let deltaArgs = argsFragment;
                if (typeof state.args === "string" && argsFragment.startsWith(state.args)) {
                  deltaArgs = argsFragment.slice(state.args.length);
                  state.args = argsFragment;
                } else {
                  state.args = (state.args ?? "") + argsFragment;
                }

                if (deltaArgs) {
                  if (state.id && state.started) {
                    yield { type: "tool_call_delta", id: state.id, delta: deltaArgs };
                  } else {
                    state.pendingArgs += deltaArgs;
                  }
                }
              }
            }
            for (const event of startToolCallsInOrder()) yield event;
          }

          if (typeof choice?.finish_reason === "string" && choice.finish_reason === "tool_calls") {
            for (const event of flushPendingToolCalls()) yield event;
            for (const id of closeOpenToolCalls()) yield { type: "tool_call_end", id };
          }
        }
      }

      for (const event of flushPendingToolCalls()) yield event;
      for (const id of closeOpenToolCalls()) yield { type: "tool_call_end", id };
      yield usage ? { type: "done", usage } : { type: "done" };
    } finally {
      clearTimeout(timeout);
      removeRequestAbortListener?.();
    }
  }
}

