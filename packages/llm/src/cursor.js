/**
 * Cursor backend LLM client.
 *
 * This package intentionally does *not* support direct provider API keys.
 * Authentication is handled by the Cursor session (cookies), so requests use
 * `credentials: "include"`.
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

function getBaseUrl() {
  const raw = readNodeEnv("CURSOR_AI_BASE_URL") ?? readViteEnv("VITE_CURSOR_AI_BASE_URL") ?? "";
  return raw.replace(/\/$/, "");
}

function getTimeoutMs() {
  const raw = readNodeEnv("CURSOR_AI_TIMEOUT_MS") ?? readViteEnv("VITE_CURSOR_AI_TIMEOUT_MS");
  if (typeof raw === "string" && raw.length > 0) {
    const parsed = Number(raw);
    if (Number.isFinite(parsed) && parsed > 0) return parsed;
  }
  return 30_000;
}

/**
 * @param {any} json
 * @returns {import("./types.js").ChatResponse}
 */
function parseChatResponse(json) {
  const message = json?.message;
  const usage = json?.usage;

  if (!message || typeof message !== "object") {
    throw new Error("Cursor LLM chat error: response missing `message`");
  }

  return {
    message: {
      role: "assistant",
      content: typeof message.content === "string" ? message.content : "",
      toolCalls: Array.isArray(message.toolCalls) ? message.toolCalls : undefined,
    },
    usage: usage && typeof usage === "object" ? usage : undefined,
    raw: json?.raw ?? json,
  };
}

/**
 * @param {any} value
 */
function isChatResponseLike(value) {
  if (!value || typeof value !== "object") return false;
  const message = value.message;
  if (!message || typeof message !== "object") return false;
  if (typeof message.content === "string") return true;
  if (Array.isArray(message.toolCalls)) return true;
  return false;
}

/**
 * @param {any} value
 * @returns {value is import("./types.js").ChatStreamEvent}
 */
function isChatStreamEvent(value) {
  if (!value || typeof value !== "object") return false;
  const type = value.type;
  if (type === "text") return typeof value.delta === "string";
  if (type === "tool_call_start") return typeof value.id === "string" && typeof value.name === "string";
  if (type === "tool_call_delta") return typeof value.id === "string" && typeof value.delta === "string";
  if (type === "tool_call_end") return typeof value.id === "string";
  if (type === "done") return true;
  return false;
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

/**
 * @param {import("./types.js").ChatResponse} full
 */
function* synthesizeStreamEvents(full) {
  const text = full?.message?.content ?? "";
  if (typeof text === "string" && text.length > 0) {
    yield { type: "text", delta: text };
  }

  const toolCalls = full?.message?.toolCalls;
  if (Array.isArray(toolCalls)) {
    for (const call of toolCalls) {
      if (!call || typeof call !== "object") continue;
      if (typeof call.id !== "string" || typeof call.name !== "string") continue;
      yield { type: "tool_call_start", id: call.id, name: call.name };
      const args = toJsonString(call.arguments ?? {});
      if (args) yield { type: "tool_call_delta", id: call.id, delta: args };
      yield { type: "tool_call_end", id: call.id };
    }
  }

  if (full?.usage) {
    yield { type: "done", usage: full.usage };
  } else {
    yield { type: "done" };
  }
}

export class CursorLLMClient {
  constructor() {
    this.baseUrl = getBaseUrl();
    this.timeoutMs = getTimeoutMs();
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
      const response = await fetch(`${this.baseUrl}/v1/chat`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          accept: "application/json",
        },
        credentials: "include",
        body: JSON.stringify({
          messages: request.messages,
          tools: request.tools,
          toolChoice: request.toolChoice,
          model: request.model,
          temperature: request.temperature,
          maxTokens: request.maxTokens,
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(`Cursor LLM chat error ${response.status}: ${text}`);
      }

      return parseChatResponse(await response.json());
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
    const removeRequestAbortListener = forwardAbortSignal(request?.signal, controller);
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);

    try {
      const response = await fetch(`${this.baseUrl}/v1/chat`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          accept: "text/event-stream, application/x-ndjson",
        },
        credentials: "include",
        body: JSON.stringify({
          messages: request.messages,
          tools: request.tools,
          toolChoice: request.toolChoice,
          model: request.model,
          temperature: request.temperature,
          maxTokens: request.maxTokens,
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(`Cursor LLM streamChat error ${response.status}: ${text}`);
      }

      const reader = response.body?.getReader();
      if (!reader) {
        const full = await this.chat({ ...request, signal: controller.signal });
        yield* synthesizeStreamEvents(full);
        return;
      }

      const contentType = response.headers.get("content-type") ?? "";
      const isSSE = /text\/event-stream/i.test(contentType);

      const decoder = new TextDecoder();
      let buffer = "";
      let sawDoneSentinel = false;
      let yieldedAny = false;

      if (isSSE) {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });

          const frames = buffer.split(/\r?\n\r?\n/);
          buffer = frames.pop() ?? "";

          for (const frame of frames) {
            const lines = frame.split(/\r?\n/);
            const dataLines = [];
            for (const line of lines) {
              if (!line.startsWith("data:")) continue;
              dataLines.push(line.slice("data:".length).trimStart());
            }
            if (!dataLines.length) continue;
            const data = dataLines.join("\n").trim();
            if (!data) continue;

            if (data === "[DONE]") {
              sawDoneSentinel = true;
              yield { type: "done" };
              return;
            }

            const parsed = JSON.parse(data);
            if (isChatStreamEvent(parsed)) {
              yieldedAny = true;
              yield parsed;
              if (parsed.type === "done") return;
              continue;
            }

            if (!yieldedAny && isChatResponseLike(parsed)) {
              const full = parseChatResponse(parsed);
              yield* synthesizeStreamEvents(full);
              return;
            }

            throw new Error("Cursor LLM streamChat error: expected a ChatStreamEvent.");
          }
        }

        // Flush trailing buffer (best effort).
        const trailing = buffer.trim();
        if (trailing) {
          for (const frame of trailing.split(/\r?\n\r?\n/)) {
            const lines = frame.split(/\r?\n/);
            const dataLines = [];
            for (const line of lines) {
              if (!line.startsWith("data:")) continue;
              dataLines.push(line.slice("data:".length).trimStart());
            }
            if (!dataLines.length) continue;
            const data = dataLines.join("\n").trim();
            if (!data) continue;

            if (data === "[DONE]") {
              sawDoneSentinel = true;
              yield { type: "done" };
              return;
            }

            const parsed = JSON.parse(data);
            if (isChatStreamEvent(parsed)) {
              yieldedAny = true;
              yield parsed;
              if (parsed.type === "done") return;
              continue;
            }

            if (!yieldedAny && isChatResponseLike(parsed)) {
              const full = parseChatResponse(parsed);
              yield* synthesizeStreamEvents(full);
              return;
            }

            throw new Error("Cursor LLM streamChat error: expected a ChatStreamEvent.");
          }
        }
      } else {
        // Default to NDJSON (newline-delimited JSON).
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });

          const lines = buffer.split(/\r?\n/);
          buffer = lines.pop() ?? "";

          for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed) continue;
            const parsed = JSON.parse(trimmed);

            if (isChatStreamEvent(parsed)) {
              yieldedAny = true;
              yield parsed;
              if (parsed.type === "done") return;
              continue;
            }

            if (!yieldedAny && isChatResponseLike(parsed)) {
              const full = parseChatResponse(parsed);
              yield* synthesizeStreamEvents(full);
              return;
            }

            throw new Error("Cursor LLM streamChat error: expected a ChatStreamEvent.");
          }
        }

        const trailing = buffer.trim();
        if (trailing) {
          const parsed = JSON.parse(trailing);
          if (isChatStreamEvent(parsed)) {
            yieldedAny = true;
            yield parsed;
            if (parsed.type === "done") return;
          } else if (!yieldedAny && isChatResponseLike(parsed)) {
            const full = parseChatResponse(parsed);
            yield* synthesizeStreamEvents(full);
            return;
          } else {
            throw new Error("Cursor LLM streamChat error: expected a ChatStreamEvent.");
          }
        }
      }

      // If the backend ends the stream without an explicit done event/sentinel,
      // emit one so downstream consumers can reliably terminate their loops.
      if (!sawDoneSentinel) yield { type: "done" };
    } finally {
      clearTimeout(timeout);
      removeRequestAbortListener?.();
    }
  }
}
