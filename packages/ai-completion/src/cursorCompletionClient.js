/**
 * Cursor backend completion client.
 *
 * This is intentionally minimal and dependency-free so it can be used in both
 * browser (Vite) and Node/Electron contexts.
 *
 * The Cursor-managed backend is responsible for selecting models/auth; this
 * client only forwards a prompt and returns the completion text.
 */
export class CursorCompletionClient {
  /**
   * @param {{
   *   baseUrl: string,
   *   fetchImpl?: typeof fetch,
   *   timeoutMs?: number,
   *   headers?: Record<string, string>
   * }} options
   */
  constructor(options) {
    if (!options || !options.baseUrl) {
      throw new Error("CursorCompletionClient requires a baseUrl");
    }
    this.baseUrl = String(options.baseUrl).trim();
    if (!this.baseUrl) {
      throw new Error("CursorCompletionClient requires a non-empty baseUrl");
    }
    // Node 18+ provides global fetch. Allow injection for tests.
    // eslint-disable-next-line no-undef
    this.fetchImpl = options.fetchImpl ?? fetch;
    this.timeoutMs = options.timeoutMs ?? 200;
    this.headers = options.headers ?? {};
  }

  /**
   * @param {string} prompt
   * @param {{ model?: string, maxTokens?: number, temperature?: number, stop?: string[], timeoutMs?: number }} [options]
   */
  async complete(prompt, options = {}) {
    const controller = new AbortController();
    const timeoutMs = Number.isFinite(options.timeoutMs) ? options.timeoutMs : this.timeoutMs;
    const timeout = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const res = await this.fetchImpl(this.baseUrl, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          ...this.headers,
        },
        body: JSON.stringify({
          prompt: (prompt ?? "").toString(),
          model: options.model,
          maxTokens: options.maxTokens,
          temperature: options.temperature,
          stop: options.stop,
        }),
        signal: controller.signal,
      });

      if (!res.ok) {
        throw new Error(`CursorCompletionClient failed: HTTP ${res.status}`);
      }

      const contentType = res.headers.get("content-type") ?? "";
      if (contentType.includes("application/json")) {
        const data = await res.json();
        return extractCompletionText(data);
      }

      const text = await res.text();
      return String(text ?? "");
    } finally {
      clearTimeout(timeout);
    }
  }
}

function extractCompletionText(data) {
  if (typeof data === "string") return data;
  if (!data || typeof data !== "object") return "";

  // Common shapes.
  if (typeof data.completion === "string") return data.completion;
  if (typeof data.text === "string") return data.text;

  // OpenAI-ish shapes.
  if (Array.isArray(data.choices) && data.choices.length > 0) {
    const choice = data.choices[0];
    if (typeof choice?.text === "string") return choice.text;
    const content = choice?.message?.content;
    if (typeof content === "string") return content;
  }

  // Fallback.
  return "";
}

