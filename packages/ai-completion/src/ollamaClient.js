/**
 * Minimal Ollama HTTP client.
 *
 * We intentionally keep this small: tab-completion should work even when Ollama
 * isn't installed/running, so callers should treat failures as "LLM unavailable"
 * and fall back to rule-based suggestions.
 */

export class OllamaClient {
  /**
   * @param {{
   *   baseUrl?: string,
   *   fetchImpl?: typeof fetch,
   *   timeoutMs?: number
   * }} [options]
   */
  constructor(options = {}) {
    this.baseUrl = (options.baseUrl ?? "http://127.0.0.1:11434").replace(/\/$/, "");
    // Node 18+ provides global fetch. Allow injection for tests.
    // eslint-disable-next-line no-undef
    this.fetchImpl = options.fetchImpl ?? fetch;
    this.timeoutMs = options.timeoutMs ?? 2000;
  }

  async health() {
    try {
      const res = await this.#request("/api/tags", { method: "GET" });
      return res.ok;
    } catch {
      return false;
    }
  }

  async listModels() {
    const res = await this.#request("/api/tags", { method: "GET" });
    if (!res.ok) {
      throw new Error(`Ollama listModels failed: HTTP ${res.status}`);
    }
    const data = await res.json();
    return Array.isArray(data.models) ? data.models : [];
  }

  /**
   * @param {string} modelName
   */
  async hasModel(modelName) {
    const models = await this.listModels();
    const normalized = normalizeModelName(modelName);
    return models.some(m => normalizeModelName(m?.name) === normalized);
  }

  /**
   * @param {string} modelName
   */
  async pullModel(modelName) {
    const res = await this.#request("/api/pull", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ name: modelName }),
    }, { timeoutMs: 600000 });

    if (!res.ok) {
      throw new Error(`Ollama pullModel failed: HTTP ${res.status}`);
    }

    // Ollama streams pull progress by default. Consume the body to completion.
    // We don't currently surface progress in tab completion.
    await res.arrayBuffer();
  }

  /**
   * @param {{
   *   model: string,
   *   prompt: string,
   *   options?: Record<string, any>,
   *   stream?: boolean
   * }} params
   */
  async generate(params) {
    const res = await this.#request("/api/generate", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        model: params.model,
        prompt: params.prompt,
        options: params.options ?? {},
        stream: params.stream ?? false,
      }),
    });

    if (!res.ok) {
      throw new Error(`Ollama generate failed: HTTP ${res.status}`);
    }
    return res.json();
  }

  /**
   * @param {string} path
   * @param {RequestInit} init
   * @param {{timeoutMs?: number}} [options]
   */
  async #request(path, init, options = {}) {
    const timeoutMs = options.timeoutMs ?? this.timeoutMs;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);
    try {
      return await this.fetchImpl(`${this.baseUrl}${path}`, {
        ...init,
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timeout);
    }
  }
}

function normalizeModelName(name) {
  return String(name ?? "").trim().toLowerCase();
}
