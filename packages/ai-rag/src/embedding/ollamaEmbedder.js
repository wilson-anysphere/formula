/**
 * Minimal Ollama embedding wrapper.
 *
 * Ollama endpoint reference:
 *   POST /api/embeddings { model, prompt }
 */
export class OllamaEmbedder {
  /**
   * @param {{ model: string, host?: string, dimension?: number }} opts
   */
  constructor(opts) {
    if (!opts?.model) throw new Error("OllamaEmbedder requires model");
    this._model = opts.model;
    this._host = opts.host ?? "http://localhost:11434";
    // Dimension depends on model; callers can provide for stores that require it.
    this._dimension = opts.dimension ?? null;
  }

  get name() {
    return `ollama:${this._model}`;
  }

  get dimension() {
    return this._dimension;
  }

  /**
   * @param {string} text
   * @returns {Promise<number[]>}
   */
  async _embedOne(text) {
    const res = await fetch(`${this._host}/api/embeddings`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ model: this._model, prompt: text }),
    });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`Ollama embeddings failed (${res.status}): ${body}`);
    }
    const json = await res.json();
    if (!json || !Array.isArray(json.embedding)) {
      throw new Error("Ollama embeddings response missing 'embedding'");
    }
    return json.embedding;
  }

  /**
   * @param {string[]} texts
   * @returns {Promise<number[][]>}
   */
  async embedTexts(texts) {
    // Ollama does not support batch embeddings in the standard endpoint; do sequential.
    const out = [];
    for (const text of texts) out.push(await this._embedOne(text));
    if (this._dimension == null && out.length > 0) this._dimension = out[0].length;
    return out;
  }
}
