/**
 * Minimal OpenAI embedding wrapper.
 *
 * Reference:
 *   POST https://api.openai.com/v1/embeddings
 *   { model, input: string[] }
 */
export class OpenAIEmbedder {
  /**
   * @param {{ apiKey: string, model: string, baseUrl?: string }} opts
   */
  constructor(opts) {
    if (!opts?.apiKey) throw new Error("OpenAIEmbedder requires apiKey");
    if (!opts?.model) throw new Error("OpenAIEmbedder requires model");
    this._apiKey = opts.apiKey;
    this._model = opts.model;
    this._baseUrl = opts.baseUrl ?? "https://api.openai.com";
    this._dimension = null;
  }

  get name() {
    return `openai:${this._model}`;
  }

  get dimension() {
    return this._dimension;
  }

  /**
   * @param {string[]} texts
   * @returns {Promise<number[][]>}
   */
  async embedTexts(texts) {
    const res = await fetch(`${this._baseUrl}/v1/embeddings`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        authorization: `Bearer ${this._apiKey}`,
      },
      body: JSON.stringify({ model: this._model, input: texts }),
    });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`OpenAI embeddings failed (${res.status}): ${body}`);
    }
    const json = await res.json();
    if (!json || !Array.isArray(json.data)) {
      throw new Error("OpenAI embeddings response missing 'data'");
    }
    const vectors = json.data
      .slice()
      .sort((a, b) => a.index - b.index)
      .map((d) => d.embedding);
    if (this._dimension == null && vectors.length > 0) this._dimension = vectors[0].length;
    return vectors;
  }
}
