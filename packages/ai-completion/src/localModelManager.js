import { LRUCache } from "./lruCache.js";

/**
 * @typedef {{
 *   model?: string,
 *   maxTokens?: number,
 *   temperature?: number,
 *   stop?: string[],
 *   timeoutMs?: number
 * }} CompletionOptions
 */

export class LocalModelManager {
  /**
   * @param {{
   *   ollamaClient: {
   *     health: () => Promise<boolean>,
   *     hasModel: (name: string) => Promise<boolean>,
   *     pullModel: (name: string) => Promise<void>,
   *     generate: (params: any) => Promise<any>
   *   },
   *   requiredModels?: string[],
   *   defaultModel?: string,
   *   cacheSize?: number
   * }} params
   */
  constructor(params) {
    if (!params || !params.ollamaClient) {
      throw new Error("LocalModelManager requires an ollamaClient");
    }
    this.ollama = params.ollamaClient;
    this.requiredModels = params.requiredModels ?? ["formula-completion"];
    this.defaultModel = params.defaultModel ?? this.requiredModels[0] ?? "formula-completion";
    this.cache = new LRUCache(params.cacheSize ?? 200);
    this.initialized = false;
  }

  async initialize() {
    const healthy = await this.ollama.health();
    if (!healthy) {
      throw new Error("Local model server not available");
    }

    for (const model of this.requiredModels) {
      const exists = await this.ollama.hasModel(model);
      if (!exists) {
        await this.ollama.pullModel(model);
      }
    }

    this.initialized = true;
  }

  /**
   * @param {string} prompt
   * @param {CompletionOptions} [options]
   * @returns {Promise<string>}
   */
  async complete(prompt, options = {}) {
    const model = options.model ?? this.defaultModel;
    const cacheKey = buildCacheKey(prompt, {
      model,
      maxTokens: options.maxTokens ?? 50,
      temperature: options.temperature ?? 0.1,
      stop: options.stop ?? [],
    });

    const cached = this.cache.get(cacheKey);
    if (cached !== undefined) return cached;

    const response = await this.ollama.generate({
      model,
      prompt,
      options: {
        temperature: options.temperature ?? 0.1,
        num_predict: options.maxTokens ?? 50,
        stop: options.stop ?? [],
      },
      stream: false,
    });

    const text = (response?.response ?? "").toString();
    this.cache.set(cacheKey, text);
    return text;
  }
}

function buildCacheKey(prompt, options) {
  return JSON.stringify({ prompt, options });
}
