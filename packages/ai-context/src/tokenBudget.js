import { throwIfAborted } from "./abort.js";

/**
 * Fast, model-agnostic token estimate. Works well enough for enforcing a budget
 * in a UI context manager; exact tokenization is model-specific and can be added
 * later behind the same interface.
 *
 * @param {string} text
 * @param {TokenEstimator} [estimator]
 */
export function estimateTokens(text, estimator = DEFAULT_TOKEN_ESTIMATOR) {
  return estimator.estimateTextTokens(text);
}

/**
 * @typedef {{
 *   /**
 *    * Estimate tokens for plain text.
 *    *\/
 *   estimateTextTokens: (text: string) => number,
 *   /**
 *    * Estimate tokens for a single LLM message.
 *    *\/
 *   estimateMessageTokens: (message: any) => number,
 *   /**
 *    * Estimate tokens for an array of LLM messages.
 *    *\/
 *   estimateMessagesTokens: (messages: any[]) => number
 * }} TokenEstimator
 */

/**
 * Deterministic JSON stringification with stable key ordering.
 * Useful for token estimation that shouldn't depend on object insertion order.
 *
 * @param {unknown} value
 * @returns {string}
 */
export function stableJsonStringify(value) {
  try {
    return JSON.stringify(stabilizeJson(value));
  } catch {
    try {
      // Avoid calling `String(...)` on arbitrary objects: custom `toString()` implementations
      // can leak non-heuristic sensitive strings into prompt context (even when higher-level
      // redaction is enabled).
      if (value !== null && (typeof value === "object" || typeof value === "function")) {
        return JSON.stringify("[Unserializable]");
      }
      return JSON.stringify(String(value));
    } catch {
      if (value !== null && (typeof value === "object" || typeof value === "function")) {
        return "[Unserializable]";
      }
      return String(value);
    }
  }
}

/**
 * @param {unknown} value
 * @returns {unknown}
 */
function stabilizeJson(value, stack = new WeakSet()) {
  if (value === undefined || value === null) return null;
  if (typeof value === "bigint") return value.toString();
  if (typeof value === "symbol") return value.toString();
  if (typeof value === "function") return `[Function ${value.name || "anonymous"}]`;
  if (typeof value !== "object") return value;

  // Avoid infinite recursion for cyclic structures (common in user-provided attachment payloads).
  if (stack.has(value)) return "[Circular]";
  stack.add(value);

  try {
    if (value instanceof Date) {
      // Avoid calling per-instance overrides (e.g. `date.toISOString = () => "secret"`), which can
      // leak sensitive strings into prompt context even when higher-level DLP redaction is enabled.
      try {
        return Date.prototype.toISOString.call(value);
      } catch {
        // Invalid dates throw in `toISOString()`; fall back to a stable, non-throwing string form.
        try {
          return Date.prototype.toString.call(value);
        } catch {
          return "";
        }
      }
    }

    if (value instanceof Map) {
      const entries = Array.from(value.entries()).map(([k, v], index) => {
        const stableKey = stabilizeJson(k, stack);
        const stableValue = stabilizeJson(v, stack);
        // Use a stable string form for sorting that is independent of insertion order.
        const keySort = JSON.stringify(stableKey) ?? "";
        const valueSort = JSON.stringify(stableValue) ?? "";
        return { stableKey, stableValue, keySort, valueSort, index };
      });

      entries.sort((a, b) => {
        if (a.keySort < b.keySort) return -1;
        if (a.keySort > b.keySort) return 1;
        if (a.valueSort < b.valueSort) return -1;
        if (a.valueSort > b.valueSort) return 1;
        return a.index - b.index;
      });

      return entries.map((e) => [e.stableKey, e.stableValue]);
    }

    if (value instanceof Set) {
      const values = Array.from(value.values()).map((v, index) => {
        const stableValue = stabilizeJson(v, stack);
        const valueSort = JSON.stringify(stableValue) ?? "";
        return { stableValue, valueSort, index };
      });

      values.sort((a, b) => {
        if (a.valueSort < b.valueSort) return -1;
        if (a.valueSort > b.valueSort) return 1;
        return a.index - b.index;
      });

      return values.map((v) => v.stableValue);
    }

    if (Array.isArray(value)) return value.map((v) => stabilizeJson(v, stack));

    const obj = /** @type {Record<string, unknown>} */ (value);
    const keys = Object.keys(obj).sort();
    /** @type {Record<string, unknown>} */
    const out = {};
    for (const key of keys) out[key] = stabilizeJson(obj[key], stack);
    return out;
  } finally {
    stack.delete(value);
  }
}

/**
 * Create a fast, heuristic TokenEstimator.
 *
 * The default implementation assumes ~4 chars/token for English-like text.
 * Message arrays are estimated by counting content + a small per-message overhead,
 * plus JSON-stringified tool call payloads.
 *
 * @param {{
 *   charsPerToken?: number,
 *   tokensPerMessageOverhead?: number
 * }} [options]
 * @returns {TokenEstimator}
 */
export function createHeuristicTokenEstimator(options = {}) {
  const charsPerToken = options.charsPerToken ?? 4;
  const tokensPerMessageOverhead = options.tokensPerMessageOverhead ?? 4;

  /**
   * @param {string} text
   */
  function estimateTextTokens(text) {
    if (!text) return 0;
    return Math.ceil(text.length / charsPerToken);
  }

  /**
   * @param {any} message
   */
  function estimateMessageTokens(message) {
    if (!message || typeof message !== "object") return 0;

    const role = typeof message.role === "string" ? message.role : "";
    const content = typeof message.content === "string" ? message.content : "";
    /** @type {string[]} */
    const parts = [role, content];

    if (message.role === "tool") {
      const toolCallId = typeof message.toolCallId === "string" ? message.toolCallId : "";
      parts.push(toolCallId);
    }

    if (message.role === "assistant" && Array.isArray(message.toolCalls) && message.toolCalls.length) {
      // Tool calls are structured, but we conservatively count the JSON payload size.
      parts.push(stableJsonStringify(message.toolCalls));
    }

    return estimateTextTokens(parts.join("\n")) + tokensPerMessageOverhead;
  }

  /**
   * @param {any[]} messages
   */
  function estimateMessagesTokens(messages) {
    if (!Array.isArray(messages) || messages.length === 0) return 0;
    let total = 0;
    for (const msg of messages) total += estimateMessageTokens(msg);
    return total;
  }

  return {
    estimateTextTokens,
    estimateMessageTokens,
    estimateMessagesTokens
  };
}

/** @type {TokenEstimator} */
export const DEFAULT_TOKEN_ESTIMATOR = createHeuristicTokenEstimator();

/**
 * Estimate tokens for an array of LLM messages.
 *
 * @param {any[]} messages
 * @param {TokenEstimator} [estimator]
 */
export function estimateMessagesTokens(messages, estimator = DEFAULT_TOKEN_ESTIMATOR) {
  return estimator.estimateMessagesTokens(messages);
}

/**
 * Estimate tokens for tool definitions passed via `ChatRequest.tools`.
 *
 * This is necessarily heuristic (providers tokenize tool schemas differently),
 * but helps prevent runaway prompts due to large JSON schemas.
 *
 * @param {any[] | null | undefined} tools
 * @param {TokenEstimator} [estimator]
 */
export function estimateToolDefinitionTokens(tools, estimator = DEFAULT_TOKEN_ESTIMATOR) {
  if (!Array.isArray(tools) || tools.length === 0) return 0;
  return estimator.estimateTextTokens(stableJsonStringify(tools));
}

/**
 * Trim a string to a maximum estimated token count.
 * @param {string} text
 * @param {number} maxTokens
 * @param {TokenEstimator} [estimator]
 * @param {{ signal?: AbortSignal }} [options]
 */
export function trimToTokenBudget(text, maxTokens, estimator = DEFAULT_TOKEN_ESTIMATOR, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  if (maxTokens <= 0) return "";
  const estimate = estimateTokens(text, estimator);
  if (estimate <= maxTokens) return text;

  const suffix = "\n…(trimmed to fit token budget)…";
  const suffixTokens = estimateTokens(suffix, estimator);

  // If the suffix itself doesn't fit, return as much of it as possible.
  if (suffixTokens >= maxTokens) {
    let lo = 0;
    let hi = suffix.length;
    while (lo < hi) {
      throwIfAborted(signal);
      const mid = Math.ceil((lo + hi) / 2);
      const candidate = suffix.slice(0, mid);
      const tokens = estimateTokens(candidate, estimator);
      if (tokens <= maxTokens) lo = mid;
      else hi = mid - 1;
    }
    return suffix.slice(0, lo);
  }

  // Find the longest prefix that fits when we append the suffix.
  let lo = 0;
  let hi = text.length;
  while (lo < hi) {
    throwIfAborted(signal);
    const mid = Math.ceil((lo + hi) / 2);
    const candidate = text.slice(0, mid) + suffix;
    const tokens = estimateTokens(candidate, estimator);
    if (tokens <= maxTokens) lo = mid;
    else hi = mid - 1;
  }

  return text.slice(0, lo) + suffix;
}

/**
 * @param {{ key: string, text: string, priority: number }[]} sections
 * @param {number} maxTokens
 * @param {TokenEstimator} [estimator]
 * @param {{ signal?: AbortSignal }} [options]
 */
export function packSectionsToTokenBudget(sections, maxTokens, estimator = DEFAULT_TOKEN_ESTIMATOR, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const ordered = orderSectionsDeterministically(sections);
  let remaining = Math.max(0, maxTokens);
  /** @type {{ key: string, text: string }[]} */
  const packed = [];

  for (const section of ordered) {
    throwIfAborted(signal);
    if (remaining <= 0) break;
    const trimmed = trimToTokenBudget(section.text, remaining, estimator, { signal });
    const used = estimateTokens(trimmed, estimator);
    remaining -= used;
    packed.push({ key: section.key, text: trimmed });
  }

  throwIfAborted(signal);
  return packed;
}

/**
 * Pack sections while returning a detailed report about token usage.
 *
 * This is useful for debugging prompt/context construction without depending on
 * model-specific tokenizers.
 *
 * @param {{ key: string, text: string, priority: number }[]} sections
 * @param {number} maxTokens
 * @param {TokenEstimator} [estimator]
 * @param {{ signal?: AbortSignal }} [options]
 * @returns {{
 *   packed: { key: string, text: string }[],
 *   report: {
 *     maxTokens: number,
 *     remainingTokens: number,
 *     sections: Array<{
 *       key: string,
 *       priority: number,
 *       tokensPreTrim: number,
 *       tokensPostTrim: number,
 *       trimmed: boolean,
 *       dropped: boolean
 *     }>
 *   }
 * }}
 */
export function packSectionsToTokenBudgetWithReport(sections, maxTokens, estimator = DEFAULT_TOKEN_ESTIMATOR, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);

  const ordered = orderSectionsDeterministically(sections);
  let remaining = Math.max(0, maxTokens);

  /** @type {{ key: string, text: string }[]} */
  const packed = [];
  /** @type {Array<{ key: string, priority: number, tokensPreTrim: number, tokensPostTrim: number, trimmed: boolean, dropped: boolean }>} */
  const reportSections = [];

  for (const section of ordered) {
    throwIfAborted(signal);

    const tokensPreTrim = estimateTokens(section.text, estimator);

    if (remaining <= 0) {
      reportSections.push({
        key: section.key,
        priority: section.priority,
        tokensPreTrim,
        tokensPostTrim: 0,
        trimmed: false,
        dropped: true
      });
      continue;
    }

    const trimmedText = trimToTokenBudget(section.text, remaining, estimator, { signal });
    const tokensPostTrim = estimateTokens(trimmedText, estimator);
    remaining -= tokensPostTrim;

    packed.push({ key: section.key, text: trimmedText });
    reportSections.push({
      key: section.key,
      priority: section.priority,
      tokensPreTrim,
      tokensPostTrim,
      trimmed: trimmedText !== section.text,
      dropped: false
    });
  }

  throwIfAborted(signal);
  return {
    packed,
    report: {
      maxTokens,
      remainingTokens: remaining,
      sections: reportSections
    }
  };
}

/**
 * @param {{ key: string, text: string, priority: number }[]} sections
 */
function orderSectionsDeterministically(sections) {
  return sections
    .map((section, originalIndex) => ({ section, originalIndex }))
    .sort((a, b) => {
      const aPriority = Number.isFinite(a.section.priority) ? a.section.priority : 0;
      const bPriority = Number.isFinite(b.section.priority) ? b.section.priority : 0;
      const priorityDiff = bPriority - aPriority;
      if (priorityDiff !== 0) return priorityDiff;
      // Deterministic tie-breaker: preserve original input order.
      return a.originalIndex - b.originalIndex;
    })
    .map(({ section }) => section);
}
