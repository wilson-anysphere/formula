import { DEFAULT_TOKEN_ESTIMATOR } from "./tokenBudget.js";
import { awaitWithAbort, throwIfAborted } from "./abort.js";

export const CONTEXT_SUMMARY_MARKER = "[CONTEXT_SUMMARY]";
const DEFAULT_SUMMARY_MAX_TOKENS = 256;
const DEFAULT_KEEP_LAST_MESSAGES = 40;
const TRIM_SUFFIX = "\n…(trimmed to fit context budget)…";

/**
 * @param {any} message
 */
function isGeneratedSummaryMessage(message) {
  return (
    message &&
    typeof message === "object" &&
    message.role === "system" &&
    typeof message.content === "string" &&
    message.content.startsWith(CONTEXT_SUMMARY_MARKER)
  );
}

/**
 * @param {any} message
 * @param {number} maxTokens
 * @param {import("./tokenBudget.js").TokenEstimator} estimator
 * @param {AbortSignal | undefined} signal
 */
function trimMessageToTokens(message, maxTokens, estimator, signal) {
  throwIfAborted(signal);
  if (!message || typeof message !== "object") return message;
  if (maxTokens <= 0) return { ...message, content: "" };
  if (estimator.estimateMessageTokens(message) <= maxTokens) return message;

  const content = typeof message.content === "string" ? message.content : "";
  if (!content) return { ...message, content: "" };

  // Binary search a prefix length that fits, including the trim suffix.
  let lo = 0;
  let hi = content.length;
  let best = "";

  while (lo <= hi) {
    throwIfAborted(signal);
    const mid = Math.floor((lo + hi) / 2);
    const candidate = content.slice(0, mid) + TRIM_SUFFIX;
    const next = { ...message, content: candidate };
    if (estimator.estimateMessageTokens(next) <= maxTokens) {
      best = candidate;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }

  if (!best) {
    // Even the suffix doesn't fit; fall back to the smallest possible content.
    return { ...message, content: "" };
  }

  return { ...message, content: best };
}

/**
 * Deterministic stub summary (no model call). Intended as a safe default until a
 * real summarization model hook is provided by the caller.
 *
 * @param {any[]} messages
 * @param {AbortSignal | undefined} signal
 * @returns {string}
 */
function defaultSummarize(messages, signal) {
  const lines = [];
  for (const msg of messages) {
    throwIfAborted(signal);
    if (!msg || typeof msg !== "object") continue;
    const role = typeof msg.role === "string" ? msg.role : "unknown";
    const content = typeof msg.content === "string" ? msg.content : "";
    const normalized = content.replace(/\s+/g, " ").trim();
    const snippet = normalized.length > 120 ? `${normalized.slice(0, 120)}…` : normalized;
    const toolCalls =
      msg.role === "assistant" && Array.isArray(msg.toolCalls) && msg.toolCalls.length
        ? ` tool_calls=[${msg.toolCalls.map((c) => c?.name).filter(Boolean).join(", ")}]`
        : "";
    lines.push(`- ${role}${toolCalls}: ${snippet || "(empty)"}`);
  }

  const count = messages.length;
  return [`Summary of earlier conversation (${count} messages):`, ...lines].join("\n");
}

/**
 * Take a suffix of `messages` that fits within `budgetTokens`. Always includes
 * the most recent message (trimming it if necessary).
 *
 * @param {any[]} messages
 * @param {number} budgetTokens
 * @param {import("./tokenBudget.js").TokenEstimator} estimator
 * @param {AbortSignal | undefined} signal
 * @returns {{ kept: any[], dropped: any[] }}
 */
function takeTailWithinBudget(messages, budgetTokens, estimator, signal) {
  throwIfAborted(signal);
  if (!Array.isArray(messages) || messages.length === 0) return { kept: [], dropped: [] };

  const last = messages[messages.length - 1];
  if (budgetTokens <= 0) {
    return { kept: [], dropped: messages.slice() };
  }

  let remaining = budgetTokens;
  /** @type {any[]} */
  const keptReversed = [];

  for (let i = messages.length - 1; i >= 0; i -= 1) {
    throwIfAborted(signal);
    const msg = messages[i];
    const tokens = estimator.estimateMessageTokens(msg);

    if (tokens <= remaining) {
      keptReversed.push(msg);
      remaining -= tokens;
      continue;
    }

    if (keptReversed.length === 0) {
      // Always keep the most recent message, but trim its content to fit.
      keptReversed.push(trimMessageToTokens(msg, remaining, estimator, signal));
      remaining = 0;
    }
    break;
  }

  throwIfAborted(signal);
  const kept = keptReversed.reverse();
  const droppedCount = Math.max(0, messages.length - kept.length);
  const dropped = droppedCount > 0 ? messages.slice(0, droppedCount) : [];
  return { kept, dropped };
}

/**
 * Trim an LLM message list to fit within a token budget.
 *
 * Policy:
 * - Preserve all system messages (except generated summary messages; those can be replaced).
 * - Preserve the most recent N user/assistant/tool messages.
 * - If older history must be dropped, replace it with a single summary message produced by
 *   a pluggable callback `summarize(messagesToSummarize)`.
 *
 * NOTE: This utility is intentionally heuristic; it is meant to prevent crashes due to
 * exceeding provider context limits, not to be a perfect tokenizer.
 *
 * @param {{
 *   messages: any[],
 *   maxTokens: number,
 *   reserveForOutputTokens?: number,
 *   estimator?: import("./tokenBudget.js").TokenEstimator,
 *   keepLastMessages?: number,
 *   summaryMaxTokens?: number,
 *   summarize?: ((messagesToSummarize: any[]) => (string | any | null | undefined) | Promise<string | any | null | undefined>) | null,
 *   summaryRole?: "system" | "assistant"
 *   signal?: AbortSignal
 * }} params
 * @returns {Promise<any[]>}
 */
export async function trimMessagesToBudget(params) {
  const signal = params.signal;
  throwIfAborted(signal);
  const messages = Array.isArray(params.messages) ? params.messages : [];
  const maxTokens = Number.isFinite(params.maxTokens) ? params.maxTokens : 0;
  const reserveForOutputTokens = Number.isFinite(params.reserveForOutputTokens) ? params.reserveForOutputTokens : 0;
  const estimator = params.estimator ?? DEFAULT_TOKEN_ESTIMATOR;
  const keepLastMessages = Number.isFinite(params.keepLastMessages) ? params.keepLastMessages : DEFAULT_KEEP_LAST_MESSAGES;
  const summaryMaxTokens = Number.isFinite(params.summaryMaxTokens) ? params.summaryMaxTokens : DEFAULT_SUMMARY_MAX_TOKENS;
  const summaryRole = params.summaryRole ?? "system";
  const summarize =
    params.summarize === null ? null : params.summarize ?? ((messagesToSummarize) => defaultSummarize(messagesToSummarize, signal));

  const allowedPromptTokens = Math.max(0, maxTokens - reserveForOutputTokens);
  if (allowedPromptTokens <= 0) return [];

  if (estimator.estimateMessagesTokens(messages) <= allowedPromptTokens) return messages;

  /** @type {any[]} */
  const systemMessages = [];
  /** @type {any[]} */
  const otherMessages = [];

  for (const msg of messages) {
    throwIfAborted(signal);
    if (msg && typeof msg === "object" && msg.role === "system" && !isGeneratedSummaryMessage(msg)) {
      systemMessages.push(msg);
    } else {
      otherMessages.push(msg);
    }
  }

  // Trim system messages only if they exceed the entire prompt budget.
  let systemTrimmed = systemMessages.map((m) => ({ ...m }));
  if (estimator.estimateMessagesTokens(systemTrimmed) > allowedPromptTokens) {
    for (let i = systemTrimmed.length - 1; i >= 0; i -= 1) {
      throwIfAborted(signal);
      const without = systemTrimmed.filter((_m, idx) => idx !== i);
      const tokensWithout = estimator.estimateMessagesTokens(without);
      const budgetForThis = Math.max(0, allowedPromptTokens - tokensWithout);
      systemTrimmed[i] = trimMessageToTokens(systemTrimmed[i], budgetForThis, estimator, signal);
    }
  }

  const systemTokens = estimator.estimateMessagesTokens(systemTrimmed);
  const otherBudget = Math.max(0, allowedPromptTokens - systemTokens);

  if (estimator.estimateMessagesTokens(otherMessages) <= otherBudget) {
    return [...systemTrimmed, ...otherMessages];
  }

  const generatedSummaries = otherMessages.filter(isGeneratedSummaryMessage);
  const nonSummaryOtherMessages = otherMessages.filter((m) => !isGeneratedSummaryMessage(m));

  // Enforce a count cap on non-system messages (prevents unbounded growth even when
  // many short messages would still fit under the token budget).
  const keepStart = Math.max(0, nonSummaryOtherMessages.length - Math.max(0, keepLastMessages));
  const olderByCount = nonSummaryOtherMessages.slice(0, keepStart);
  const recentByCount = nonSummaryOtherMessages.slice(keepStart);

  let summaryBudget = summarize ? Math.min(summaryMaxTokens, otherBudget) : 0;
  let recentBudget = Math.max(0, otherBudget - summaryBudget);
  // Never steal *all* budget from the most recent messages just to include a summary.
  // If the budget is too tight, prefer keeping the latest user/tool context.
  if (recentBudget <= 0) {
    summaryBudget = 0;
    recentBudget = otherBudget;
  }

  const { kept: recentKept, dropped: droppedFromRecent } = takeTailWithinBudget(recentByCount, recentBudget, estimator, signal);
  const toSummarize = [...generatedSummaries, ...olderByCount, ...droppedFromRecent];

  /** @type {any[]} */
  const out = [...systemTrimmed];

  if (summarize && toSummarize.length > 0 && summaryBudget > 0) {
    throwIfAborted(signal);
    const summaryValue = await awaitWithAbort(summarize(toSummarize), signal);
    throwIfAborted(signal);
    if (summaryValue) {
      const summaryMessage =
        typeof summaryValue === "string"
          ? { role: summaryRole, content: `${CONTEXT_SUMMARY_MARKER}\n${summaryValue}`.trim() }
          : {
              role: summaryValue.role ?? summaryRole,
              content: `${CONTEXT_SUMMARY_MARKER}\n${String(summaryValue.content ?? "")}`.trim(),
            };
      out.push(trimMessageToTokens(summaryMessage, summaryBudget, estimator, signal));
    }
  }

  out.push(...recentKept);

  // Final safety check: if we still exceed, drop/truncate from the start of non-system messages.
  while (out.length > 0 && estimator.estimateMessagesTokens(out) > allowedPromptTokens) {
    throwIfAborted(signal);
    // Preserve all system messages; remove the first non-system if possible.
    const firstNonSystemIdx = out.findIndex((m) => !(m && typeof m === "object" && m.role === "system" && !isGeneratedSummaryMessage(m)));
    if (firstNonSystemIdx === -1) {
      // Only system messages remain; truncate the last one.
      const idx = out.length - 1;
      out[idx] = trimMessageToTokens(
        out[idx],
        Math.max(0, allowedPromptTokens - estimator.estimateMessagesTokens(out.slice(0, idx))),
        estimator,
        signal
      );
      break;
    }
    out.splice(firstNonSystemIdx, 1);
  }

  // Ensure we never exceed the budget due to unexpected estimator behavior.
  if (estimator.estimateMessagesTokens(out) > allowedPromptTokens) {
    return takeTailWithinBudget(out, allowedPromptTokens, estimator, signal).kept;
  }

  throwIfAborted(signal);
  return out;
}
