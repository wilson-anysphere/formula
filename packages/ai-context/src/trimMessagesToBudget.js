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
    (message.role === "system" || message.role === "assistant") &&
    typeof message.content === "string" &&
    message.content.startsWith(CONTEXT_SUMMARY_MARKER)
  );
}

/**
 * @param {any} message
 */
function isAssistantToolCallMessage(message) {
  return (
    message &&
    typeof message === "object" &&
    message.role === "assistant" &&
    Array.isArray(message.toolCalls) &&
    message.toolCalls.length > 0
  );
}

/**
 * @param {any} message
 */
function isToolMessage(message) {
  return message && typeof message === "object" && message.role === "tool";
}

/**
 * Find the bounds of a tool-call group containing the message at `idx`.
 *
 * A group is:
 * - An assistant message with `toolCalls`, followed by 0+ contiguous tool messages.
 *
 * Tool messages are assumed to immediately follow their originating assistant tool call.
 *
 * @param {any[]} messages
 * @param {number} idx
 * @returns {{ start: number, end: number } | null}
 */
function getToolCallGroupBounds(messages, idx) {
  if (!Array.isArray(messages) || messages.length === 0) return null;
  if (!Number.isInteger(idx) || idx < 0 || idx >= messages.length) return null;

  const msg = messages[idx];
  if (isAssistantToolCallMessage(msg)) {
    let end = idx;
    for (let i = idx + 1; i < messages.length; i += 1) {
      const next = messages[i];
      if (!isToolMessage(next)) break;
      end = i;
    }
    return { start: idx, end };
  }

  if (isToolMessage(msg)) {
    // Rewind to the start of the contiguous tool block.
    let j = idx;
    while (j - 1 >= 0 && isToolMessage(messages[j - 1])) j -= 1;
    const assistantIdx = j - 1;
    if (assistantIdx >= 0 && isAssistantToolCallMessage(messages[assistantIdx])) {
      let end = j;
      for (let i = j + 1; i < messages.length; i += 1) {
        const next = messages[i];
        if (!isToolMessage(next)) break;
        end = i;
      }
      return { start: assistantIdx, end };
    }
  }

  return null;
}

/**
 * If `startIdx` points into the middle of a tool-call group (i.e. at a tool message),
 * rewind to the group start so we don't keep orphaned tool messages.
 *
 * @param {any[]} messages
 * @param {number} startIdx
 */
function rewindStartIdxForToolCallPairs(messages, startIdx) {
  if (!Array.isArray(messages) || messages.length === 0) return startIdx;
  if (!Number.isInteger(startIdx) || startIdx <= 0 || startIdx >= messages.length) return startIdx;

  const bounds = getToolCallGroupBounds(messages, startIdx);
  if (!bounds) return startIdx;
  // `startIdx` is in a tool message block; move it to the start of the group.
  if (bounds.start < startIdx) return bounds.start;
  return startIdx;
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
 * @param {{ preserveToolCallPairs?: boolean } | undefined} [options]
 * @returns {{ kept: any[], dropped: any[] }}
 */
function takeTailWithinBudget(messages, budgetTokens, estimator, signal, options) {
  throwIfAborted(signal);
  if (!Array.isArray(messages) || messages.length === 0) return { kept: [], dropped: [] };

  if (budgetTokens <= 0) {
    return { kept: [], dropped: messages.slice() };
  }

  const preserveToolCallPairs = options?.preserveToolCallPairs !== false;

  if (!preserveToolCallPairs) {
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

  let remaining = budgetTokens;
  /** @type {any[]} */
  const keptSegmentsReversed = [];

  let keptStartIdx = messages.length;

  // Walk backwards, but treat assistant(toolCalls)+tool* as an atomic group.
  for (let i = messages.length - 1; i >= 0; i -= 1) {
    throwIfAborted(signal);
    const group = getToolCallGroupBounds(messages, i);
    const groupStart = group?.start ?? i;
    const groupEnd = group?.end ?? i;
    const slice = messages.slice(groupStart, groupEnd + 1);
    const tokens = estimator.estimateMessagesTokens(slice);

    if (tokens <= remaining) {
      keptSegmentsReversed.push(slice);
      keptStartIdx = Math.min(keptStartIdx, groupStart);
      remaining -= tokens;
      // Skip over the group we just consumed.
      i = groupStart;
      continue;
    }

    if (keptSegmentsReversed.length === 0) {
      // Always keep something. For tool-call groups, we try to keep the full group
      // by aggressively trimming message *contents* (but not tool call metadata).
      if (group) {
        // Compute the minimum token cost of keeping the group structure with empty content.
        const emptyContentGroup = slice.map((m) => (m && typeof m === "object" ? { ...m, content: "" } : m));
        const minTokens = estimator.estimateMessagesTokens(emptyContentGroup);
        if (minTokens <= remaining) {
          const fullTokensPerMsg = slice.map((m) => estimator.estimateMessageTokens(m));
          const minTokensPerMsg = emptyContentGroup.map((m) => estimator.estimateMessageTokens(m));
          let extra = remaining - minTokens;
          /** @type {number[]} */
          const extraAlloc = new Array(slice.length).fill(0);
          for (let j = slice.length - 1; j >= 0; j -= 1) {
            throwIfAborted(signal);
            const delta = Math.max(0, fullTokensPerMsg[j] - minTokensPerMsg[j]);
            const take = Math.min(delta, extra);
            extraAlloc[j] = take;
            extra -= take;
          }
          const trimmedGroup = slice.map((m, idx) =>
            trimMessageToTokens(m, minTokensPerMsg[idx] + extraAlloc[idx], estimator, signal)
          );
          keptSegmentsReversed.push(trimmedGroup);
          keptStartIdx = groupStart;
          remaining = 0;
        }
      } else {
        keptSegmentsReversed.push([trimMessageToTokens(messages[i], remaining, estimator, signal)]);
        keptStartIdx = i;
        remaining = 0;
      }
    }
    break;
  }

  throwIfAborted(signal);
  const kept = keptSegmentsReversed
    .reverse()
    .flatMap((segment) => segment);
  const dropped = keptStartIdx > 0 ? messages.slice(0, keptStartIdx) : [];
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
 *   /**
 *    * When true (default), preserve tool-call coherence:
 *    * - Never keep a `role: "tool"` message without also keeping the assistant message that
 *    *   issued the matching tool call.
 *    * - Never keep an assistant message with `toolCalls` without also keeping the subsequent
 *    *   `role: "tool"` messages (when present).
 *    *\/
 *   preserveToolCallPairs?: boolean,
 *   /**
 *    * When true, prefer dropping completed tool-call groups (assistant(toolCalls)+tool*)
 *    * before dropping other messages. This can help retain more conversational context when
 *    * tool outputs are large.
 *    *\/
 *   dropToolMessagesFirst?: boolean,
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
  const preserveToolCallPairs = params.preserveToolCallPairs !== false;
  const dropToolMessagesFirst = params.dropToolMessagesFirst === true;
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
  let nonSummaryOtherMessages = otherMessages.filter((m) => !isGeneratedSummaryMessage(m));

  /** @type {any[]} */
  const orphanToolMessages = [];
  if (preserveToolCallPairs && nonSummaryOtherMessages.length > 0) {
    /** @type {any[]} */
    const filtered = [];
    let inToolGroup = false;
    for (const msg of nonSummaryOtherMessages) {
      throwIfAborted(signal);
      if (isAssistantToolCallMessage(msg)) {
        filtered.push(msg);
        inToolGroup = true;
        continue;
      }
      if (isToolMessage(msg)) {
        if (inToolGroup) {
          filtered.push(msg);
        } else {
          // If the input already contains orphan tool messages, avoid preserving them when trimming.
          orphanToolMessages.push(msg);
        }
        continue;
      }
      filtered.push(msg);
      inToolGroup = false;
    }
    nonSummaryOtherMessages = filtered;
  }

  // Enforce a count cap on non-system messages (prevents unbounded growth even when
  // many short messages would still fit under the token budget).
  let keepStart = Math.max(0, nonSummaryOtherMessages.length - Math.max(0, keepLastMessages));
  if (preserveToolCallPairs) keepStart = rewindStartIdxForToolCallPairs(nonSummaryOtherMessages, keepStart);
  const olderByCount = nonSummaryOtherMessages.slice(0, keepStart);
  const recentByCountRaw = nonSummaryOtherMessages.slice(keepStart);

  let summaryBudget = summarize ? Math.min(summaryMaxTokens, otherBudget) : 0;
  let recentBudget = Math.max(0, otherBudget - summaryBudget);
  // Never steal *all* budget from the most recent messages just to include a summary.
  // If the budget is too tight, prefer keeping the latest user/tool context.
  if (recentBudget <= 0) {
    summaryBudget = 0;
    recentBudget = otherBudget;
  }

  /** @type {any[]} */
  let droppedToolGroups = [];
  let recentByCount = recentByCountRaw;

  if (preserveToolCallPairs && dropToolMessagesFirst && recentByCount.length > 0 && recentBudget > 0) {
    // Prefer dropping completed tool-call groups (assistant(toolCalls)+tool*) before other messages.
    // Only drop groups that are followed by at least one non-tool message, so we don't
    // remove the tool output the model is about to respond to.
    let recentTokens = estimator.estimateMessagesTokens(recentByCount);
    while (recentTokens > recentBudget) {
      throwIfAborted(signal);
      let removed = false;
      for (let i = 0; i < recentByCount.length; i += 1) {
        throwIfAborted(signal);
        if (!isAssistantToolCallMessage(recentByCount[i])) continue;
        const group = getToolCallGroupBounds(recentByCount, i);
        if (!group) continue;
        // Skip groups that reach the end (likely in-progress tool call).
        if (group.end >= recentByCount.length - 1) continue;
        const groupMessages = recentByCount.slice(group.start, group.end + 1);
        droppedToolGroups.push(...groupMessages);
        recentByCount = [...recentByCount.slice(0, group.start), ...recentByCount.slice(group.end + 1)];
        recentTokens = estimator.estimateMessagesTokens(recentByCount);
        removed = true;
        break;
      }
      if (!removed) break;
    }
  }

  const { kept: recentKept, dropped: droppedFromRecent } = takeTailWithinBudget(recentByCount, recentBudget, estimator, signal, {
    preserveToolCallPairs,
  });
  const toSummarize = [...generatedSummaries, ...orphanToolMessages, ...olderByCount, ...droppedToolGroups, ...droppedFromRecent];

  /** @type {any[]} */
  const out = [...systemTrimmed];

  // If we couldn't keep any recent messages (e.g. the most recent coherent tool-call group
  // is too large to fit), fall back to using the summary budget even if we would otherwise
  // prefer keeping recent messages over a summary. This avoids returning an empty/non-actionable
  // context under tight budgets.
  if (summarize && toSummarize.length > 0 && summaryBudget === 0 && recentKept.length === 0 && otherBudget > 0) {
    summaryBudget = Math.min(summaryMaxTokens, otherBudget);
  }

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
    const firstNonSystemIdx = out.findIndex(
      (m) => !(m && typeof m === "object" && m.role === "system" && !isGeneratedSummaryMessage(m))
    );
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

    if (preserveToolCallPairs) {
      const bounds = getToolCallGroupBounds(out, firstNonSystemIdx);
      if (bounds) {
        out.splice(bounds.start, bounds.end - bounds.start + 1);
        continue;
      }
    }

    out.splice(firstNonSystemIdx, 1);
  }

  // Ensure we never exceed the budget due to unexpected estimator behavior.
  if (estimator.estimateMessagesTokens(out) > allowedPromptTokens) {
    return takeTailWithinBudget(out, allowedPromptTokens, estimator, signal, { preserveToolCallPairs }).kept;
  }

  throwIfAborted(signal);
  return out;
}
