import {
  DEFAULT_TOKEN_ESTIMATOR,
  estimateMessagesTokens,
  estimateTokens,
  estimateToolDefinitionTokens,
} from "./tokenBudget.js";

/**
 * @typedef {{ key: string, tokens: number }} SectionTarget
 */

/**
 * Normalize a section targets input into a deterministic list.
 *
 * - For object inputs, keys are sorted so allocations are stable regardless of
 *   insertion order.
 * - For array inputs, the given order is preserved.
 *
 * @param {Record<string, number> | SectionTarget[] | null | undefined} sectionTargets
 * @returns {Array<{ key: string, targetTokens: number, index: number }>}
 */
function normalizeSectionTargets(sectionTargets) {
  if (Array.isArray(sectionTargets)) {
    return sectionTargets
      .map((t, index) => ({
        key: typeof t?.key === "string" ? t.key : "",
        targetTokens: Number.isFinite(t?.tokens) ? Math.max(0, Math.trunc(t.tokens)) : 0,
        index,
      }))
      .filter((t) => t.key);
  }

  if (!sectionTargets || typeof sectionTargets !== "object") return [];

  const obj = /** @type {Record<string, number>} */ (sectionTargets);
  const keys = Object.keys(obj).sort();
  return keys
    .map((key, index) => ({
      key,
      targetTokens: Number.isFinite(obj[key]) ? Math.max(0, Math.trunc(obj[key])) : 0,
      index,
    }))
    .filter((t) => t.key);
}

/**
 * Allocate tokens across targets, preserving the target proportions when the
 * budget is too small.
 *
 * @param {number} availableTokens
 * @param {Array<{ key: string, targetTokens: number, index: number }>} targets
 * @returns {{
 *   allocations: Array<{ key: string, targetTokens: number, allocatedTokens: number }>,
 *   allocationByKey: Record<string, number>,
 *   totalTargetTokens: number,
 *   allocatedTokens: number,
 *   unallocatedTokens: number
 * }}
 */
function allocateSectionTokens(availableTokens, targets) {
  const total = Math.max(0, Math.trunc(availableTokens));
  const totalTargetTokens = targets.reduce((sum, t) => sum + (t.targetTokens ?? 0), 0);

  /** @type {number[]} */
  const allocated = new Array(targets.length).fill(0);

  if (total <= 0 || targets.length === 0) {
    const allocations = targets.map((t) => ({ key: t.key, targetTokens: t.targetTokens, allocatedTokens: 0 }));
    return {
      allocations,
      allocationByKey: Object.fromEntries(allocations.map((a) => [a.key, a.allocatedTokens])),
      totalTargetTokens,
      allocatedTokens: 0,
      unallocatedTokens: 0,
    };
  }

  // If the full target budget fits, allocate exactly and return the remainder.
  if (totalTargetTokens <= total) {
    for (let i = 0; i < targets.length; i += 1) allocated[i] = targets[i].targetTokens;
    const allocations = targets.map((t, i) => ({ key: t.key, targetTokens: t.targetTokens, allocatedTokens: allocated[i] }));
    const allocatedTokens = totalTargetTokens;
    return {
      allocations,
      allocationByKey: Object.fromEntries(allocations.map((a) => [a.key, a.allocatedTokens])),
      totalTargetTokens,
      allocatedTokens,
      unallocatedTokens: total - allocatedTokens,
    };
  }

  // Scale proportionally to fit, distributing rounding remainder deterministically.
  const ratio = totalTargetTokens > 0 ? total / totalTargetTokens : 0;
  /** @type {Array<{ idx: number, frac: number, key: string, index: number }>} */
  const remainderRanking = [];
  let floorSum = 0;

  for (let i = 0; i < targets.length; i += 1) {
    const t = targets[i];
    const raw = t.targetTokens * ratio;
    const floored = Math.floor(raw);
    allocated[i] = floored;
    floorSum += floored;
    const frac = raw - floored;
    // Skip 0-target sections from ever receiving remainder tokens.
    if (t.targetTokens > 0) {
      remainderRanking.push({ idx: i, frac, key: t.key, index: t.index });
    }
  }

  let remaining = total - floorSum;
  if (remaining > 0 && remainderRanking.length > 0) {
    remainderRanking.sort((a, b) => {
      if (b.frac !== a.frac) return b.frac - a.frac;
      if (a.index !== b.index) return a.index - b.index;
      return a.key.localeCompare(b.key);
    });
    for (let i = 0; i < remainderRanking.length && remaining > 0; i += 1) {
      const idx = remainderRanking[i].idx;
      if (allocated[idx] < targets[idx].targetTokens) {
        allocated[idx] += 1;
        remaining -= 1;
      }
    }
  }

  const allocations = targets.map((t, i) => ({ key: t.key, targetTokens: t.targetTokens, allocatedTokens: allocated[i] }));
  const allocatedTokens = allocations.reduce((sum, a) => sum + a.allocatedTokens, 0);
  return {
    allocations,
    allocationByKey: Object.fromEntries(allocations.map((a) => [a.key, a.allocatedTokens])),
    totalTargetTokens,
    allocatedTokens,
    unallocatedTokens: Math.max(0, total - allocatedTokens),
  };
}

/**
 * Plan a token budget for an LLM request.
 *
 * This is a lightweight, deterministic helper for callers that need to split a
 * model context window into:
 * - reserved output tokens
 * - "fixed" prompt overhead (system prompt + tools + messages)
 * - remaining tokens available for sheet/workbook context sections
 *
 * @param {{
 *   maxContextTokens: number,
 *   reserveForOutputTokens: number,
 *   systemPrompt?: string,
 *   tools?: any[],
 *   messages?: any[],
 *   estimator?: import("./tokenBudget.js").TokenEstimator,
 *   sectionTargets?: Record<string, number> | SectionTarget[]
 * }} params
 */
export function planTokenBudget(params) {
  const total = Number.isFinite(params?.maxContextTokens) ? Math.max(0, Math.trunc(params.maxContextTokens)) : 0;
  const reserveRequested = Number.isFinite(params?.reserveForOutputTokens)
    ? Math.max(0, Math.trunc(params.reserveForOutputTokens))
    : 0;
  const reserved = Math.min(total, reserveRequested);
  const available = Math.max(0, total - reserved);
  const estimator = params?.estimator ?? DEFAULT_TOKEN_ESTIMATOR;

  const systemPrompt = typeof params?.systemPrompt === "string" ? params.systemPrompt : "";
  const tools = Array.isArray(params?.tools) ? params.tools : [];
  const messages = Array.isArray(params?.messages) ? params.messages : [];

  const systemPromptTokens = systemPrompt ? estimateTokens(systemPrompt, estimator) : 0;
  const toolDefinitionTokens = estimateToolDefinitionTokens(tools, estimator);
  const messageTokens = estimateMessagesTokens(messages, estimator);

  const fixedPromptTokens = systemPromptTokens + toolDefinitionTokens + messageTokens;
  const remainingForContextTokens = Math.max(0, available - fixedPromptTokens);

  const targets = normalizeSectionTargets(params?.sectionTargets);
  const sections = targets.length ? allocateSectionTokens(remainingForContextTokens, targets) : null;

  return {
    total,
    reserved,
    available,
    systemPromptTokens,
    toolDefinitionTokens,
    messageTokens,
    fixedPromptTokens,
    remainingForContextTokens,
    ...(sections ? { sections } : {}),
  };
}
