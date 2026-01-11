export type ContextBudgetMode = "chat" | "agent" | "inline_edit";

/**
 * Best-effort model context window lookup.
 *
 * NOTE: This is intentionally conservative for unknown models. Callers can
 * override via `contextWindowTokens` in orchestrator options.
 */
export function getModelContextWindowTokens(model: string): number {
  const m = String(model ?? "").toLowerCase();

  // OpenAI
  if (m.startsWith("gpt-4o") || m.startsWith("gpt-4.1") || m.includes("gpt-4-turbo")) return 128_000;
  if (m.startsWith("gpt-4")) return 32_000;
  if (m.startsWith("gpt-3.5")) return 16_000;

  // Anthropic
  if (m.startsWith("claude-3") || m.startsWith("claude-2")) return 200_000;

  // Default fallback (kept modest to avoid blowing up smaller provider limits).
  return 16_000;
}

/**
 * Effective context window for a given UI surface.
 *
 * Inline edit is intentionally capped to keep prompts small and latency low.
 */
export function getModeContextWindowTokens(mode: ContextBudgetMode, model: string): number {
  const base = getModelContextWindowTokens(model);
  if (mode === "inline_edit") return Math.min(base, 4_096);
  return base;
}

/**
 * Reserve some tokens for the completion to avoid "prompt too long" errors.
 */
export function getDefaultReserveForOutputTokens(mode: ContextBudgetMode, contextWindowTokens: number): number {
  const total = Math.max(0, contextWindowTokens);
  if (total === 0) return 0;

  const min = mode === "inline_edit" ? 256 : 512;
  const max = mode === "inline_edit" ? 1024 : 4096;
  const fraction = mode === "inline_edit" ? 0.05 : 0.1;
  return clamp(Math.floor(total * fraction), min, max);
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

