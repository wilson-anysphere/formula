export type ContextBudgetMode = "chat" | "agent" | "inline_edit";

/**
 * Best-effort model context window lookup.
 *
 * NOTE: This is intentionally conservative for unknown models. Callers can
 * override via `contextWindowTokens` in orchestrator options.
 */
export function getModelContextWindowTokens(model: string): number {
  const m = String(model ?? "").toLowerCase();

  // If the model name includes an explicit context-window hint (e.g. "32k", "128k"),
  // trust it. This keeps the logic deterministic without relying on any specific
  // model/provider naming scheme.
  const hintMatch = m.match(/(\d+)\s*k\b/i);
  if (hintMatch) {
    const thousands = Number(hintMatch[1]);
    if (Number.isFinite(thousands) && thousands > 0) return Math.floor(thousands * 1000);
  }

  // Cursor-managed backends generally support large context windows.
  if (m.includes("cursor")) return 128_000;

  // Conservative default for unknown models.
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
