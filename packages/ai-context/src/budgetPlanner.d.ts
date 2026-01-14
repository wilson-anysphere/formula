import type { TokenEstimator } from "./tokenBudget.js";

export type SectionTarget = { key: string; tokens: number };

export type TokenBudgetPlan = {
  total: number;
  reserved: number;
  available: number;
  systemPromptTokens: number;
  toolDefinitionTokens: number;
  messageTokens: number;
  fixedPromptTokens: number;
  /**
   * Remaining tokens available for caller-provided sheet/workbook context after
   * accounting for reserve + fixed prompt overhead.
   */
  remainingForContextTokens: number;
  sections?: {
    allocations: Array<{ key: string; targetTokens: number; allocatedTokens: number }>;
    allocationByKey: Record<string, number>;
    totalTargetTokens: number;
    allocatedTokens: number;
    unallocatedTokens: number;
  };
};

export function planTokenBudget(params: {
  maxContextTokens: number;
  reserveForOutputTokens: number;
  systemPrompt?: string;
  tools?: unknown[];
  messages?: unknown[];
  estimator?: TokenEstimator;
  sectionTargets?: Record<string, number> | SectionTarget[];
}): TokenBudgetPlan;
