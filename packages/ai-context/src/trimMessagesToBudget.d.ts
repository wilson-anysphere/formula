import type { TokenEstimator } from "./tokenBudget.js";

export const CONTEXT_SUMMARY_MARKER: string;

export function trimMessagesToBudget(params: {
  messages: any[];
  maxTokens: number;
  reserveForOutputTokens?: number;
  estimator?: TokenEstimator;
  keepLastMessages?: number;
  summaryMaxTokens?: number;
  summarize?:
    | null
    | ((messagesToSummarize: any[]) => string | any | null | undefined | Promise<string | any | null | undefined>);
  summaryRole?: "system" | "assistant";
  signal?: AbortSignal;
}): Promise<any[]>;
