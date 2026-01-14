export interface TokenEstimator {
  estimateTextTokens(text: string): number;
  estimateMessageTokens(message: unknown): number;
  estimateMessagesTokens(messages: unknown[]): number;
}

export function estimateTokens(text: string, estimator?: TokenEstimator): number;

export function stableJsonStringify(value: unknown): string;

export function createHeuristicTokenEstimator(options?: {
  charsPerToken?: number;
  tokensPerMessageOverhead?: number;
}): TokenEstimator;

export const DEFAULT_TOKEN_ESTIMATOR: TokenEstimator;

export function estimateMessagesTokens(messages: unknown[], estimator?: TokenEstimator): number;

export function estimateToolDefinitionTokens(tools: unknown[] | null | undefined, estimator?: TokenEstimator): number;

export function trimToTokenBudget(
  text: string,
  maxTokens: number,
  estimator?: TokenEstimator,
  options?: { signal?: AbortSignal },
): string;

export function packSectionsToTokenBudget(
  sections: Array<{ key: string; text: string; priority: number }>,
  maxTokens: number,
  estimator?: TokenEstimator,
  options?: { signal?: AbortSignal },
): Array<{ key: string; text: string }>;

export interface TokenBudgetSectionReport {
  key: string;
  priority: number;
  tokensPreTrim: number;
  tokensPostTrim: number;
  trimmed: boolean;
  dropped: boolean;
}

export interface TokenBudgetReport {
  maxTokens: number;
  remainingTokens: number;
  sections: TokenBudgetSectionReport[];
}

export function packSectionsToTokenBudgetWithReport(
  sections: Array<{ key: string; text: string; priority: number }>,
  maxTokens: number,
  estimator?: TokenEstimator,
  options?: { signal?: AbortSignal },
): { packed: Array<{ key: string; text: string }>; report: TokenBudgetReport };
