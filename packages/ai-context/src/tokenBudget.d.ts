export interface TokenEstimator {
  estimateTextTokens(text: string): number;
  estimateMessageTokens(message: any): number;
  estimateMessagesTokens(messages: any[]): number;
}

export function estimateTokens(text: string, estimator?: TokenEstimator): number;

export function stableJsonStringify(value: unknown): string;

export function createHeuristicTokenEstimator(options?: {
  charsPerToken?: number;
  tokensPerMessageOverhead?: number;
}): TokenEstimator;

export const DEFAULT_TOKEN_ESTIMATOR: TokenEstimator;

export function estimateMessagesTokens(messages: any[], estimator?: TokenEstimator): number;

export function estimateToolDefinitionTokens(tools: any[] | null | undefined, estimator?: TokenEstimator): number;

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
