export class DlpViolationError extends Error {
  decision: any;
  constructor(decision: any);
}

export function formatDlpDecisionMessage(decision: any): string;

