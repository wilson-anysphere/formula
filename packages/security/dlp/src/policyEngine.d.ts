export const DLP_DECISION: Readonly<{
  ALLOW: "allow";
  BLOCK: "block";
  REDACT: "redact";
}>;

export const DLP_REASON_CODE: Readonly<{
  BLOCKED_BY_POLICY: string;
  INVALID_POLICY: string;
}>;

export function evaluatePolicy(params: {
  action: string;
  classification: { level: string; labels?: string[] };
  policy: any;
  options?: { includeRestrictedContent?: boolean };
}): any;

export function isClassificationAllowed(classification: { level: string }, maxAllowed: string | null): boolean;

