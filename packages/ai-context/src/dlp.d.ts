export type DlpFinding = "email" | "ssn" | "credit_card";
export type DlpLevel = "public" | "sensitive";

export function classifyText(text: string): { level: DlpLevel; findings: DlpFinding[] };

export function redactText(text: string): string;
