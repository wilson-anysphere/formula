export type DlpFinding = "email" | "ssn" | "credit_card" | "phone_number" | "api_key" | "iban";
export type DlpLevel = "public" | "sensitive";

export function classifyText(text: string): { level: DlpLevel; findings: DlpFinding[] };
export function redactText(text: string): string;
