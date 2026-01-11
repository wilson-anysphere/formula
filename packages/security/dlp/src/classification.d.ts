export const CLASSIFICATION_LEVEL: Readonly<{
  PUBLIC: string;
  INTERNAL: string;
  CONFIDENTIAL: string;
  RESTRICTED: string;
}>;

export const CLASSIFICATION_LEVELS: readonly string[];

export const DEFAULT_CLASSIFICATION: Readonly<{ level: string; labels: string[] }>;

export function classificationRank(level: string): number;
export function normalizeClassification(classification: any): { level: string; labels: string[] };
export function maxClassification(a: any, b: any): { level: string; labels: string[] };
export function compareClassification(a: any, b: any): number;

