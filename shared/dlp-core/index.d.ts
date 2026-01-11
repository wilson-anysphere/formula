export const CLASSIFICATION_LEVEL: Readonly<{
  PUBLIC: "Public";
  INTERNAL: "Internal";
  CONFIDENTIAL: "Confidential";
  RESTRICTED: "Restricted";
}>;

export const CLASSIFICATION_LEVELS: readonly ["Public", "Internal", "Confidential", "Restricted"];

export type ClassificationLevel = (typeof CLASSIFICATION_LEVELS)[number];

export interface Classification {
  level: ClassificationLevel;
  labels: string[];
}

export const DEFAULT_CLASSIFICATION: Readonly<Classification>;

export interface DlpRuleBase {
  maxAllowed: ClassificationLevel | null;
}

export interface DlpAiRule extends DlpRuleBase {
  allowRestrictedContent: boolean;
  redactDisallowed: boolean;
}

export interface DlpPolicy {
  version: number;
  allowDocumentOverrides: boolean;
  rules: Record<string, DlpRuleBase | DlpAiRule>;
}

export const DLP_ACTION: Readonly<{
  SHARE_EXTERNAL_LINK: "sharing.externalLink";
  EXPORT_CSV: "export.csv";
  EXPORT_PDF: "export.pdf";
  EXPORT_XLSX: "export.xlsx";
  CLIPBOARD_COPY: "clipboard.copy";
  AI_CLOUD_PROCESSING: "ai.cloudProcessing";
  EXTERNAL_CONNECTOR: "connector.external";
}>;

export const DLP_POLICY_VERSION: 1;

export function validateDlpPolicy(policy: unknown): asserts policy is DlpPolicy;
export function normalizeDlpPolicy(policy: unknown): DlpPolicy;

export function normalizeClassification(classification: unknown): Classification;
export function classificationRank(level: ClassificationLevel): number;
export function compareClassification(a: unknown, b: unknown): -1 | 0 | 1;
export function maxClassification(a: unknown, b: unknown): Classification;

export type CellCoord = { row: number; col: number };
export type CellRange = { start: CellCoord; end: CellCoord };

export function normalizeRange(range: CellRange): CellRange;
export function selectorKey(selector: unknown): string;

export type ClassificationScope = "document" | "sheet" | "column" | "range" | "cell";

export type DocumentSelector = { scope: "document"; documentId: string };
export type SheetSelector = { scope: "sheet"; documentId: string; sheetId: string };
export type ColumnSelector = {
  scope: "column";
  documentId: string;
  sheetId: string;
  columnIndex?: number;
  tableId?: string;
  columnId?: string;
};
export type CellSelector = {
  scope: "cell";
  documentId: string;
  sheetId: string;
  row: number;
  col: number;
  tableId?: string;
  columnId?: string;
};
export type RangeSelector = { scope: "range"; documentId: string; sheetId: string; range: CellRange };

export type ClassificationSelector = DocumentSelector | SheetSelector | ColumnSelector | CellSelector | RangeSelector;

export function normalizeSelector(selector: unknown): ClassificationSelector;

export type ClassificationRecord = { selector: unknown; classification: unknown };

export type ResolvedClassificationMatch = {
  selector: ClassificationSelector;
  selectorKey: string;
  classification: Classification;
};

export interface ResolveClassificationOptions {
  includeMatchedSelectors?: boolean;
  maxMatchedSelectors?: number;
  maxRangeCellsForMatchedSelectors?: number;
}

export function resolveClassification(params: {
  querySelector: unknown;
  records: ClassificationRecord[];
  options?: ResolveClassificationOptions;
}): {
  effectiveClassification: Classification;
  matchedCount: number;
  matchedSelectors?: ResolvedClassificationMatch[];
};

export const DLP_DECISION: Readonly<{
  ALLOW: "allow";
  BLOCK: "block";
  REDACT: "redact";
}>;

export const DLP_REASON_CODE: Readonly<{
  BLOCKED_BY_POLICY: "dlp.blockedByPolicy";
  INVALID_POLICY: "dlp.invalidPolicy";
}>;

export type DlpDecision = (typeof DLP_DECISION)[keyof typeof DLP_DECISION];

export type PolicyEvaluationResult = {
  action: string;
  decision: DlpDecision;
  reasonCode?: string;
  classification: Classification;
  maxAllowed: ClassificationLevel | null;
};

export function evaluatePolicy(params: {
  action: string;
  classification: unknown;
  policy: unknown;
  options?: { includeRestrictedContent?: boolean };
}): PolicyEvaluationResult;

export function isClassificationAllowed(classification: unknown, maxAllowed: ClassificationLevel | null): boolean;
export function isAllowed(
  action: string,
  classification: unknown,
  policy: unknown,
  options?: { includeRestrictedContent?: boolean }
): boolean;

export const REDACTION_PLACEHOLDER: "[REDACTED]";
export function redact(value?: unknown, policyRule?: unknown): string;
