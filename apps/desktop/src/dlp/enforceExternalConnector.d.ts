export function enforceExternalConnector(args: {
  documentId: string;
  sheetId?: string;
  range?: unknown;
  classificationStore: { list: (documentId: string) => Array<{ selector: unknown; classification: unknown }> };
  policy: unknown | (() => unknown | Promise<unknown>);
}): { decision: unknown; selectionClassification: unknown };
