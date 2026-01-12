export type ClipboardSelectionRange = {
  start: { row: number; col: number };
  end: { row: number; col: number };
};

export type ClipboardCopyClassificationStore = {
  list(documentId: string): Array<{ selector: any; classification: any }>;
};

export type EnforceClipboardCopyParams = {
  documentId: string;
  sheetId: string;
  range: ClipboardSelectionRange;
  classificationStore: ClipboardCopyClassificationStore;
  policy: any;
};

export function enforceClipboardCopy(params: EnforceClipboardCopyParams): {
  decision: any;
  selectionClassification: any;
};

