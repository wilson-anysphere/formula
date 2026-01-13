export type YjsSpreadsheetDocAdapter = {
  encodeState(): Uint8Array;
  applyState(snapshot: Uint8Array): void;
  on(event: "update", listener: () => void): () => void;
};

export function createYjsSpreadsheetDocAdapter(...args: any[]): YjsSpreadsheetDocAdapter;
