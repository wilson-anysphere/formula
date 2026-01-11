export type CellChange = { sheetId: string; row: number; col: number; cell: any };

export interface Engine {
  applyChanges(changes: readonly CellChange[]): void;
  recalculate(): void;
  beginBatch?: () => void;
  endBatch?: () => void;
}

export class MockEngine implements Engine {
  sheets: Map<string, any>;
  recalcCount: number;
  batchDepth: number;
  appliedChanges: CellChange[];

  applyChanges(changes: readonly CellChange[]): void;
  recalculate(): void;
  beginBatch(): void;
  endBatch(): void;

  getCell(sheetId: string, row: number, col: number): any;
}

