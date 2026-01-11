import type * as Y from "yjs";

export interface FormulaConflict {
  id: string;
  cell: { sheetId: string; row: number; col: number };
  cellKey: string;
  localFormula: string;
  remoteFormula: string;
  remoteUserId: string;
  detectedAt: number;
  localPreview?: any;
  remotePreview?: any;
}

export class FormulaConflictMonitor {
  constructor(opts: {
    doc: Y.Doc;
    localUserId: string;
    cells?: Y.Map<any>;
    origin?: object;
    localOrigins?: Set<any>;
    onConflict: (conflict: FormulaConflict) => void;
    getCellValue?: (ref: { sheetId: string; row: number; col: number }) => any;
    concurrencyWindowMs?: number;
  });

  dispose(): void;
  listConflicts(): Array<FormulaConflict>;
  setLocalFormula(cellKey: string, formula: string): void;
  resolveConflict(conflictId: string, chosenFormula: string): boolean;
}

