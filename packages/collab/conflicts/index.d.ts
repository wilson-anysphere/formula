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

export interface CellConflict {
  id: string;
  cell: { sheetId: string; row: number; col: number };
  cellKey: string;
  field: "value";
  localValue: any;
  remoteValue: any;
  remoteUserId: string;
  detectedAt: number;
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

export interface CellStructuralConflict {
  id: string;
  type: "move" | "cell";
  reason: "move-destination" | "delete-vs-edit" | "content" | "format";
  sheetId: string;
  cell: string;
  cellKey: string;
  local: any;
  remote: any;
  remoteUserId: string;
  detectedAt: number;
}

export type CellStructuralConflictResolution = {
  choice: "ours" | "theirs" | "manual";
  to?: string;
  cell?: { value?: unknown; formula?: string; enc?: unknown; format?: Record<string, unknown> | null } | null;
};

export class CellStructuralConflictMonitor {
  constructor(opts: {
    doc: Y.Doc;
    localUserId: string;
    cells?: Y.Map<any>;
    origin?: any;
    localOrigins?: Set<any>;
    onConflict: (conflict: CellStructuralConflict) => void;
  });

  dispose(): void;
  listConflicts(): Array<CellStructuralConflict>;
  resolveConflict(conflictId: string, resolution: CellStructuralConflictResolution): boolean;
}

export class CellConflictMonitor {
  constructor(opts: {
    doc: Y.Doc;
    localUserId: string;
    cells?: Y.Map<any>;
    origin?: object;
    localOrigins?: Set<any>;
    onConflict: (conflict: CellConflict) => void;
  });

  dispose(): void;
  listConflicts(): Array<CellConflict>;
  setLocalValue(cellKey: string, value: any): void;
  resolveConflict(conflictId: string, chosenValue: any): boolean;
}
