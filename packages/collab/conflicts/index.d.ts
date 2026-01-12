import type * as Y from "yjs";

export type CellContentChoice = { type: "formula"; formula: string; preview?: any } | { type: "value"; value: any };

export type FormulaConflict =
  | {
      kind: "formula";
      id: string;
      cell: { sheetId: string; row: number; col: number };
      cellKey: string;
      localFormula: string;
      remoteFormula: string;
      /**
       * Best-effort id of the remote user who overwrote the local change.
       *
       * May be an empty string when unavailable (e.g. legacy writers that do not
       * update the cell's `modifiedBy` field).
       */
      remoteUserId: string;
      detectedAt: number;
      localPreview?: any;
      remotePreview?: any;
    }
  | {
      kind: "value";
      id: string;
      cell: { sheetId: string; row: number; col: number };
      cellKey: string;
      localValue: any;
      remoteValue: any;
      /**
       * Best-effort id of the remote user who overwrote the local change.
       *
       * May be an empty string when unavailable (e.g. legacy writers that do not
       * update the cell's `modifiedBy` field).
       */
      remoteUserId: string;
      detectedAt: number;
    }
  | {
      kind: "content";
      id: string;
      cell: { sheetId: string; row: number; col: number };
      cellKey: string;
      local: CellContentChoice;
      remote: CellContentChoice;
      /**
       * Best-effort id of the remote user who overwrote the local change.
       *
       * May be an empty string when unavailable (e.g. legacy writers that do not
       * update the cell's `modifiedBy` field).
       */
      remoteUserId: string;
      detectedAt: number;
    };

export interface CellConflict {
  id: string;
  cell: { sheetId: string; row: number; col: number };
  cellKey: string;
  field: "value";
  localValue: any;
  remoteValue: any;
  /**
   * Best-effort id of the remote user who overwrote the local change.
   *
   * May be an empty string when unavailable (e.g. legacy writers that do not
   * update the cell's `modifiedBy` field).
   */
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
    /**
     * Deprecated/ignored. Former wall-clock heuristic for inferring concurrency.
     *
     * Conflict detection is now causal (Yjs-based) and works across long offline periods.
     */
    concurrencyWindowMs?: number;
    mode?: "formula" | "formula+value";
    includeValueConflicts?: boolean;
  });

  dispose(): void;
  listConflicts(): Array<FormulaConflict>;
  setLocalFormula(cellKey: string, formula: string): void;
  setLocalValue(cellKey: string, value: any): void;
  resolveConflict(conflictId: string, chosen: string | CellContentChoice | any): boolean;
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
    maxOpRecordsPerUser?: number;
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
