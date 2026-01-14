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
    /**
     * Transaction origins that should be treated as local:
     * - Conflicts are not emitted for these transactions.
     * - Local edit tracking is updated from observed Yjs changes so later remote
     *   overwrites can be detected causally (even if callers don't use
     *   `setLocalFormula` / `setLocalValue`).
     */
    localOrigins?: Set<any>;
    /**
     * Transaction origins that should be ignored entirely (no conflicts emitted,
     * no local-edit tracking updates).
     */
    ignoredOrigins?: Set<any>;
    onConflict: (conflict: FormulaConflict) => void;
    getCellValue?: (ref: { sheetId: string; row: number; col: number }) => any;
    /**
     * Deprecated/ignored. Former wall-clock heuristic for inferring concurrency.
     *
     * Conflict detection is now causal (Yjs-based) and works across long offline periods.
     *
     * @deprecated
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
  /**
   * Best-effort id of the remote user who made the conflicting structural change.
   *
   * May be an empty string when unavailable (e.g. legacy/unknown op records).
   */
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
    /**
     * Transaction origins that should be ignored entirely (no conflicts emitted,
     * no local-edit tracking updates).
     */
    ignoredOrigins?: Set<any>;
    onConflict: (conflict: CellStructuralConflict) => void;
    maxOpRecordsPerUser?: number;
    /**
     * Optional age-based pruning window (in milliseconds) for records in the shared
     * `cellStructuralOps` log. When set, records older than `Date.now() - maxOpRecordAgeMs`
     * may be deleted by any client (best-effort).
     *
     * Pruning is conservative: records are not deleted in the same op-log transaction
     * they are added, so late-arriving/offline records have a chance to be ingested
     * by other clients before being removed.
     *
     * Defaults to null/disabled.
     */
    maxOpRecordAgeMs?: number | null;
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
    /**
     * Transaction origins that should be treated as local:
     * - Conflicts are not emitted for these transactions.
     * - Local edit tracking is updated from observed Yjs changes so later remote
     *   overwrites can be detected causally (even if callers don't use
     *   `setLocalValue`).
     */
    localOrigins?: Set<any>;
    /**
     * Transaction origins that should be ignored entirely (no conflicts emitted,
     * no local-edit tracking updates).
     */
    ignoredOrigins?: Set<any>;
    onConflict: (conflict: CellConflict) => void;
  });

  dispose(): void;
  listConflicts(): Array<CellConflict>;
  setLocalValue(cellKey: string, value: any): void;
  resolveConflict(conflictId: string, chosenValue: any): boolean;
}
