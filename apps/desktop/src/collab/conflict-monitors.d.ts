export const VERSIONING_RESTORE_ORIGIN: string;
export const BRANCHING_APPLY_ORIGIN: string;

import type * as Y from "yjs";
import type {
  CellStructuralConflict,
  CellStructuralConflictMonitor,
  FormulaConflict,
  FormulaConflictMonitor,
} from "../../../../packages/collab/conflicts/index.js";

export function createDesktopFormulaConflictMonitor(opts: {
  doc: Y.Doc;
  cells?: Y.Map<any>;
  localUserId: string;
  sessionOrigin: any;
  binderOrigin: any;
  undoLocalOrigins?: Set<any>;
  onConflict: (conflict: FormulaConflict) => void;
  getCellValue?: (ref: { sheetId: string; row: number; col: number }) => any;
  mode?: "formula" | "formula+value";
}): FormulaConflictMonitor;

export function createDesktopCellStructuralConflictMonitor(opts: {
  doc: Y.Doc;
  cells?: Y.Map<any>;
  localUserId: string;
  sessionOrigin: any;
  binderOrigin: any;
  undoLocalOrigins?: Set<any>;
  onConflict: (conflict: CellStructuralConflict) => void;
  maxOpRecordsPerUser?: number;
  maxOpRecordAgeMs?: number | null;
}): CellStructuralConflictMonitor;
