export type FindReplaceCellRef = { sheetName: string; row: number; col: number };
export type FindReplaceSelectionRange = { startRow: number; startCol: number; endRow: number; endCol: number };

export type FindReplaceControllerParams = {
  workbook: any;
  getCurrentSheetName?: () => string;
  getActiveCell?: () => FindReplaceCellRef;
  setActiveCell?: (cell: FindReplaceCellRef) => void;
  getSelectionRanges?: () => FindReplaceSelectionRange[];
  beginBatch?: (options?: unknown) => void;
  endBatch?: () => void;
  /**
   * Optional guard for replace operations.
   *
   * When provided and it returns `false` (or `{ allowed: false }`), `replaceNext`/`replaceAll`
   * will no-op and (when `showToast` is provided) surface the rejection reason to the user.
   */
  canReplace?: () => boolean | { allowed: boolean; reason?: string | null };
  /**
   * Optional toast helper for user-facing replace guard failures.
   */
  showToast?: (message: string, type?: any) => void;
};

export class FindReplaceController {
  constructor(params: FindReplaceControllerParams);
  query: string;
  replacement: string;
  scope: string;
  lookIn: string;
  valueMode: string;
  matchCase: boolean;
  matchEntireCell: boolean;
  useWildcards: boolean;
  searchOrder: string;

  findNext(): Promise<any>;
  findAll(): Promise<any[]>;
  replaceNext(): Promise<any>;
  replaceAll(): Promise<any>;
}

export type RegisterFindReplaceShortcutsParams = {
  controller: FindReplaceController;
  workbook: any;
  getCurrentSheetName: () => string;
  setActiveCell: (cell: FindReplaceCellRef) => void;
  selectRange: (selection: { sheetName: string; range: FindReplaceSelectionRange }) => void;
  mount?: HTMLElement;
};

export type FindReplaceShortcutDialogs = {
  findDialog: HTMLDialogElement;
  replaceDialog: HTMLDialogElement;
  goToDialog: HTMLDialogElement;
};

export function registerFindReplaceShortcuts(params: RegisterFindReplaceShortcutsParams): FindReplaceShortcutDialogs;
