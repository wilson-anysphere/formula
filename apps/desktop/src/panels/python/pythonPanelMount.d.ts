export function mountPythonPanel(params: {
  doc: any;
  container: HTMLElement;
  workbookId?: string;
  invoke?: (cmd: string, args?: any) => Promise<any>;
  drainBackendSync?: () => Promise<void>;
  getActiveSheetId?: () => string;
  getSelection?: () => { sheet_id: string; start_row: number; start_col: number; end_row: number; end_col: number };
  setSelection?: (selection: { sheet_id: string; start_row: number; start_col: number; end_row: number; end_col: number }) => void;
}): { dispose: () => void };

