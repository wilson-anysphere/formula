export class DocumentControllerWorkbookAdapter {
  constructor(
    documentController: any,
    options?: {
      activeSheetName?: string;
      getActiveSheetName?: () => string;
      getSelection?: () => { sheetName: string; address: string };
      setSelection?: (sheetName: string, address: string) => void;
      onDidMutate?: () => void;
    }
  );

  documentController: any;
  events: any;
  sheets: Map<string, any>;
  activeSheetName: string;
  selection: { sheetName: string; address: string } | null;

  dispose(): void;
  getActiveSheet(): any;
  getSheet(sheetName: string): any;
}

