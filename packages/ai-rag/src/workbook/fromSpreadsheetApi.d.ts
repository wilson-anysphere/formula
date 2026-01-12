export function workbookFromSpreadsheetApi(params: {
  spreadsheet: {
    listSheets(): string[];
    listNonEmptyCells(
      sheet?: string
    ): Array<{ address: { sheet: string; row: number; col: number }; cell: { value?: any; formula?: string } }>;
  };
  workbookId: string;
  coordinateBase?: "one" | "zero" | "auto";
  signal?: AbortSignal;
}): any;
