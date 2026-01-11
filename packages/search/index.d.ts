export type A1Address = {
  row: number;
  col: number;
  rowAbsolute: boolean;
  colAbsolute: boolean;
};

export type A1Range = {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
};

export function colToIndex(col: string): number;
export function indexToCol(index: number): string;
export function parseA1Address(input: string): A1Address;
export function formatA1Address(params: {
  row: number;
  col: number;
  rowAbsolute?: boolean;
  colAbsolute?: boolean;
}): string;

export function parseA1Range(input: string): A1Range;
export function formatA1Range(range: A1Range): string;

export function splitSheetQualifier(input: string): { sheetName: string | null; ref: string };

export type GoToParseResult = {
  type: "range";
  source: "table" | "a1" | "name";
  sheetName: string;
  range: A1Range;
};

export type GoToWorkbookLookup = {
  getTable(name: string): null | {
    sheetName: string;
    startRow: number;
    endRow: number;
    startCol: number;
    endCol: number;
    columns?: string[];
  };
  getName(name: string): null | { sheetName?: string; range: A1Range };
};

export function parseGoTo(
  input: string,
  options: { workbook: GoToWorkbookLookup; currentSheetName: string }
): GoToParseResult;

