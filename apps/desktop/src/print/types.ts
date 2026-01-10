export type Orientation = "portrait" | "landscape";

// Mirrors OpenXML ST_PaperSize codes where possible (1 = Letter, 9 = A4).
export type PaperSizeCode = number;

export type Scaling =
  | { kind: "percent"; percent: number }
  | { kind: "fitTo"; widthPages: number; heightPages: number };

export type PageMarginsInches = {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
};

export type PageSetup = {
  orientation: Orientation;
  paperSize: PaperSizeCode;
  margins: PageMarginsInches;
  scaling: Scaling;
};

// 1-based inclusive coordinates (Excel-style).
export type CellRange = {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
};

export type PrintTitles = {
  repeatRows?: { start: number; end: number };
  repeatCols?: { start: number; end: number };
};

export type ManualPageBreaks = {
  rowBreaksAfter: number[]; // 1-based row number after which a break occurs
  colBreaksAfter: number[]; // 1-based column number after which a break occurs
};

export type SheetPrintSettings = {
  sheetName: string;
  printArea?: CellRange[];
  printTitles?: PrintTitles;
  pageSetup: PageSetup;
  manualPageBreaks?: ManualPageBreaks;
};

export type Page = {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
};

