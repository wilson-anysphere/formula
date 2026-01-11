export type OutlineHiddenState = { user: boolean; outline: boolean; filter: boolean };

export type OutlineEntry = { level: number; hidden: OutlineHiddenState; collapsed: boolean };

export class OutlineAxis {
  entries: Map<number, OutlineEntry>;
  entry(index: number): OutlineEntry;
  entryMut(index: number): OutlineEntry;
  clearOutlineHidden(): void;
}

export function isHidden(hidden: OutlineHiddenState): boolean;

export function groupDetailRange(
  axis: OutlineAxis,
  summaryIndex: number,
  summaryLevel: number,
  summaryAfterDetails: boolean
): [start: number, end: number, level: number] | null;

export class Outline {
  pr: { summaryBelow: boolean; summaryRight: boolean; showOutlineSymbols: boolean };
  rows: OutlineAxis;
  cols: OutlineAxis;
  toggleRowGroup(summaryIndex: number): void;
  toggleColGroup(summaryIndex: number): void;
  groupRows(start: number, end: number): void;
  ungroupRows(start: number, end: number): void;
  groupCols(start: number, end: number): void;
  ungroupCols(start: number, end: number): void;
  recomputeOutlineHiddenRows(): void;
  recomputeOutlineHiddenCols(): void;
}

