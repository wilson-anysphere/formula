export type WorkbookSearchResult = {
  id: string;
  score: number;
  metadata?: {
    workbookId?: string;
    kind?: "table" | "namedRange" | "dataRegion" | "formulaRegion" | string;
    title?: string;
    sheetName?: string;
    rect?: { r0: number; c0: number; r1: number; c1: number };
    tokenCount?: number;
    [key: string]: any;
  };
  [key: string]: any;
};

export function rerankWorkbookResults<T extends WorkbookSearchResult>(
  query: string,
  results: T[],
  opts?: {
    kindBoost?: Partial<Record<string, number>>;
    titleTokenBoost?: number;
    sheetTokenBoost?: number;
    tokenPenaltyThreshold?: number;
    tokenPenaltyScale?: number;
    tokenPenaltyMax?: number;
  },
): T[];

export function dedupeOverlappingResults<T extends WorkbookSearchResult>(
  results: T[],
  opts?: {
    overlapRatioThreshold?: number;
  },
): T[];
