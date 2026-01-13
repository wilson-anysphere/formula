import { suggestRanges, type SurroundingCellsContext } from "@formula/ai-completion";

const surroundingCells: SurroundingCellsContext = {
  getCellValue: () => 1,
};

// Ensure TypeScript consumers can pass `maxScanCols` to expand 2D tables.
suggestRanges({
  currentArgText: "A",
  cellRef: { row: 0, col: 0 },
  surroundingCells,
  maxScanCols: 25,
});

