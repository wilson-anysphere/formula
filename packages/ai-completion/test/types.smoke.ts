import { suggestRanges, TabCompletionEngine, type PartialFormulaContext, type SurroundingCellsContext } from "@formula/ai-completion";

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

// Ensure TypeScript consumers can provide an async parsePartialFormula implementation.
const completion = new TabCompletionEngine({
  parsePartialFormula: async (_input, _cursorPosition, _registry): Promise<PartialFormulaContext> => {
    return { isFormula: false, inFunctionCall: false };
  },
});

// And a sync implementation remains valid.
new TabCompletionEngine({
  parsePartialFormula: (_input, _cursorPosition, _registry): PartialFormulaContext => {
    return { isFormula: false, inFunctionCall: false };
  },
});

void completion;
