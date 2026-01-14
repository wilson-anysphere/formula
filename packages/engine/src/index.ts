export { createEngineClient } from "./client.ts";
export type { EngineClient } from "./client.ts";
export type {
  CalcSettings,
  CalculationMode,
  IterativeCalcSettings,
  CellChange,
  CellData,
  CellDataCompact,
  CellDataRich,
  CellScalar,
  CellValueRich,
  PivotAggregationType,
  PivotCalculationResult,
  PivotConfig,
  PivotField,
  PivotFieldItems,
  PivotFieldRef,
  PivotFieldType,
  PivotGrandTotals,
  PivotLayout,
  PivotSchema,
  PivotSchemaField,
  PivotShowAsType,
  PivotSortOrder,
  PivotSubtotalPosition,
  PivotValue,
  PivotValueField,
  EditCellChange,
  EditCellSnapshot,
  EditFormulaRewrite,
  EditMovedRange,
  EditOp,
  EditResult,
  GoalSeekRecalcMode,
  GoalSeekResult,
  GoalSeekStatus,
  GoalSeekRequest,
  GoalSeekResponse,
  FormulaParseError,
  FormulaLocaleInfo,
  FormulaParseOptions,
  FormulaPartialLexResult,
  FormulaPartialParseResult,
  FormulaCoord,
  FormulaSpan,
  FormulaToken,
  FunctionContext,
  EngineInfoDto,
  FormatRun,
  RpcMethod,
  WorkbookStyleDto,
  RewriteFormulaForCopyDeltaRequest,
  RpcOptions,
} from "./protocol.ts";
export { defaultWasmBinaryUrl, defaultWasmModuleUrl } from "./wasm.ts";

export { EngineWorker } from "./EngineWorker.ts";
export type { MessageChannelLike, MessagePortLike, WorkerLike } from "./EngineWorker.ts";

export { WasmWorkbookBackend } from "./backend/WasmWorkbookBackend.ts";
export type {
  WorkbookBackend,
  WorkbookInfo,
  SheetInfo,
  SheetVisibility,
  TabColor,
  SheetUsedRange,
  RangeCellEdit,
  RangeData,
  CellValue,
} from "@formula/workbook-backend";

export { isFormulaInput, normalizeFormulaText, normalizeFormulaTextOpt } from "./backend/formula.ts";
export {
  isMissingGetLocaleInfoError,
  isMissingGetCellPhoneticError,
  isMissingGetRangeCompactError,
  isMissingSetCellPhoneticError,
  isMissingSupportedLocaleIdsError,
} from "./compat.ts";

export {
  EXCEL_DEFAULT_CELL_PADDING_PX,
  EXCEL_DEFAULT_MAX_DIGIT_WIDTH_PX,
  excelColWidthCharsToPixels,
  pixelsToExcelColWidthChars,
  type ExcelColumnWidthConversionOptions,
} from "./columnWidth.ts";

export {
  engineApplyDeltas,
  engineApplyDocumentChange,
  engineHydrateFromDocument,
  exportDocumentToEngineWorkbookJson,
} from "./documentControllerSync.ts";
export type {
  DocumentCellDelta,
  DocumentCellState,
  EngineApplyDocumentChangeOptions,
  DocumentControllerChangePayload,
  EngineCellScalar,
  EngineSyncTarget,
  EngineWorkbookJson,
} from "./documentControllerSync.ts";
