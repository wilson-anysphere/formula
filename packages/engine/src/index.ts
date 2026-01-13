export { createEngineClient } from "./client.ts";
export type { EngineClient } from "./client.ts";
export type {
  CellChange,
  CellData,
  CellDataRich,
  CellScalar,
  CellValueRich,
  EditCellChange,
  EditCellSnapshot,
  EditFormulaRewrite,
  EditMovedRange,
  EditOp,
  EditResult,
  GoalSeekRecalcMode,
  GoalSeekRequest,
  GoalSeekResponse,
  FormulaParseError,
  FormulaParseOptions,
  FormulaPartialLexResult,
  FormulaPartialParseResult,
  FormulaCoord,
  FormulaSpan,
  FormulaToken,
  FunctionContext,
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
  engineApplyDeltas,
  engineHydrateFromDocument,
  exportDocumentToEngineWorkbookJson,
} from "./documentControllerSync.ts";
export type {
  DocumentCellDelta,
  DocumentCellState,
  EngineCellScalar,
  EngineSyncTarget,
  EngineWorkbookJson,
} from "./documentControllerSync.ts";
