export { createEngineClient } from "./client";
export type { EngineClient } from "./client";
export type { CellChange, CellData, CellScalar, RpcOptions } from "./protocol";
export { defaultWasmBinaryUrl, defaultWasmModuleUrl } from "./wasm";

export { EngineWorker } from "./EngineWorker";
export type { MessageChannelLike, MessagePortLike, WorkerLike } from "./EngineWorker";

export { WasmWorkbookBackend } from "./backend/WasmWorkbookBackend";
export type {
  WorkbookBackend,
  WorkbookInfo,
  SheetInfo,
  SheetUsedRange,
  RangeCellEdit,
  RangeData,
  CellValue,
} from "./backend/workbookBackend";

export {
  engineApplyDeltas,
  engineHydrateFromDocument,
  exportDocumentToEngineWorkbookJson,
} from "./documentControllerSync";
export type {
  DocumentCellDelta,
  DocumentCellState,
  EngineCellScalar,
  EngineSyncTarget,
  EngineWorkbookJson,
} from "./documentControllerSync";
