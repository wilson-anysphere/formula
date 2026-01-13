export { HashEmbedder } from "./embedding/hashEmbedder.js";

export { InMemoryVectorStore } from "./store/inMemoryVectorStore.js";
export { JsonVectorStore } from "./store/jsonVectorStore.js";
export { SqliteVectorStore } from "./store/sqliteVectorStore.js";
export type {
  SqliteVectorStoreDimensionMismatchError,
  SqliteVectorStoreInvalidMetadataError,
} from "./store/sqliteVectorStore.js";

export {
  InMemoryBinaryStorage,
  LocalStorageBinaryStorage,
  ChunkedLocalStorageBinaryStorage,
  IndexedDBBinaryStorage,
} from "./store/binaryStorage.js";
export type { BinaryStorage } from "./store/binaryStorage.js";

export { chunkWorkbook } from "./workbook/chunkWorkbook.js";
export { chunkToText } from "./workbook/chunkToText.js";
export { cellToA1, rectToA1 } from "./workbook/rect.js";
export { workbookFromSpreadsheetApi } from "./workbook/fromSpreadsheetApi.js";

export { indexWorkbook, approximateTokenCount } from "./pipeline/indexWorkbook.js";

export { searchWorkbookRag } from "./retrieval/searchWorkbookRag.js";
export { rerankWorkbookResults, dedupeOverlappingResults } from "./retrieval/rankResults.js";
