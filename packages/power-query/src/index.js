/**
 * `@formula/power-query`
 *
 * Formula's Power Query engine (query model + execution engine + refresh orchestration).
 *
 * Public consumers should import from this entrypoint rather than deep-linking into
 * individual files inside `src/`.
 *
 * Optional dependency:
 * - `@formula/data-io` is only required for Arrow + Parquet helpers. The core engine
 *   can execute CSV/JSON/API/SQL sources without it, but Arrow IPC caching / parquet-bytes
 *   loading will throw if the optional dependency is missing.
 */

export { DataTable } from "./table.js";
export { ArrowTableAdapter } from "./arrowTable.js";
export { QueryEngine } from "./engine.js";
export { applyOperation, compileRowFormula } from "./steps.js";
export { valueKey } from "./valueKey.js";
export { QueryFoldingEngine } from "./folding/sql.js";
export { ODataFoldingEngine, buildODataUrl } from "./folding/odata.js";
export { RefreshManager, QueryScheduler } from "./refresh.js";
export { RefreshOrchestrator, computeQueryDependencies } from "./refreshGraph.js";
export { InMemoryRefreshStateStore } from "./refreshStateStore.js";
export { parseCronExpression } from "./cron.js";
export { InMemorySheet, writeTableToSheet } from "./sheet.js";
export {
  FileConnector,
  HttpConnector,
  ODataConnector,
  SharePointConnector,
  SqlConnector,
  decodeBinaryText,
  decodeBinaryTextStream,
  parseCsv,
  parseCsvCell,
  parseCsvStream,
  parseCsvStreamBatches,
} from "./connectors/index.js";
export { CacheManager } from "./cache/cache.js";
export { MemoryCacheStore } from "./cache/memory.js";
export { FileSystemCacheStore } from "./cache/filesystem.js";
export { IndexedDBCacheStore } from "./cache/indexeddb.js";
export { EncryptedCacheStore } from "./cache/encryptedStore.js";
export { createWebCryptoCacheProvider } from "./cache/webCryptoProvider.js";
export { stableStringify, fnv1a64, hashValue } from "./cache/key.js";
export {
  OAuth2Manager,
  OAuth2TokenClient,
  OAuth2TokenError,
  InMemoryOAuthTokenStore,
  CredentialStoreOAuthTokenStore,
  createCodeVerifier,
  createCodeChallenge,
  normalizeScopes,
} from "./oauth2/index.js";

// Privacy levels / formula firewall helpers.
export { getPrivacyLevel, privacyRank } from "./privacy/levels.js";
export {
  getFileSourceId,
  getHttpSourceId,
  getSharePointSourceId,
  getSqlSourceId,
  getSourceIdForProvenance,
  getSourceIdForQuerySource,
  normalizeFilePath,
} from "./privacy/sourceId.js";

export { parseM, compileMToQuery, prettyPrintQueryToM } from "./m/index.js";

export * from "./credentials/index.js";

export * from "./values.js";
