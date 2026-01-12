import {
  CacheManager,
  DataTable,
  EncryptedCacheStore,
  FileConnector,
  HttpConnector,
  MemoryCacheStore,
  MS_PER_DAY,
  ODataConnector,
  PqDateTimeZone,
  PqDecimal,
  PqDuration,
  PqTime,
  QueryEngine,
  RefreshManager,
  RefreshOrchestrator,
  SharePointConnector,
  SqlConnector,
  createWebCryptoCacheProvider,
  hashValue,
  httpScope,
  oauth2Scope,
  parseCronExpression,
  randomId,
} from "@formula/power-query";

import type { Query, QueryExecutionContext, QueryOperation, QuerySource } from "@formula/power-query";

import { createNodeCryptoCacheProvider } from "@formula/power-query/node";

const store = new MemoryCacheStore();
const cache = new CacheManager({ store });

const engine = new QueryEngine({ cache });

const query: Query = {
  id: "q1",
  name: "Query 1",
  source: { type: "range", range: { values: [["A"], [1]], hasHeaders: true } },
  steps: [],
  refreshPolicy: { type: "manual" },
};

const context: QueryExecutionContext = {};
void engine.executeQuery(query, context, { limit: 10 });

const mgr = new RefreshManager({ engine, concurrency: 1 });
mgr.registerQuery(query);
mgr.refresh(query.id);

const orchestrator = new RefreshOrchestrator({ engine, concurrency: 1 });
orchestrator.registerQuery(query);
orchestrator.refreshAll();

const op: QueryOperation = { type: "take", count: 1 };
void hashValue(op);

const source: QuerySource = { type: "csv", path: "/tmp/data.csv" };
void source;

void new FileConnector();
void new HttpConnector();
void new ODataConnector();
void new SharePointConnector();
void new SqlConnector();

void DataTable.fromGrid([["A"], [1]], { hasHeaders: true });

void parseCronExpression("* * * * *");
void randomId(8);
void httpScope({ url: "https://example.com/api" });
void oauth2Scope({ providerId: "example", scopesHash: "hash" });
void createWebCryptoCacheProvider({ keyVersion: 1, keyBytes: new Uint8Array(32) });
void createNodeCryptoCacheProvider({ keyVersion: 1, keyBytes: new Uint8Array(32) });

void new EncryptedCacheStore({ store, crypto: {} as any });
void MS_PER_DAY;
void new PqDecimal("1.23");
void new PqTime(0);
void new PqDuration(0);
void new PqDateTimeZone(new Date(), 0);
