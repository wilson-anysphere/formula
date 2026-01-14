import crypto from "node:crypto";
import http from "node:http";
import type { IncomingMessage } from "node:http";
import https from "node:https";
import { promises as fs, readFileSync } from "node:fs";
import { createRequire } from "node:module";
import type { AddressInfo } from "node:net";
import { monitorEventLoopDelay, type IntervalHistogram } from "node:perf_hooks";
import path from "node:path";
import type { Duplex } from "node:stream";

import type { Logger } from "pino";
import WebSocket, { WebSocketServer } from "ws";

// y-websocket ships its server utilities as a CommonJS file under bin/.
// eslint-disable-next-line @typescript-eslint/ban-ts-comment
// @ts-ignore
import ywsUtils from "y-websocket/bin/utils";

import type { SyncServerConfig } from "./config.js";
import {
  AuthError,
  authenticateRequest,
  extractToken,
  type AuthContext,
  type IntrospectCache,
} from "./auth.js";
import { acquireDataDirLock, type DataDirLockHandle } from "./dataDirLock.js";
import {
  LeveldbDocNameHashingLayer,
  persistedDocName as derivePersistedDocName,
} from "./leveldb-docname.js";
import {
  ConnectionTracker,
  SlidingWindowRateLimiter,
  TokenBucketRateLimiter,
} from "./limits.js";
import {
  FilePersistence,
  migrateLegacyPlaintextFilesToEncryptedFormat,
  type PersistenceOverloadScope,
} from "./persistence.js";
import { createEncryptedLevelAdapter } from "./leveldbEncryption.js";
import {
  DocConnectionTracker,
  LeveldbRetentionManager,
  type LeveldbPersistenceLike,
  type RetentionSweepResult,
} from "./retention.js";
import { requireLevelForYLeveldb } from "./leveldbLevel.js";
import { TombstoneStore, docKeyFromDocName } from "./tombstones.js";
import { Y } from "./yjs.js";
import { installYwsSecurity } from "./ywsSecurity.js";
import { createSyncServerMetrics, type WsConnectionRejectionReason } from "./metrics.js";
import {
  createSyncTokenIntrospectionClient,
  SyncTokenIntrospectionOverCapacityError,
  type SyncTokenIntrospectionClient,
} from "./introspection.js";
import { statusCodeForIntrospectionReason } from "./introspection-reasons.js";

const { setupWSConnection, setPersistence, getYDoc, docs: ywsDocs } = ywsUtils as {
  setupWSConnection: (
    conn: WebSocket,
    req: IncomingMessage,
    opts?: { gc?: boolean }
  ) => void;
  setPersistence: (persistence: unknown) => void;
  getYDoc: (docName: string, gc?: boolean) => any;
  docs: Map<string, any>;
};

function pickIp(req: IncomingMessage, trustProxy: boolean): string {
  if (trustProxy) {
    const forwarded = req.headers["x-forwarded-for"];
    const raw =
      typeof forwarded === "string"
        ? forwarded
        : Array.isArray(forwarded)
          ? forwarded[0]
          : null;
    if (raw && raw.length > 0) {
      const first = raw.split(",")[0]?.trim();
      if (first) {
        return first.length > MAX_CLIENT_IP_CHARS
          ? first.slice(0, MAX_CLIENT_IP_CHARS)
          : first;
      }
    }
  }
  const remote = req.socket.remoteAddress ?? "unknown";
  return remote.length > MAX_CLIENT_IP_CHARS
    ? remote.slice(0, MAX_CLIENT_IP_CHARS)
    : remote;
}

function timingSafeEqualStrings(a: string, b: string): boolean {
  const aBuf = Buffer.from(a);
  const bBuf = Buffer.from(b);
  if (aBuf.length !== bBuf.length) return false;
  return crypto.timingSafeEqual(aBuf, bBuf);
}

function rawPathnameFromUrl(requestUrl: string): string {
  const queryIndex = requestUrl.indexOf("?");
  const withoutQuery = queryIndex === -1 ? requestUrl : requestUrl.slice(0, queryIndex);

  // Typical HTTP request targets are "origin-form" (e.g. "/path"). Only treat
  // `scheme://...` as an absolute-form URL when the string does not start with
  // a slash so document ids like `/doc://name` don't get misclassified.
  if (withoutQuery.startsWith("/")) return withoutQuery;

  const schemeIndex = withoutQuery.indexOf("://");
  if (schemeIndex !== -1) {
    const pathIndex = withoutQuery.indexOf("/", schemeIndex + 3);
    return pathIndex === -1 ? "/" : withoutQuery.slice(pathIndex);
  }

  return withoutQuery;
}

function rawDataByteLength(raw: WebSocket.RawData): number {
  if (typeof raw === "string") return Buffer.byteLength(raw);
  if (Array.isArray(raw)) {
    return raw.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  }
  if (raw instanceof ArrayBuffer) return raw.byteLength;
  // Buffer is a Uint8Array
  if (raw instanceof Uint8Array) return raw.byteLength;
  return 0;
}

const MAX_CLIENT_IP_CHARS = 128;
const MAX_USER_AGENT_CHARS = 512;
const MAX_DOC_NAME_BYTES = 1024;

let eventLoopDelayHistogram: IntervalHistogram | null = null;
let eventLoopDelayHistogramUsers = 0;

function acquireEventLoopDelayHistogram(): IntervalHistogram | null {
  if (eventLoopDelayHistogram) {
    eventLoopDelayHistogramUsers += 1;
    return eventLoopDelayHistogram;
  }

  try {
    // `monitorEventLoopDelay()` was added in Node 11 and is guaranteed in our
    // supported runtime (Node >= 20), but guard so the server can still start
    // in older environments.
    if (typeof monitorEventLoopDelay !== "function") return null;
    eventLoopDelayHistogram = monitorEventLoopDelay({ resolution: 20 });
    eventLoopDelayHistogram.enable();
    eventLoopDelayHistogramUsers = 1;
    return eventLoopDelayHistogram;
  } catch {
    eventLoopDelayHistogram = null;
    eventLoopDelayHistogramUsers = 0;
    return null;
  }
}

function releaseEventLoopDelayHistogram(): void {
  if (eventLoopDelayHistogramUsers <= 0) return;
  eventLoopDelayHistogramUsers -= 1;
  if (eventLoopDelayHistogramUsers > 0) return;

  const histogram = eventLoopDelayHistogram;
  eventLoopDelayHistogram = null;

  if (!histogram) return;
  try {
    histogram.disable();
  } catch {
    // ignore
  }
  try {
    histogram.reset();
  } catch {
    // ignore
  }
}

function sendUpgradeRejection(
  socket: Duplex,
  statusCode: number,
  message: string
) {
  const statusText = http.STATUS_CODES[statusCode] ?? "Error";
  const body = message ?? statusText;

  const headers = [
    `HTTP/1.1 ${statusCode} ${statusText}`,
    "Connection: close",
    "Content-Type: text/plain; charset=utf-8",
    `Content-Length: ${Buffer.byteLength(body)}`,
    "",
    body,
  ].join("\r\n");

  try {
    socket.end(headers);
  } catch {
    socket.destroy();
  }
}

function sendJson(
  res: http.ServerResponse,
  statusCode: number,
  body: unknown
): void {
  res.writeHead(statusCode, { "content-type": "application/json" });
  res.end(JSON.stringify(body));
}

function sendText(
  res: http.ServerResponse,
  statusCode: number,
  body: string,
  contentType: string
): void {
  res.writeHead(statusCode, { "content-type": contentType });
  res.end(body);
}

async function countPersistedDocBlobsOnDisk(dir: string): Promise<number> {
  try {
    const entries = await fs.readdir(dir, { withFileTypes: true });
    return entries.filter((e) => e.isFile() && e.name.endsWith(".yjs")).length;
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "ENOENT") return 0;
    throw err;
  }
}

async function getDirectoryStats(dir: string): Promise<{
  fileCount: number;
  sizeBytes: number;
}> {
  try {
    const entries = await fs.readdir(dir, { withFileTypes: true });
    let fileCount = 0;
    let sizeBytes = 0;
    for (const entry of entries) {
      if (!entry.isFile()) continue;
      fileCount += 1;
      const st = await fs.stat(path.join(dir, entry.name));
      sizeBytes += st.size;
    }
    return { fileCount, sizeBytes };
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "ENOENT") return { fileCount: 0, sizeBytes: 0 };
    throw err;
  }
}

function toErrorMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  return typeof err === "string" ? err : JSON.stringify(err);
}

export type SyncServerHandle = {
  start: () => Promise<{ port: number }>;
  stop: () => Promise<void>;
  getHttpUrl: () => string;
  getWsUrl: () => string;
};

type LeveldbPersistence = LeveldbPersistenceLike & {
  getYDoc: (docName: string) => Promise<any>;
  storeUpdate: (docName: string, update: Uint8Array) => Promise<void>;
  flushDocument: (docName: string) => Promise<void>;
  getMetas?: (docName: string) => Promise<Map<string, unknown>>;
  destroy: () => Promise<void>;
};

export type SyncServerCreateOptions = {
  createLeveldbPersistence?: (
    location: string,
    opts: { encryption: SyncServerConfig["persistence"]["encryption"] }
  ) => LeveldbPersistence;
};

export function createSyncServer(
  config: SyncServerConfig,
  logger: Logger,
  { createLeveldbPersistence }: SyncServerCreateOptions = {}
) {
  const nodeEnv = process.env.NODE_ENV ?? "development";
  const parseCommaList = (value: string | undefined): string[] | null => {
    if (value === undefined) return null;
    return value
      .split(",")
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);
  };
  const reservedRootGuardEnabled = (() => {
    const raw = process.env.SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED;
    if (raw === undefined) {
      // Default to enabled in production, but allow tests/dev environments to opt out
      // so existing workflows that store versioning/branching metadata in the Y.Doc
      // remain compatible.
      return nodeEnv === "production";
    }
    const normalized = raw.trim().toLowerCase();
    return normalized === "1" || normalized === "true";
  })();
  const reservedRootNames =
    parseCommaList(process.env.SYNC_SERVER_RESERVED_ROOT_NAMES) ?? [
      "versions",
      "versionsMeta",
    ];
  const reservedRootPrefixes =
    parseCommaList(process.env.SYNC_SERVER_RESERVED_ROOT_PREFIXES) ?? ["branching:"];
  const reservedRootGuard = {
    enabled: reservedRootGuardEnabled,
    reservedRootNames,
    reservedRootPrefixes,
  };

  const allowedOrigins = (() => {
    const list = config.allowedOrigins;
    if (!list || list.length === 0) return null;
    return new Set(list);
  })();

  const connectionTracker = new ConnectionTracker(
    config.limits.maxConnections,
    config.limits.maxConnectionsPerIp
  );
  const docConnectionTracker = new DocConnectionTracker();
  const connectionAttemptLimiter = new TokenBucketRateLimiter(
    config.limits.maxConnAttemptsPerWindow,
    config.limits.connAttemptWindowMs
  );
  const docMessageLimiter = new SlidingWindowRateLimiter(
    config.limits.maxMessagesPerDocWindow,
    config.limits.docMessageWindowMs
  );
  const ipMessageLimiter = new SlidingWindowRateLimiter(
    config.limits.maxMessagesPerIpWindow,
    config.limits.ipMessageWindowMs
  );

  const metrics = createSyncServerMetrics();
  let draining = false;
  let stopInFlight: Promise<void> | null = null;
  let drainWaiter: (() => void) | null = null;
  let eventLoopDelay: IntervalHistogram | null = null;
  let processMetricsTimer: NodeJS.Timeout | null = null;
  const updateProcessMetrics = () => {
    const mem = process.memoryUsage();
    metrics.processResidentMemoryBytes.set(mem.rss);
    metrics.processHeapUsedBytes.set(mem.heapUsed);
    metrics.processHeapTotalBytes.set(mem.heapTotal);
    // monitorEventLoopDelay reports nanoseconds. Store p99 in milliseconds.
    if (eventLoopDelay) {
      metrics.eventLoopDelayMs.set(eventLoopDelay.percentile(99) / 1e6);
      // Reset so each collection is roughly per-interval, not cumulative.
      eventLoopDelay.reset();
    }
  };
  const getEventLoopDelayMs = (): number => {
    if (!eventLoopDelay) return 0;
    const p99Ns = eventLoopDelay.percentile(99);
    // Values are reported in nanoseconds.
    return Number.isFinite(p99Ns) ? p99Ns / 1e6 : 0;
  };

  const closeCodeLabel = (code: number): string => {
    switch (code) {
      case 1000:
      case 1003:
      case 1006:
      case 1008:
      case 1009:
      case 1011:
      case 1013:
        return String(code);
      default:
        return "other";
    }
  };

  const recordUpgradeRejection = (
    reason: WsConnectionRejectionReason
  ) => {
    metrics.wsConnectionsRejectedTotal.inc({ reason });
  };

  const recordLeveldbRetentionSweep = (result: { purged: number; errors: unknown[] }) => {
    metrics.retentionSweepsTotal.inc({ sweep: "leveldb" });
    if (result.purged > 0) {
      metrics.retentionDocsPurgedTotal.inc({ sweep: "leveldb" }, result.purged);
    } else {
      metrics.retentionDocsPurgedTotal.inc({ sweep: "leveldb" }, 0);
    }
    metrics.retentionSweepErrorsTotal.inc({ sweep: "leveldb" }, result.errors.length);
  };

  const recordTombstoneSweep = (result: {
    docBlobsDeleted: number;
    errors: unknown[];
  }) => {
    metrics.retentionSweepsTotal.inc({ sweep: "tombstone" });
    if (result.docBlobsDeleted > 0) {
      metrics.retentionDocsPurgedTotal.inc(
        { sweep: "tombstone" },
        result.docBlobsDeleted
      );
    } else {
      metrics.retentionDocsPurgedTotal.inc({ sweep: "tombstone" }, 0);
    }
    metrics.retentionSweepErrorsTotal.inc(
      { sweep: "tombstone" },
      result.errors.length
    );
  };

  const activeSocketsByDoc = new Map<string, Set<WebSocket>>();

  const closeActiveSocketsForDoc = (docName: string, code: number, reason: string) => {
    const sockets = activeSocketsByDoc.get(docName);
    if (!sockets || sockets.size === 0) return;
    for (const ws of Array.from(sockets)) {
      try {
        ws.close(code, reason);
      } catch {
        // Ignore failures (socket may already be closing).
      }
    }
  };

  const triggerPersistenceOverload = (docName: string, scope: PersistenceOverloadScope) => {
    metrics.persistenceOverloadTotal.inc({ scope });
    closeActiveSocketsForDoc(docName, 1013, "persistence overloaded");
  };

  const cleanupOrphanedYwsDoc = (docName: string) => {
    try {
      const doc = ywsDocs.get(docName);
      if (!doc) return;
      const conns = (doc as any)?.conns;
      const connCount = conns && typeof conns.size === "number" ? conns.size : 0;
      if (connCount > 0) return;
      ywsDocs.delete(docName);
      if (typeof (doc as any)?.destroy === "function") {
        (doc as any).destroy();
      }
    } catch (err) {
      logger.warn({ err, docName }, "yws_doc_cleanup_failed");
    }
  };

  const tombstones = new TombstoneStore(config.dataDir, logger);
  const shouldPersist = (docName: string) => !tombstones.has(docKeyFromDocName(docName));

  const syncTokenIntrospection: SyncTokenIntrospectionClient | null = config.introspection
    ? createSyncTokenIntrospectionClient({
        ...config.introspection,
        metrics,
        maxTokenBytes: config.limits.maxTokenBytes,
      })
    : null;

  type TombstoneSweepResult = {
    expiredTombstonesRemoved: number;
    tombstonesProcessed: number;
    docBlobsDeleted: number;
    errors: Array<{ docKey: string; error: string }>;
  };

  let tombstoneSweepTimer: NodeJS.Timeout | null = null;
  let tombstoneSweepInFlight: Promise<TombstoneSweepResult> | null = null;
  let tombstoneSweepCursor = 0;
  let leveldbTombstoneSweep:
    | {
        getAllDocNames: () => Promise<string[]>;
        clearDocument: (docName: string) => Promise<void>;
      }
    | null = null;

  let dataDirLock: DataDirLockHandle | null = null;
  let persistenceInitialized = false;
  let persistenceCleanup: (() => Promise<void>) | null = null;
  let persistenceBackend: "file" | "leveldb" | null = null;
  let clearPersistedDocument: ((docName: string) => Promise<void>) | null = null;
  let waitForDocLoaded: ((docName: string) => Promise<void>) | null = null;

  let retentionManager: LeveldbRetentionManager | null = null;
  let retentionSweepTimer: NodeJS.Timeout | null = null;
  let retentionSweepInFlight: Promise<RetentionSweepResult> | null = null;

  const leveldbDocNameHashingEnabled =
    config.persistence.backend === "leveldb" && config.persistence.leveldbDocNameHashing;
  const persistedDocNameForLeveldb = (docName: string) =>
    derivePersistedDocName(docName, leveldbDocNameHashingEnabled);

  const maybeStartRetentionSweeper = () => {
    if (
      retentionSweepTimer ||
      !retentionManager ||
      config.retention.ttlMs <= 0 ||
      config.retention.sweepIntervalMs <= 0
    ) {
      return;
    }

    retentionSweepTimer = setInterval(() => {
      void triggerRetentionSweep();
    }, config.retention.sweepIntervalMs);
    retentionSweepTimer.unref();
  };

  const triggerRetentionSweep = async (): Promise<RetentionSweepResult> => {
    if (!retentionManager) {
      throw new Error("Retention is not initialized");
    }

    if (retentionSweepInFlight) return await retentionSweepInFlight;

    retentionSweepInFlight = retentionManager
      .sweep()
      .then((result) => {
        recordLeveldbRetentionSweep(result);
        logger.info(
          {
            ...result,
            ttlMs: config.retention.ttlMs,
            intervalMs: config.retention.sweepIntervalMs,
          },
          "retention_sweep_completed"
        );
        return result;
      })
      .catch((err) => {
        metrics.retentionSweepsTotal.inc({ sweep: "leveldb" });
        metrics.retentionSweepErrorsTotal.inc({ sweep: "leveldb" }, 1);
        logger.error({ err }, "retention_sweep_failed");
        throw err;
      })
      .finally(() => {
        retentionSweepInFlight = null;
      });

    return await retentionSweepInFlight;
  };

  const initPersistence = async () => {
    if (persistenceInitialized) return;

    const nodeEnv = process.env.NODE_ENV ?? "development";
    await tombstones.init();

    if (config.persistence.backend === "file") {
      if (config.persistence.encryption.mode === "keyring") {
        await migrateLegacyPlaintextFilesToEncryptedFormat({
          dir: config.dataDir,
          logger,
          keyRing: config.persistence.encryption.keyRing,
        });
      }
  
      persistenceInitialized = true;
      persistenceCleanup = null;
      retentionManager = null;
      leveldbTombstoneSweep = null;
      persistenceBackend = "file";
      const persistence = new FilePersistence(
        config.dataDir,
        logger,
        config.persistence.compactAfterUpdates,
        config.persistence.encryption,
        shouldPersist,
        {
          maxQueueDepthPerDoc: config.persistence.maxQueueDepthPerDoc ?? 0,
          maxQueueDepthTotal: config.persistence.maxQueueDepthTotal ?? 0,
          onOverload: (docName, scope) => triggerPersistenceOverload(docName, scope),
        }
      );
      persistenceCleanup = () => persistence.flush();
      waitForDocLoaded = (docName: string) => persistence.waitForLoaded(docName);
      setPersistence(persistence);
      metrics.setPersistenceInfo({
        backend: "file",
        encryptionEnabled: config.persistence.encryption.mode !== "off",
      });
      clearPersistedDocument = (docName: string) => persistence.clearDocument(docName);
      logger.info(
        {
          dir: config.dataDir,
          encryption: config.persistence.encryption.mode,
        },
        "persistence_file_enabled"
      );
      return;
    }

    if (createLeveldbPersistence) {
      const ldb = createLeveldbPersistence(config.dataDir, {
        encryption: config.persistence.encryption,
      });
      const hashingEnabled = leveldbDocNameHashingEnabled;
      const hashedLdb = new LeveldbDocNameHashingLayer(ldb, hashingEnabled);
      const docsNeedingMigration = new Set<string>();
      const emptyStateVector = Y.encodeStateVector(new Y.Doc());
      const isEmptyYDoc = (doc: unknown) => {
        const sv = Y.encodeStateVector(doc as any);
        if (sv.byteLength !== emptyStateVector.byteLength) return false;
        for (let i = 0; i < sv.byteLength; i += 1) {
          if (sv[i] !== emptyStateVector[i]) return false;
        }
        return true;
      };

      persistenceBackend = "leveldb";

      const queues = new Map<string, Promise<unknown>>();
      const pendingCountsByDoc = new Map<string, number>();
      let pendingTotal = 0;
      const disabledDocs = new Set<string>();

      const maxQueueDepthPerDoc = config.persistence.maxQueueDepthPerDoc ?? 0;
      const maxQueueDepthTotal = config.persistence.maxQueueDepthTotal ?? 0;

      const enqueue = <T>(docName: string, task: () => Promise<T>) => {
        const pendingForDoc = pendingCountsByDoc.get(docName) ?? 0;
        pendingCountsByDoc.set(docName, pendingForDoc + 1);
        pendingTotal += 1;

        const prev = queues.get(docName) ?? Promise.resolve();
        const next = prev
          .catch(() => {
            // Keep the queue alive even if a previous write failed.
          })
          .then(task);
        const nextForQueue = next as Promise<unknown>;
        queues.set(docName, nextForQueue);
        void nextForQueue
          .finally(() => {
            pendingTotal = Math.max(0, pendingTotal - 1);
            const remaining = (pendingCountsByDoc.get(docName) ?? 1) - 1;
            if (remaining <= 0) pendingCountsByDoc.delete(docName);
            else pendingCountsByDoc.set(docName, remaining);
            if (queues.get(docName) === nextForQueue) queues.delete(docName);
          })
          .catch(() => {
            // Best-effort: the returned `next` promise is handled by callers; avoid an unhandled
            // rejection from this internal `.finally` bookkeeping chain.
          });
        return next;
      };
      const docLoadPromises = new Map<string, Promise<void>>();

      leveldbTombstoneSweep = {
        getAllDocNames: () => ldb.getAllDocNames(),
        clearDocument: (docName: string) =>
          enqueue(docName, () => ldb.clearDocument(docName)),
      };

      const retentionProvider: LeveldbPersistenceLike = {
        getAllDocNames: () => ldb.getAllDocNames(),
        clearDocument: (docName: string) =>
          enqueue(docName, () => ldb.clearDocument(docName)),
        getMeta: (docName: string, metaKey: string) => ldb.getMeta(docName, metaKey),
        setMeta: (docName: string, metaKey: string, value: unknown) =>
          ldb.setMeta(docName, metaKey, value),
      };

      retentionManager = new LeveldbRetentionManager(
        retentionProvider,
        docConnectionTracker,
        logger,
        config.retention.ttlMs
      );

      persistenceInitialized = true;
      persistenceCleanup = async () => {
        await Promise.allSettled([...queues.values()]);
        await ldb.destroy();
      };
      clearPersistedDocument = (docName: string) =>
        enqueue(persistedDocNameForLeveldb(docName), async () => {
          await hashedLdb.clearDocument(docName);
          if (hashingEnabled) {
            // Legacy namespace cleanup (raw docName keys).
            await ldb.clearDocument(docName);
          }
        });

      setPersistence({
        provider: hashedLdb,
        bindState: (docName: string, ydoc: any) => {
          // Important: `y-websocket` does not await `bindState()`. Track the
          // async loading promise so websocket upgrades can wait for the doc to
          // be fully hydrated before allowing the initial sync handshake.
          const loadPromise = (async () => {
            const persistenceOrigin = "persistence:leveldb";
            const retentionDocName = persistedDocNameForLeveldb(docName);
 
            ydoc.on("update", (update: Uint8Array, origin: unknown) => {
              if (origin === persistenceOrigin) return;
              if (!shouldPersist(docName)) return;
              if (retentionManager?.isPurging(retentionDocName)) return;
              if (disabledDocs.has(docName)) return;
 
              const pendingForDoc = pendingCountsByDoc.get(retentionDocName) ?? 0;
              if (maxQueueDepthPerDoc > 0 && pendingForDoc >= maxQueueDepthPerDoc) {
                disabledDocs.add(docName);
                triggerPersistenceOverload(docName, "doc");
                return;
              }
              if (maxQueueDepthTotal > 0 && pendingTotal >= maxQueueDepthTotal) {
                disabledDocs.add(docName);
                triggerPersistenceOverload(docName, "total");
                return;
              }
 
              void enqueue(retentionDocName, async () => {
                await hashedLdb.storeUpdate(docName, update);
              });
              void retentionManager?.markSeen(retentionDocName);
            });
            ydoc.on("destroy", () => {
              disabledDocs.delete(docName);
            });
 
            if (shouldPersist(docName)) {
              void retentionManager?.markSeen(retentionDocName, { force: true });

              const persistedYdoc = await hashedLdb.getYDoc(docName);
              Y.applyUpdate(
                ydoc,
                Y.encodeStateAsUpdate(persistedYdoc),
                persistenceOrigin
              );

              if (hashingEnabled) {
                const legacyYdoc = await ldb.getYDoc(docName);
                const legacyMetas =
                  typeof ldb.getMetas === "function"
                    ? await ldb.getMetas(docName)
                    : null;

                let needsMigration = false;

                if (!isEmptyYDoc(legacyYdoc)) {
                  needsMigration = true;
                  Y.applyUpdate(
                    ydoc,
                    Y.encodeStateAsUpdate(legacyYdoc),
                    persistenceOrigin
                  );
                }

                if (legacyMetas && legacyMetas.size > 0) {
                  needsMigration = true;
                  await enqueue(retentionDocName, async () => {
                    for (const [metaKey, value] of legacyMetas.entries()) {
                      await hashedLdb.setMeta(docName, metaKey, value);
                    }
                  });
                }

                if (needsMigration) {
                  docsNeedingMigration.add(docName);
                }
              }
            }
          })();

          docLoadPromises.set(docName, loadPromise);
          loadPromise.catch(() => {});
          ydoc.on("destroy", () => {
            if (docLoadPromises.get(docName) === loadPromise) {
              docLoadPromises.delete(docName);
            }
          });
          return loadPromise;
        },
        writeState: async (docName: string, ydoc: any) => {
          if (!shouldPersist(docName)) return;
          const retentionDocName = persistedDocNameForLeveldb(docName);
          if (retentionManager?.isPurging(retentionDocName)) return;
          if (disabledDocs.has(docName)) return;

          if (hashingEnabled && docsNeedingMigration.has(docName)) {
            if (ydoc) {
              await enqueue(retentionDocName, () =>
                hashedLdb.storeUpdate(docName, Y.encodeStateAsUpdate(ydoc))
              );
            }

            await enqueue(retentionDocName, () => hashedLdb.flushDocument(docName));
            await enqueue(retentionDocName, () => ldb.clearDocument(docName));
            docsNeedingMigration.delete(docName);
            void retentionManager?.markFlushed(retentionDocName);
            return;
          }

          await enqueue(retentionDocName, () => hashedLdb.flushDocument(docName));
          void retentionManager?.markFlushed(retentionDocName);
        },
      });
      waitForDocLoaded = (docName: string) => docLoadPromises.get(docName) ?? Promise.resolve();

      metrics.setPersistenceInfo({
        backend: "leveldb",
        encryptionEnabled: config.persistence.encryption.mode !== "off",
      });
      maybeStartRetentionSweeper();
      logger.info(
        {
          dir: config.dataDir,
          docNameHashing: hashingEnabled,
          encryption: config.persistence.encryption.mode,
          strict:
            config.persistence.encryption.mode === "keyring"
              ? config.persistence.encryption.strict
              : undefined,
        },
        "persistence_leveldb_enabled"
      );
      return;
    }

    const require = createRequire(import.meta.url);
    let LeveldbPersistenceCtor: new (
      location: string,
      opts?: any
    ) => LeveldbPersistence;
    try {
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      ({ LeveldbPersistence: LeveldbPersistenceCtor } = require("y-leveldb") as {
        LeveldbPersistence: new (location: string, opts?: any) => LeveldbPersistence;
      });
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code === "MODULE_NOT_FOUND") {
        if (nodeEnv === "production") {
          throw new Error(
            "y-leveldb is required for LevelDB persistence. Install y-leveldb or set SYNC_SERVER_PERSISTENCE_BACKEND=file."
          );
        }

        logger.warn(
          { err },
          "y-leveldb_not_installed_falling_back_to_file_persistence"
        );

        if (config.persistence.encryption.mode === "keyring") {
          await migrateLegacyPlaintextFilesToEncryptedFormat({
            dir: config.dataDir,
            logger,
            keyRing: config.persistence.encryption.keyRing,
          });
        }

        persistenceInitialized = true;
        persistenceCleanup = null;
        retentionManager = null;
        leveldbTombstoneSweep = null;
        persistenceBackend = "file";
        const persistence = new FilePersistence(
          config.dataDir,
          logger,
          config.persistence.compactAfterUpdates,
          config.persistence.encryption,
          shouldPersist,
          {
            maxQueueDepthPerDoc: config.persistence.maxQueueDepthPerDoc ?? 0,
            maxQueueDepthTotal: config.persistence.maxQueueDepthTotal ?? 0,
            onOverload: (docName, scope) => triggerPersistenceOverload(docName, scope),
          }
        );
        waitForDocLoaded = (docName: string) => persistence.waitForLoaded(docName);
        setPersistence(persistence);
        metrics.setPersistenceInfo({
          backend: "file",
          encryptionEnabled: config.persistence.encryption.mode !== "off",
        });
        clearPersistedDocument = (docName: string) => persistence.clearDocument(docName);
        return;
      }

      throw err;
    }

    await fs.mkdir(config.dataDir, { recursive: true });

    const isLockError = (err: unknown) => {
      const msg =
        err instanceof Error ? err.message : typeof err === "string" ? err : "";
      return msg.toLowerCase().includes("lock");
    };

    const maxAttempts = 5;
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      try {
        const levelAdapter =
          config.persistence.encryption.mode === "keyring"
            ? createEncryptedLevelAdapter({
                keyRing: config.persistence.encryption.keyRing,
                strict: config.persistence.encryption.strict,
              })(requireLevelForYLeveldb())
            : undefined;

        const ldb = new LeveldbPersistenceCtor(
          config.dataDir,
          levelAdapter ? { level: levelAdapter } : undefined
        );
        const hashingEnabled = leveldbDocNameHashingEnabled;
        const hashedLdb = new LeveldbDocNameHashingLayer(ldb, hashingEnabled);
        const docsNeedingMigration = new Set<string>();
        const emptyStateVector = Y.encodeStateVector(new Y.Doc());
        const isEmptyYDoc = (doc: unknown) => {
          const sv = Y.encodeStateVector(doc as any);
          if (sv.byteLength !== emptyStateVector.byteLength) return false;
          for (let i = 0; i < sv.byteLength; i += 1) {
            if (sv[i] !== emptyStateVector[i]) return false;
          }
          return true;
        };

        persistenceBackend = "leveldb";

        const queues = new Map<string, Promise<unknown>>();
        const pendingCountsByDoc = new Map<string, number>();
        let pendingTotal = 0;
        const disabledDocs = new Set<string>();

        const maxQueueDepthPerDoc = config.persistence.maxQueueDepthPerDoc ?? 0;
        const maxQueueDepthTotal = config.persistence.maxQueueDepthTotal ?? 0;

        const enqueue = <T>(docName: string, task: () => Promise<T>) => {
          const pendingForDoc = pendingCountsByDoc.get(docName) ?? 0;
          pendingCountsByDoc.set(docName, pendingForDoc + 1);
          pendingTotal += 1;

          const prev = queues.get(docName) ?? Promise.resolve();
          const next = prev
            .catch(() => {
              // Keep the queue alive even if a previous write failed.
            })
            .then(task);
          const nextForQueue = next as Promise<unknown>;
          queues.set(docName, nextForQueue);
          void nextForQueue
            .finally(() => {
              pendingTotal = Math.max(0, pendingTotal - 1);
              const remaining = (pendingCountsByDoc.get(docName) ?? 1) - 1;
              if (remaining <= 0) pendingCountsByDoc.delete(docName);
              else pendingCountsByDoc.set(docName, remaining);
              if (queues.get(docName) === nextForQueue) queues.delete(docName);
            })
            .catch(() => {
              // Best-effort: the returned `next` promise is handled by callers; avoid an unhandled
              // rejection from this internal `.finally` bookkeeping chain.
            });
          return next;
        };
        const docLoadPromises = new Map<string, Promise<void>>();

        leveldbTombstoneSweep = {
          getAllDocNames: () => ldb.getAllDocNames(),
          clearDocument: (docName: string) =>
            enqueue(docName, () => ldb.clearDocument(docName)),
        };

        const retentionProvider: LeveldbPersistenceLike = {
          getAllDocNames: () => ldb.getAllDocNames(),
          clearDocument: (docName: string) =>
            enqueue(docName, () => ldb.clearDocument(docName)),
          getMeta: (docName: string, metaKey: string) =>
            ldb.getMeta(docName, metaKey),
          setMeta: (docName: string, metaKey: string, value: unknown) =>
            ldb.setMeta(docName, metaKey, value),
        };

        retentionManager = new LeveldbRetentionManager(
          retentionProvider,
          docConnectionTracker,
          logger,
          config.retention.ttlMs
        );

        persistenceInitialized = true;
        persistenceCleanup = async () => {
          await Promise.allSettled([...queues.values()]);
          await ldb.destroy();
        };
        clearPersistedDocument = (docName: string) =>
          enqueue(persistedDocNameForLeveldb(docName), async () => {
            await hashedLdb.clearDocument(docName);
            if (hashingEnabled) {
              await ldb.clearDocument(docName);
            }
          });

        setPersistence({
          provider: hashedLdb,
          bindState: (docName: string, ydoc: any) => {
            // Important: `y-websocket` does not await `bindState()`. Track the
            // async loading promise so websocket upgrades can wait for the doc to
            // be fully hydrated before allowing the initial sync handshake.
            const loadPromise = (async () => {
              // Attach the update listener first so we don't miss early client updates.
              const persistenceOrigin = "persistence:leveldb";
              const retentionDocName = persistedDocNameForLeveldb(docName);
 
              ydoc.on("update", (update: Uint8Array, origin: unknown) => {
                if (origin === persistenceOrigin) return;
                if (!shouldPersist(docName)) return;
                if (retentionManager?.isPurging(retentionDocName)) return;
                if (disabledDocs.has(docName)) return;
 
                const pendingForDoc = pendingCountsByDoc.get(retentionDocName) ?? 0;
                if (maxQueueDepthPerDoc > 0 && pendingForDoc >= maxQueueDepthPerDoc) {
                  disabledDocs.add(docName);
                  triggerPersistenceOverload(docName, "doc");
                  return;
                }
                if (maxQueueDepthTotal > 0 && pendingTotal >= maxQueueDepthTotal) {
                  disabledDocs.add(docName);
                  triggerPersistenceOverload(docName, "total");
                  return;
                }
                void enqueue(retentionDocName, async () => {
                  await hashedLdb.storeUpdate(docName, update);
                });
                void retentionManager?.markSeen(retentionDocName);
              });
              ydoc.on("destroy", () => {
                disabledDocs.delete(docName);
              });

              if (shouldPersist(docName)) {
                void retentionManager?.markSeen(retentionDocName, { force: true });
                const persistedYdoc = await hashedLdb.getYDoc(docName);
                Y.applyUpdate(
                  ydoc,
                  Y.encodeStateAsUpdate(persistedYdoc),
                  persistenceOrigin
                );

                if (hashingEnabled) {
                  const legacyYdoc = await ldb.getYDoc(docName);
                  const legacyMetas =
                    typeof ldb.getMetas === "function"
                      ? await ldb.getMetas(docName)
                      : null;

                  let needsMigration = false;

                  if (!isEmptyYDoc(legacyYdoc)) {
                    needsMigration = true;
                    Y.applyUpdate(
                      ydoc,
                      Y.encodeStateAsUpdate(legacyYdoc),
                      persistenceOrigin
                    );
                  }

                  if (legacyMetas && legacyMetas.size > 0) {
                    needsMigration = true;
                    await enqueue(retentionDocName, async () => {
                      for (const [metaKey, value] of legacyMetas.entries()) {
                        await hashedLdb.setMeta(docName, metaKey, value);
                      }
                    });
                  }

                  if (needsMigration) {
                    docsNeedingMigration.add(docName);
                  }
                }
              }
            })();

            docLoadPromises.set(docName, loadPromise);
            loadPromise.catch(() => {});
            ydoc.on("destroy", () => {
              if (docLoadPromises.get(docName) === loadPromise) {
                docLoadPromises.delete(docName);
              }
            });
            return loadPromise;
          },
          writeState: async (docName: string, ydoc: any) => {
            // Compact updates on last client disconnect to keep DB size bounded.
            if (!shouldPersist(docName)) return;
            const retentionDocName = persistedDocNameForLeveldb(docName);
            if (retentionManager?.isPurging(retentionDocName)) return;
            if (disabledDocs.has(docName)) return;

            if (hashingEnabled && docsNeedingMigration.has(docName)) {
              if (ydoc) {
                await enqueue(retentionDocName, () =>
                  hashedLdb.storeUpdate(docName, Y.encodeStateAsUpdate(ydoc))
                );
              }

              await enqueue(retentionDocName, () => hashedLdb.flushDocument(docName));
              await enqueue(retentionDocName, () => ldb.clearDocument(docName));
              docsNeedingMigration.delete(docName);
              void retentionManager?.markFlushed(retentionDocName);
              return;
            }

            await enqueue(retentionDocName, () => hashedLdb.flushDocument(docName));
            void retentionManager?.markFlushed(retentionDocName);
          },
        });
        waitForDocLoaded = (docName: string) => docLoadPromises.get(docName) ?? Promise.resolve();

        metrics.setPersistenceInfo({
          backend: "leveldb",
          encryptionEnabled: config.persistence.encryption.mode !== "off",
        });
        maybeStartRetentionSweeper();
        logger.info(
          {
            dir: config.dataDir,
            docNameHashing: hashingEnabled,
            encryption: config.persistence.encryption.mode,
            strict:
              config.persistence.encryption.mode === "keyring"
                ? config.persistence.encryption.strict
                : undefined,
          },
          "persistence_leveldb_enabled"
        );
        return;
      } catch (err) {
        if (attempt < maxAttempts && isLockError(err)) {
          const delayMs = 50 * Math.pow(2, attempt - 1);
          logger.warn({ err, attempt, delayMs }, "leveldb_locked_retrying_open");
          await new Promise((r) => setTimeout(r, delayMs));
          continue;
        }
        throw err;
      }
    }
  };

  const sweepTombstonesOnce = async (): Promise<TombstoneSweepResult> => {
    await initPersistence();

    const { expiredDocKeys } = await tombstones.sweepExpired(
      config.retention.tombstoneTtlMs
    );

    const errors: Array<{ docKey: string; error: string }> = [];

    const SWEEP_DOC_LIMIT = 100;

    if (persistenceBackend === "file") {
      const docKeys = tombstones
        .entries()
        .map(([docKey]) => docKey)
        .sort((a, b) => a.localeCompare(b));

      const limit = Math.min(SWEEP_DOC_LIMIT, docKeys.length);

      let processed = 0;
      let deleted = 0;

      if (limit > 0) {
        const start = tombstoneSweepCursor % docKeys.length;
        for (let i = 0; i < limit; i += 1) {
          const docKey = docKeys[(start + i) % docKeys.length]!;
          processed += 1;
          try {
            await fs.rm(path.join(config.dataDir, `${docKey}.yjs`), { force: true });
            deleted += 1;
          } catch (err) {
            errors.push({ docKey, error: toErrorMessage(err) });
          }
        }
        tombstoneSweepCursor = (start + processed) % docKeys.length;
      }

      return {
        expiredTombstonesRemoved: expiredDocKeys.length,
        tombstonesProcessed: processed,
        docBlobsDeleted: deleted,
        errors,
      };
    }

    if (persistenceBackend === "leveldb") {
      const leveldb = leveldbTombstoneSweep;
      if (!leveldb) {
        tombstoneSweepCursor = 0;
        return {
          expiredTombstonesRemoved: expiredDocKeys.length,
          tombstonesProcessed: 0,
          docBlobsDeleted: 0,
          errors,
        };
      }

      const docNames = (await leveldb.getAllDocNames()).sort((a, b) =>
        a.localeCompare(b)
      );

      const limit = Math.min(SWEEP_DOC_LIMIT, docNames.length);

      let processed = 0;
      let deleted = 0;

      if (limit > 0) {
        const start = tombstoneSweepCursor % docNames.length;
        for (let i = 0; i < limit; i += 1) {
          const persistedName = docNames[(start + i) % docNames.length]!;
          processed += 1;

          const docKey = tombstones.has(persistedName)
            ? persistedName
            : docKeyFromDocName(persistedName);

          if (!tombstones.has(docKey)) continue;

          try {
            await leveldb.clearDocument(persistedName);
            deleted += 1;
          } catch (err) {
            errors.push({ docKey, error: toErrorMessage(err) });
          }
        }

        tombstoneSweepCursor = (start + processed) % docNames.length;
      } else {
        tombstoneSweepCursor = 0;
      }

      return {
        expiredTombstonesRemoved: expiredDocKeys.length,
        tombstonesProcessed: processed,
        docBlobsDeleted: deleted,
        errors,
      };
    }

    tombstoneSweepCursor = 0;
    return {
      expiredTombstonesRemoved: expiredDocKeys.length,
      tombstonesProcessed: 0,
      docBlobsDeleted: 0,
      errors,
    };
  };

  const triggerTombstoneSweep = async (): Promise<TombstoneSweepResult> => {
    if (tombstoneSweepInFlight) return await tombstoneSweepInFlight;
    tombstoneSweepInFlight = sweepTombstonesOnce()
      .then((result) => {
        recordTombstoneSweep(result);
        return result;
      })
      .catch((err) => {
        metrics.retentionSweepsTotal.inc({ sweep: "tombstone" });
        metrics.retentionDocsPurgedTotal.inc({ sweep: "tombstone" }, 0);
        metrics.retentionSweepErrorsTotal.inc({ sweep: "tombstone" }, 1);
        throw err;
      })
      .finally(() => {
        tombstoneSweepInFlight = null;
      });
    return await tombstoneSweepInFlight;
  };

  const wss = new WebSocketServer({
    noServer: true,
    maxPayload:
      config.limits.maxMessageBytes > 0 ? config.limits.maxMessageBytes : undefined,
    // Explicitly disable compression (defense-in-depth against compression bombs).
    perMessageDeflate: false,
  });

  const introspectCache: IntrospectCache | null =
    config.auth.mode === "introspect" ? new Map() : null;

  const handler: http.RequestListener = (req, res) => {
    void (async () => {
      // This server does not accept request bodies. Drain and discard any bytes
      // to avoid holding sockets open (or buffering data) under hostile clients.
      req.resume();

      if (!req.url) {
        res.writeHead(400).end();
        return;
      }

      const pathname = rawPathnameFromUrl(req.url);

      if (req.method === "GET" && pathname === "/metrics") {
        if (!config.metrics.public) {
          sendText(res, 404, "not_found", "text/plain; charset=utf-8");
          return;
        }
        res.setHeader("cache-control", "no-store");
        const body = await metrics.metricsText();
        sendText(res, 200, body, metrics.registry.contentType);
        return;
      }

      if (req.method === "GET" && pathname === "/readyz") {
        res.setHeader("cache-control", "no-store");
        if (draining) {
          sendJson(res, 503, { status: "not_ready", reason: "draining" });
          return;
        }

        if (!dataDirLock) {
          sendJson(res, 503, {
            status: "not_ready",
            reason: config.disableDataDirLock
              ? "data_dir_lock_disabled"
              : "data_dir_lock_not_acquired",
          });
          return;
        }

        const ready = persistenceInitialized && tombstones.isInitialized();
        sendJson(res, ready ? 200 : 503, { status: ready ? "ready" : "not_ready" });
        return;
      }

      if (req.method === "GET" && pathname === "/healthz") {
        const snapshot = connectionTracker.snapshot();
        const mem = process.memoryUsage();
        res.setHeader("cache-control", "no-store");
        sendJson(res, 200, {
          status: "ok",
          uptimeSec: Math.round(process.uptime()),
          rssBytes: mem.rss,
          heapUsedBytes: mem.heapUsed,
          heapTotalBytes: mem.heapTotal,
          eventLoopDelayMs: getEventLoopDelayMs(),
          draining,
          connections: snapshot,
          backend: persistenceBackend ?? config.persistence.backend,
          encryptionEnabled: config.persistence.encryption.mode !== "off",
          tombstonesCount: tombstones.count(),
        });
        return;
      }

      if (pathname.startsWith("/internal/")) {
        res.setHeader("cache-control", "no-store");

        if (!config.internalAdminToken) {
          sendJson(res, 404, { error: "not_found" });
          return;
        }

        const header = req.headers["x-internal-admin-token"];
        const provided =
          typeof header === "string"
            ? header
            : Array.isArray(header)
              ? header[0]
              : undefined;
        if (
          typeof provided !== "string" ||
          !timingSafeEqualStrings(provided, config.internalAdminToken)
        ) {
          sendJson(res, 403, { error: "forbidden" });
          return;
        }

        if (req.method === "GET" && pathname === "/internal/metrics") {
          const body = await metrics.metricsText();
          sendText(res, 200, body, metrics.registry.contentType);
          return;
        }

        if (req.method === "GET" && pathname === "/internal/stats") {
          const backend = persistenceBackend ?? config.persistence.backend;
          const persistedDocBlobsCount =
            backend === "file"
              ? await countPersistedDocBlobsOnDisk(config.dataDir)
              : null;
          const leveldbStats =
            backend === "leveldb" ? await getDirectoryStats(config.dataDir) : null;

          const topDocs = [...activeSocketsByDoc.entries()]
            .map(([docName, sockets]) => ({ docName, connections: sockets.size }))
            .filter((d) => d.connections > 0)
            .sort((a, b) => b.connections - a.connections)
            .slice(0, 10);

          const mem = process.memoryUsage();

          sendJson(res, 200, {
            ok: true,
            draining,
            rssBytes: mem.rss,
            heapUsedBytes: mem.heapUsed,
            heapTotalBytes: mem.heapTotal,
            eventLoopDelayMs: getEventLoopDelayMs(),
            persistence: {
              backend,
              dataDir: config.dataDir,
              persistedDocBlobsCount,
              leveldbStats,
            },
            encryptionEnabled: config.persistence.encryption.mode !== "off",
            tombstonesCount: tombstones.count(),
            connections: {
              ...connectionTracker.snapshot(),
              wsTotal: wss.clients.size,
              activeDocs: activeSocketsByDoc.size,
              topDocs,
            },
          });
          return;
        }

        if (req.method === "POST" && pathname === "/internal/retention/sweep") {
          // Always run the tombstone sweep. For file persistence this prunes
          // tombstones + deletes any `.yjs` blobs that should never be served
          // again. For LevelDB it is best-effort (we only delete docs whose
          // persisted name can be matched against a tombstone).
          const tombstoneResult = await triggerTombstoneSweep();

          if (retentionManager) {
            if (config.retention.ttlMs <= 0) {
              sendJson(res, 400, {
                error: "retention_disabled",
                message:
                  "Retention is disabled. Set SYNC_SERVER_RETENTION_TTL_MS to a positive integer (milliseconds).",
              });
              return;
            }

            const result = await triggerRetentionSweep();
            sendJson(res, 200, { ok: true, ...result });
            return;
          }

          sendJson(res, 200, { ok: true, ...tombstoneResult });
          return;
        }

        if (req.method === "DELETE" && pathname.startsWith("/internal/docs/")) {
          const ip = pickIp(req, config.trustProxy);
          let docName: string;
          try {
            docName = decodeURIComponent(
              pathname.slice("/internal/docs/".length)
            );
          } catch {
            logger.warn({ ip, reason: "invalid_doc_name" }, "internal_doc_purge_rejected");
            sendJson(res, 400, { error: "bad_request" });
            return;
          }

          if (!docName) {
            sendJson(res, 400, { error: "missing_doc_id" });
            return;
          }
          const docNameBytes = Buffer.byteLength(docName, "utf8");
          if (docNameBytes > MAX_DOC_NAME_BYTES) {
            logger.warn({ ip, docNameBytes }, "internal_doc_purge_rejected");
            sendJson(res, 414, { error: "doc_id_too_long" });
            return;
          }

          const docKey = docKeyFromDocName(docName);
          await tombstones.set(docKey);

          const sockets = activeSocketsByDoc.get(docName);
          const socketCount = sockets?.size ?? 0;
          if (sockets) {
            for (const ws of Array.from(sockets)) ws.terminate();
          }

          if (!clearPersistedDocument || !persistenceBackend) {
            throw new Error("Persistence is not initialized");
          }

          await clearPersistedDocument(docName);

          logger.info(
            { ip, docName, docKey, backend: persistenceBackend, terminated: socketCount },
            "internal_doc_purge_completed"
          );
          sendJson(res, 200, { ok: true });
          return;
        }

        sendJson(res, 404, { error: "not_found" });
        return;
      }

      sendJson(res, 404, { error: "not_found" });
    })().catch((err) => {
      logger.error({ err }, "http_request_failed");
      if (res.headersSent) {
        res.end();
        return;
      }
      sendJson(res, 500, { error: "internal_error" });
    });
  };

  const server = config.tls
    ? https.createServer(
        {
          cert: readFileSync(config.tls.certPath),
          key: readFileSync(config.tls.keyPath),
        },
        handler
      )
    : http.createServer(handler);

  // Reduce exposure to slowloris-style attacks (clients that slowly drip headers).
  // The sync-server request surface area is small, so a short header timeout is safe.
  server.headersTimeout = 10_000;

  wss.on("connection", (ws, req) => {
    const ip = pickIp(req, config.trustProxy);
    // Match y-websocket docName extraction (no normalization/decoding).
    const pathName = rawPathnameFromUrl(req.url ?? "/");
    const docName = pathName.startsWith("/") ? pathName.slice(1) : pathName;
    const persistedName = persistedDocNameForLeveldb(docName);
    const active = activeSocketsByDoc.get(docName) ?? new Set<WebSocket>();
    active.add(ws);
    activeSocketsByDoc.set(docName, active);
    const authCtx = (req as IncomingMessage & { auth?: unknown }).auth as
      | AuthContext
      | undefined;

    metrics.wsConnectionsTotal.inc();
    metrics.wsConnectionsCurrent.set(wss.clients.size);
    metrics.wsActiveDocsCurrent.set(activeSocketsByDoc.size);
    metrics.wsUniqueIpsCurrent.set(connectionTracker.snapshot().uniqueIps);

    let messagesInWindow = 0;
    let messageWindowStartedAtMs = Date.now();
    let oversizeMessageCounted = false;
    let oversizeMessageBytesCounted = false;
    const messageWindowMs = config.limits.messageWindowMs;
    const maxMessagesPerWindow = config.limits.maxMessagesPerWindow;

    ws.on("message", (data) => {
      const messageBytes = rawDataByteLength(data);
      if (
        config.limits.maxMessageBytes > 0 &&
        messageBytes > config.limits.maxMessageBytes
      ) {
        logger.warn(
          {
            ip,
            docName,
            userId: authCtx?.userId,
            messageBytes,
            limitBytes: config.limits.maxMessageBytes,
          },
          "ws_message_too_large"
        );
        metrics.wsMessageBytesRejectedTotal.inc(messageBytes);
        oversizeMessageBytesCounted = true;
        ws.close(1009, "Message too big");
        return;
      }

      const nowMs = Date.now();

      if (!ipMessageLimiter.consume(ip, nowMs)) {
        metrics.wsMessageBytesRejectedTotal.inc(messageBytes);
        metrics.wsMessagesRateLimitedTotal.inc();
        logger.warn(
          { ip, docName, userId: authCtx?.userId },
          "ws_ip_message_rate_limited"
        );
        ws.close(1013, "Rate limit exceeded");
        return;
      }

      if (maxMessagesPerWindow > 0 && messageWindowMs > 0) {
        if (nowMs - messageWindowStartedAtMs >= messageWindowMs) {
          messageWindowStartedAtMs = nowMs;
          messagesInWindow = 0;
        }

        messagesInWindow += 1;
        if (messagesInWindow > maxMessagesPerWindow) {
          metrics.wsMessageBytesRejectedTotal.inc(messageBytes);
          metrics.wsMessagesRateLimitedTotal.inc();
          logger.warn(
            { ip, docName, userId: authCtx?.userId },
            "ws_message_rate_limited"
          );
          ws.close(1013, "Rate limit exceeded");
          return;
        }
      }

      if (!docMessageLimiter.consume(docName, nowMs)) {
        metrics.wsMessageBytesRejectedTotal.inc(messageBytes);
        metrics.wsMessagesRateLimitedTotal.inc();
        logger.warn(
          { ip, docName, userId: authCtx?.userId },
          "ws_doc_message_rate_limited"
        );
        ws.close(1013, "Rate limit exceeded");
        return;
      }

      metrics.wsMessageBytesTotal.inc(messageBytes);
    });

    // Treat an active websocket session as document activity. y-websocket sends
    // periodic pings (30s) and receives pongs, which lets us refresh `lastSeenMs`
    // even for read-only sessions with no Yjs updates.
    ws.on("pong", () => {
      if (!shouldPersist(docName)) return;
      void retentionManager?.markSeen(persistedName);
    });

    ws.on("close", (code) => {
      metrics.wsClosesTotal.inc({ code: closeCodeLabel(code) });
      if (code === 1009) {
        // Count oversize payloads closed via 1009 (Message Too Big). Avoid
        // double-counting when `ws` also emitted a maxPayload error.
        if (!oversizeMessageCounted) {
          oversizeMessageCounted = true;
          metrics.wsMessagesTooLargeTotal.inc();
        }
        if (!oversizeMessageBytesCounted) {
          const limit = config.limits.maxMessageBytes;
          if (limit > 0) {
            metrics.wsMessageBytesRejectedTotal.inc(limit + 1);
            oversizeMessageBytesCounted = true;
          }
        }
      }
      connectionTracker.unregister(ip);
      docConnectionTracker.unregister(persistedName);
      if (leveldbDocNameHashingEnabled) {
        docConnectionTracker.unregister(docName);
      }
      const sockets = activeSocketsByDoc.get(docName);
      if (sockets) {
        sockets.delete(ws);
        if (sockets.size === 0) {
          activeSocketsByDoc.delete(docName);
          docMessageLimiter.reset(docName);
        }
      }
      metrics.wsConnectionsCurrent.set(wss.clients.size);
      metrics.wsActiveDocsCurrent.set(activeSocketsByDoc.size);
      metrics.wsUniqueIpsCurrent.set(connectionTracker.snapshot().uniqueIps);
      if (draining && drainWaiter && wss.clients.size === 0) {
        const resolve = drainWaiter;
        drainWaiter = null;
        resolve();
      }
      logger.info({ ip, docName, userId: authCtx?.userId }, "ws_connection_closed");
    });

    ws.on("error", (err) => {
      const wsErrorCode = (err as { code?: unknown }).code;
      if (
        wsErrorCode === "WS_ERR_UNSUPPORTED_MESSAGE_LENGTH" ||
        (err instanceof RangeError && err.message.includes("Max payload size exceeded"))
      ) {
        // When `ws` rejects a frame because it exceeds `maxPayload`, it may emit an
        // error and close the connection without completing a close handshake. In
        // that case the `close` event is reported as 1006 (abnormal closure) even
        // though the peer observed a 1009 close code. Count the oversize payload
        // via the error event so operators can monitor abuse.
        if (!oversizeMessageCounted) {
          oversizeMessageCounted = true;
          metrics.wsMessagesTooLargeTotal.inc();
        }
        if (!oversizeMessageBytesCounted) {
          let rejectedBytes: number | null = null;
          try {
            const receiver = (ws as any)._receiver as { _totalPayloadLength?: unknown } | undefined;
            const totalPayloadLength = receiver?._totalPayloadLength;
            if (typeof totalPayloadLength === "number" && Number.isFinite(totalPayloadLength)) {
              rejectedBytes = totalPayloadLength;
            }
          } catch {
            // ignore
          }
          if (!rejectedBytes || rejectedBytes <= 0) {
            const limit = config.limits.maxMessageBytes;
            if (limit > 0) rejectedBytes = limit + 1;
          }
          if (rejectedBytes && rejectedBytes > 0) {
            metrics.wsMessageBytesRejectedTotal.inc(rejectedBytes);
            oversizeMessageBytesCounted = true;
          }
        }
      }
      logger.warn({ err, ip, docName }, "ws_connection_error");
    });

    logger.info(
      {
        ip,
        docName,
        userId: authCtx?.userId,
        role: authCtx?.role,
        tokenType: authCtx?.tokenType,
      },
      "ws_connection_open"
    );

    const purging =
      retentionManager?.isPurging(persistedName) ||
      (leveldbDocNameHashingEnabled && retentionManager?.isPurging(docName));
    if (purging) {
      logger.warn({ ip, docName }, "ws_connection_rejected_doc_purging");
      recordUpgradeRejection("retention_purging");
      ws.close(1013, "Document is being purged");
      return;
    }

    if (shouldPersist(docName)) {
      void retentionManager?.markSeen(persistedName);
    }

    const ydoc = getYDoc(docName, config.gc);

    installYwsSecurity(ws, {
      docName,
      auth: authCtx,
      logger,
      ydoc,
      metrics,
      limits: {
        maxMessageBytes: config.limits.maxMessageBytes,
        maxAwarenessStateBytes: config.limits.maxAwarenessStateBytes,
        maxAwarenessEntries: config.limits.maxAwarenessEntries,
        maxBranchingCommitsPerDoc: config.limits.maxBranchingCommitsPerDoc,
        maxVersionsPerDoc: config.limits.maxVersionsPerDoc,
      },
      enforceRangeRestrictions: config.enforceRangeRestrictions,
      reservedRootGuard,
    });
    setupWSConnection(ws, req, { gc: config.gc });
  });

  server.on("upgrade", (req, socket, head) => {
    void (async () => {
      try {
        const ip = pickIp(req, config.trustProxy);
        if (draining) {
          recordUpgradeRejection("draining");
          sendUpgradeRejection(socket, 503, "draining");
          return;
        }
        const uaHeader = req.headers["user-agent"];
        const userAgent = (() => {
          const raw =
            typeof uaHeader === "string"
              ? uaHeader
              : Array.isArray(uaHeader)
                ? uaHeader[0]
                : undefined;
          if (!raw) return undefined;
          const trimmed = raw.trim();
          if (trimmed.length === 0) return undefined;
          return trimmed.length > MAX_USER_AGENT_CHARS
            ? trimmed.slice(0, MAX_USER_AGENT_CHARS)
            : trimmed;
        })();

        if (!connectionAttemptLimiter.consume(ip)) {
          recordUpgradeRejection("rate_limit");
          sendUpgradeRejection(socket, 429, "Too Many Requests");
          return;
        }

        if (req.method !== "GET") {
          recordUpgradeRejection("method_not_allowed");
          sendUpgradeRejection(socket, 405, "Method Not Allowed");
          return;
        }

        if (allowedOrigins) {
          const originHeader = req.headers["origin"];
          const origin =
            typeof originHeader === "string"
              ? originHeader
              : Array.isArray(originHeader)
                ? originHeader[0]
                : undefined;
          if (origin !== undefined && !allowedOrigins.has(origin.trim())) {
            recordUpgradeRejection("origin_not_allowed");
            sendUpgradeRejection(socket, 403, "Origin not allowed");
            return;
          }
        }

        if (!req.url) {
          recordUpgradeRejection("missing_doc_id");
          sendUpgradeRejection(socket, 400, "Missing URL");
          return;
        }

        const maxUrlBytes = config.limits.maxUrlBytes ?? 0;
        if (maxUrlBytes > 0 && Buffer.byteLength(req.url, "utf8") > maxUrlBytes) {
          recordUpgradeRejection("url_too_long");
          sendUpgradeRejection(socket, 414, "URL too long");
          return;
        }

        // Match y-websocket docName extraction (no normalization/decoding).
        const pathName = rawPathnameFromUrl(req.url);
        const docName = pathName.startsWith("/") ? pathName.slice(1) : pathName;
        if (!docName) {
          recordUpgradeRejection("missing_doc_id");
          sendUpgradeRejection(socket, 400, "Missing document id");
          return;
        }
        if (Buffer.byteLength(docName, "utf8") > MAX_DOC_NAME_BYTES) {
          recordUpgradeRejection("doc_id_too_long");
          sendUpgradeRejection(socket, 414, "Document id too long");
          return;
        }

        // Parse the query string for token extraction.
        const url = new URL(req.url, "http://localhost");

        const persistedName = persistedDocNameForLeveldb(docName);

        const purging =
          retentionManager?.isPurging(persistedName) ||
          (leveldbDocNameHashingEnabled && retentionManager?.isPurging(docName));
        if (purging) {
          recordUpgradeRejection("retention_purging");
          sendUpgradeRejection(socket, 503, "Document is being purged");
          return;
        }

        const token = extractToken(req, url);
        const maxTokenBytes = config.limits.maxTokenBytes ?? 0;
        if (token && maxTokenBytes > 0 && Buffer.byteLength(token, "utf8") > maxTokenBytes) {
          recordUpgradeRejection("token_too_long");
          sendUpgradeRejection(socket, 414, "Token too long");
          return;
        }
        let authCtx: AuthContext;
        try {
          authCtx = await authenticateRequest(config.auth, token, docName, {
            introspectCache,
            clientIp: ip,
            userAgent,
            metrics,
            maxTokenBytes,
          });
        } catch (err) {
          if (err instanceof AuthError) {
            recordUpgradeRejection("auth_failure");
            sendUpgradeRejection(socket, err.statusCode, err.message);
            return;
          }
          throw err;
        }

        const docKey = docKeyFromDocName(docName);
        if (tombstones.has(docKey)) {
          recordUpgradeRejection("tombstone");
          sendUpgradeRejection(socket, 410, "Gone");
          return;
        }

        // Optional revalidation for locally-verified JWTs. This prevents
        // long-lived/replayed sync tokens from granting access after sessions are
        // revoked or document permissions change.
        if (syncTokenIntrospection && authCtx.tokenType === "jwt") {
          try {
            const introspection = await syncTokenIntrospection.introspect({
              token: token!,
              docId: docName,
              clientIp: ip,
              userAgent,
            });

            if (!introspection.active) {
              const statusCode = statusCodeForIntrospectionReason(introspection.reason);
              recordUpgradeRejection("auth_failure");
              logger.warn(
                {
                  ip,
                  docName,
                  statusCode,
                  reason: introspection.reason ?? "inactive",
                  userId: authCtx.userId,
                },
                "ws_connection_rejected_introspection_inactive"
              );
              sendUpgradeRejection(socket, statusCode, introspection.reason ?? "inactive");
              return;
            }

            // Sync token introspection may clamp the effective role to the current
            // DB membership. Prefer the introspected role so demotions take effect
            // even if the client is still holding an older token.
            if (introspection.role) {
              authCtx.role = introspection.role;
            }
          } catch (err) {
            recordUpgradeRejection("auth_failure");
            if (err instanceof SyncTokenIntrospectionOverCapacityError) {
              logger.warn({ err, ip, docName }, "ws_introspection_over_capacity");
              sendUpgradeRejection(socket, 503, "Introspection over capacity");
            } else {
              logger.error({ err, ip, docName }, "ws_introspection_failed");
              sendUpgradeRejection(socket, 503, "Introspection failed");
            }
            return;
          }
        }

        // Drain mode can begin while an upgrade is in-flight (e.g. during async
        // auth/introspection). Re-check before we allocate connection slots or
        // start persistence hydration so we stop accepting new websockets as soon
        // as draining starts.
        if (draining) {
          recordUpgradeRejection("draining");
          sendUpgradeRejection(socket, 503, "draining");
          return;
        }

        const maxConnectionsPerDoc = config.limits.maxConnectionsPerDoc ?? 0;
        if (maxConnectionsPerDoc > 0) {
          const activeForDoc = activeSocketsByDoc.get(docName);
          const activeCount = activeForDoc?.size ?? 0;
          if (activeCount >= maxConnectionsPerDoc) {
            recordUpgradeRejection("max_connections_per_doc");
            sendUpgradeRejection(socket, 429, "max_connections_per_doc_exceeded");
            return;
          }
        }

        const connResult = connectionTracker.tryRegister(ip);
        if (!connResult.ok) {
          recordUpgradeRejection("rate_limit");
          sendUpgradeRejection(socket, 429, connResult.reason);
          return;
        }
        docConnectionTracker.register(persistedName);
        if (leveldbDocNameHashingEnabled) {
          docConnectionTracker.register(docName);
        }

        (req as IncomingMessage & { auth?: unknown }).auth = authCtx;

        // `y-websocket` does not await persistence `bindState()`. Ensure the
        // document is fully hydrated from persistence before allowing the client
        // to complete the websocket upgrade, so the initial sync never observes
        // a transient empty Y.Doc.
        try {
          getYDoc(docName, config.gc);
          await waitForDocLoaded?.(docName);
        } catch (err) {
          cleanupOrphanedYwsDoc(docName);
          connectionTracker.unregister(ip);
          docConnectionTracker.unregister(persistedName);
          if (leveldbDocNameHashingEnabled) {
            docConnectionTracker.unregister(docName);
          }
          recordUpgradeRejection("persistence_load_failed");
          logger.error(
            {
              err,
              ip,
              docName,
              userId: authCtx.userId,
              backend: persistenceBackend,
            },
            "ws_connection_rejected_persistence_load_failed"
          );
          sendUpgradeRejection(socket, 503, "Persistence unavailable");
          return;
        }

        // Drain mode may have begun while we were waiting for persistence
        // hydration. If so, release the allocated connection slot and reject the
        // upgrade before completing the websocket handshake.
        if (draining) {
          cleanupOrphanedYwsDoc(docName);
          connectionTracker.unregister(ip);
          docConnectionTracker.unregister(persistedName);
          if (leveldbDocNameHashingEnabled) {
            docConnectionTracker.unregister(docName);
          }
          recordUpgradeRejection("draining");
          sendUpgradeRejection(socket, 503, "draining");
          return;
        }

        try {
          wss.handleUpgrade(req, socket, head, (ws) => {
            wss.emit("connection", ws, req);
          });
        } catch (err) {
          cleanupOrphanedYwsDoc(docName);
          connectionTracker.unregister(ip);
          docConnectionTracker.unregister(persistedName);
          if (leveldbDocNameHashingEnabled) {
            docConnectionTracker.unregister(docName);
          }
          throw err;
        }
      } catch (err) {
        logger.error({ err }, "upgrade_failed");
        try {
          sendUpgradeRejection(socket, 500, "Internal Server Error");
        } catch {
          // ignore
        }
      }
    })();
  });

  const handle: SyncServerHandle = {
    async start() {
      if (!config.disableDataDirLock && !dataDirLock) {
        dataDirLock = await acquireDataDirLock(config.dataDir);
        logger.info({ lockPath: dataDirLock.lockPath }, "data_dir_lock_acquired");
      } else if (config.disableDataDirLock) {
        logger.warn(
          { dir: config.dataDir },
          "data_dir_lock_disabled_unsafe_for_multi_process"
        );
      }

      if (!eventLoopDelay) {
        eventLoopDelay = acquireEventLoopDelayHistogram();
      }

      try {
        await initPersistence();
        await new Promise<void>((resolve, reject) => {
          const onError = (err: unknown) => {
            server.off("error", onError);
            reject(err);
          };
          server.once("error", onError);
          server.listen(config.port, config.host, () => {
            server.off("error", onError);
            resolve();
          });
        });
      } catch (err) {
        if (eventLoopDelay) {
          releaseEventLoopDelayHistogram();
          eventLoopDelay = null;
        }

        if (persistenceCleanup) {
          try {
            await persistenceCleanup();
          } catch (cleanupErr) {
            logger.warn(
              { err: cleanupErr },
              "startup_persistence_cleanup_failed_after_error"
            );
          } finally {
            persistenceCleanup = null;
          }
        }

        if (dataDirLock) {
          try {
            await dataDirLock.release();
          } finally {
            dataDirLock = null;
          }
        }
        throw err;
      }
      const addr = server.address() as AddressInfo | null;
      const port = addr?.port ?? config.port;

      if (!processMetricsTimer) {
        updateProcessMetrics();
        processMetricsTimer = setInterval(updateProcessMetrics, 5_000);
        processMetricsTimer.unref();
      }

      if (config.retention.sweepIntervalMs > 0) {
        tombstoneSweepTimer = setInterval(() => {
          void triggerTombstoneSweep()
            .then((result) => {
              if (result.errors.length > 0) {
                logger.warn({ result }, "tombstone_sweep_completed_with_errors");
              } else if (
                result.expiredTombstonesRemoved > 0 ||
                result.tombstonesProcessed > 0
              ) {
                logger.info({ result }, "tombstone_sweep_completed");
              }
            })
            .catch((err) => logger.error({ err }, "tombstone_sweep_failed"));
        }, config.retention.sweepIntervalMs);
        tombstoneSweepTimer.unref();
      }

      logger.info(
        {
          host: config.host,
          port,
          dataDir: config.dataDir,
          authMode: config.auth.mode,
        },
        "sync_server_started"
      );
      return { port };
    },

    async stop() {
      if (stopInFlight) return await stopInFlight;

      stopInFlight = (async () => {
        const errors: unknown[] = [];
        const graceMs = Math.max(0, config.shutdownGraceMs ?? 0);

        // Enter drain mode: stop accepting new websocket upgrades and fail /readyz
        // so load balancers stop routing new connections.
        if (!draining) {
          draining = true;
          metrics.shutdownDrainingCurrent.set(1);
        }
        logger.info({ graceMs, wsConnections: wss.clients.size }, "shutdown_draining_started");

        if (processMetricsTimer) {
          clearInterval(processMetricsTimer);
          processMetricsTimer = null;
        }
        if (eventLoopDelay) {
          releaseEventLoopDelayHistogram();
          eventLoopDelay = null;
        }

        if (tombstoneSweepTimer) {
          clearInterval(tombstoneSweepTimer);
          tombstoneSweepTimer = null;
        }
        if (tombstoneSweepInFlight) {
          try {
            await tombstoneSweepInFlight;
          } catch (err) {
            errors.push(err);
            logger.warn({ err }, "shutdown_tombstone_sweep_failed");
          }
        }

        if (graceMs > 0 && wss.clients.size > 0) {
          await new Promise<void>((resolve) => {
            let finished = false;
            let timeout: NodeJS.Timeout | null = null;

            const finish = () => {
              if (finished) return;
              finished = true;
              if (timeout) clearTimeout(timeout);
              timeout = null;
              if (drainWaiter === finish) {
                drainWaiter = null;
              }
              resolve();
            };

            // Let `ws.on("close")` resolve early once the last client disconnects.
            drainWaiter = finish;

            if (wss.clients.size === 0) {
              finish();
              return;
            }

            timeout = setTimeout(finish, graceMs);
            timeout.unref();
          });
        }

        if (wss.clients.size > 0) {
          logger.warn(
            { graceMs, remainingWsConnections: wss.clients.size },
            "shutdown_draining_grace_expired"
          );
        } else {
          logger.info({ graceMs }, "shutdown_draining_complete");
        }

        // Force terminate remaining websocket clients (if any).
        try {
          for (const ws of wss.clients) ws.terminate();
        } catch (err) {
          errors.push(err);
          logger.warn({ err }, "shutdown_ws_terminate_failed");
        }

        try {
          await new Promise<void>((resolve, reject) => {
            try {
              wss.close((closeErr) => (closeErr ? reject(closeErr) : resolve()));
            } catch (err) {
              reject(err);
            }
          });
        } catch (err) {
          errors.push(err);
          logger.warn({ err }, "shutdown_wss_close_failed");
        }

        try {
          await new Promise<void>((resolve, reject) =>
            server.close((err) => (err ? reject(err) : resolve()))
          );
        } catch (err) {
          errors.push(err);
          logger.warn({ err }, "shutdown_http_close_failed");
        }

        if (retentionSweepTimer) {
          clearInterval(retentionSweepTimer);
          retentionSweepTimer = null;
        }
        if (retentionSweepInFlight) {
          try {
            await retentionSweepInFlight;
          } catch (err) {
            errors.push(err);
            logger.warn({ err }, "shutdown_retention_sweep_failed");
          }
        }

        if (persistenceCleanup) {
          try {
            await persistenceCleanup();
          } catch (err) {
            errors.push(err);
            logger.warn({ err }, "shutdown_persistence_cleanup_failed");
          } finally {
            persistenceCleanup = null;
          }
        }

        if (dataDirLock) {
          try {
            await dataDirLock.release();
          } catch (err) {
            errors.push(err);
            logger.warn({ err }, "data_dir_lock_release_failed");
          } finally {
            dataDirLock = null;
          }
        }

        drainWaiter = null;

        if (errors.length > 0) {
          // Keep draining mode enabled when shutdown fails so the instance
          // remains non-ready and rejects new websocket upgrades. This avoids
          // accidentally re-advertising readiness if the process keeps running
          // (e.g. due to a cleanup error).
          logger.error(
            { graceMs, wsConnections: wss.clients.size, errors: errors.length },
            "shutdown_failed_while_draining"
          );
          throw new AggregateError(errors, "sync-server shutdown failed");
        }

        draining = false;
        metrics.shutdownDrainingCurrent.set(0);
      })();

      return await stopInFlight;
    },

    getHttpUrl() {
      const addr = server.address() as AddressInfo | null;
      const port = addr?.port ?? config.port;
      const scheme = config.tls ? "https" : "http";
      return `${scheme}://${config.host}:${port}`;
    },

    getWsUrl() {
      const addr = server.address() as AddressInfo | null;
      const port = addr?.port ?? config.port;
      const scheme = config.tls ? "wss" : "ws";
      return `${scheme}://${config.host}:${port}`;
    },
  };

  return handle;
}
