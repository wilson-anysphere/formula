import http from "node:http";
import type { IncomingMessage } from "node:http";
import { promises as fs } from "node:fs";
import { createRequire } from "node:module";
import type { AddressInfo } from "node:net";
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
} from "./auth.js";
import { acquireDataDirLock, type DataDirLockHandle } from "./dataDirLock.js";
import {
  LeveldbDocNameHashingLayer,
  persistedDocName as derivePersistedDocName,
} from "./leveldb-docname.js";
import { ConnectionTracker, TokenBucketRateLimiter } from "./limits.js";
import {
  FilePersistence,
  migrateLegacyPlaintextFilesToEncryptedFormat,
} from "./persistence.js";
import { createEncryptedLevelAdapter } from "./leveldbEncryption.js";
import {
  DocConnectionTracker,
  LeveldbRetentionManager,
  type LeveldbPersistenceLike,
} from "./retention.js";
import { requireLevelForYLeveldb } from "./leveldbLevel.js";
import { TombstoneStore, docKeyFromDocName } from "./tombstones.js";
import { Y } from "./yjs.js";
import { installYwsSecurity } from "./ywsSecurity.js";

const { setupWSConnection, setPersistence } = ywsUtils as {
  setupWSConnection: (
    conn: WebSocket,
    req: IncomingMessage,
    opts?: { gc?: boolean }
  ) => void;
  setPersistence: (persistence: unknown) => void;
};

function pickIp(req: IncomingMessage, trustProxy: boolean): string {
  if (trustProxy) {
    const forwarded = req.headers["x-forwarded-for"];
    if (typeof forwarded === "string" && forwarded.length > 0) {
      return forwarded.split(",")[0]!.trim();
    }
  }
  return req.socket.remoteAddress ?? "unknown";
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

  socket.write(headers);
  socket.destroy();
}

function sendJson(
  res: http.ServerResponse,
  statusCode: number,
  body: unknown
): void {
  res.writeHead(statusCode, { "content-type": "application/json" });
  res.end(JSON.stringify(body));
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
  const connectionTracker = new ConnectionTracker(
    config.limits.maxConnections,
    config.limits.maxConnectionsPerIp
  );
  const docConnectionTracker = new DocConnectionTracker();
  const connectionAttemptLimiter = new TokenBucketRateLimiter(
    config.limits.maxConnAttemptsPerWindow,
    config.limits.connAttemptWindowMs
  );

  const activeSocketsByDoc = new Map<string, Set<WebSocket>>();

  const tombstones = new TombstoneStore(config.dataDir, logger);
  const shouldPersist = (docName: string) => !tombstones.has(docKeyFromDocName(docName));

  type TombstoneSweepResult = {
    expiredTombstonesRemoved: number;
    tombstonesProcessed: number;
    docBlobsDeleted: number;
    errors: Array<{ docKey: string; error: string }>;
  };

  let tombstoneSweepTimer: NodeJS.Timeout | null = null;
  let tombstoneSweepInFlight: Promise<TombstoneSweepResult> | null = null;
  let tombstoneSweepCursor = 0;

  let dataDirLock: DataDirLockHandle | null = null;
  let persistenceInitialized = false;
  let persistenceCleanup: (() => Promise<void>) | null = null;
  let persistenceBackend: "file" | "leveldb" | null = null;
  let clearPersistedDocument: ((docName: string) => Promise<void>) | null = null;

  let retentionManager: LeveldbRetentionManager | null = null;
  let retentionSweepTimer: NodeJS.Timeout | null = null;

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
      void retentionManager
        ?.sweep()
        .then((result) => {
          logger.info(
            {
              ...result,
              ttlMs: config.retention.ttlMs,
              intervalMs: config.retention.sweepIntervalMs,
            },
            "retention_sweep_completed"
          );
        })
        .catch((err) => logger.error({ err }, "retention_sweep_failed"));
    }, config.retention.sweepIntervalMs);
    retentionSweepTimer.unref();
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
      persistenceBackend = "file";
      const persistence = new FilePersistence(
        config.dataDir,
        logger,
        config.persistence.compactAfterUpdates,
        config.persistence.encryption,
        shouldPersist
      );
      setPersistence(persistence);
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
      const enqueue = <T>(docName: string, task: () => Promise<T>) => {
        const prev = queues.get(docName) ?? Promise.resolve();
        const next = prev
          .catch(() => {
            // Keep the queue alive even if a previous write failed.
          })
          .then(task);
        const nextForQueue = next as Promise<unknown>;
        queues.set(docName, nextForQueue);
        void nextForQueue.finally(() => {
          if (queues.get(docName) === nextForQueue) queues.delete(docName);
        });
        return next;
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
        bindState: async (docName: string, ydoc: any) => {
          const persistenceOrigin = "persistence:leveldb";
          const retentionDocName = persistedDocNameForLeveldb(docName);

           ydoc.on("update", (update: Uint8Array, origin: unknown) => {
             if (origin === persistenceOrigin) return;
             if (!shouldPersist(docName)) return;
             if (retentionManager?.isPurging(retentionDocName)) return;
             void enqueue(retentionDocName, async () => {
                await hashedLdb.storeUpdate(docName, update);
              });
              void retentionManager?.markSeen(retentionDocName);
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
        },
        writeState: async (docName: string, ydoc: any) => {
          if (!shouldPersist(docName)) return;
          const retentionDocName = persistedDocNameForLeveldb(docName);
          if (retentionManager?.isPurging(retentionDocName)) return;

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
        persistenceBackend = "file";
        const persistence = new FilePersistence(
          config.dataDir,
          logger,
          config.persistence.compactAfterUpdates,
          config.persistence.encryption,
          shouldPersist
        );
        setPersistence(persistence);
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
        const enqueue = <T>(docName: string, task: () => Promise<T>) => {
          const prev = queues.get(docName) ?? Promise.resolve();
          const next = prev
            .catch(() => {
              // Keep the queue alive even if a previous write failed.
            })
            .then(task);
          const nextForQueue = next as Promise<unknown>;
          queues.set(docName, nextForQueue);
          void nextForQueue.finally(() => {
            if (queues.get(docName) === nextForQueue) queues.delete(docName);
          });
          return next;
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
          bindState: async (docName: string, ydoc: any) => {
            // Important: `y-websocket` does not await `bindState()`. Attach the
            // update listener first so we don't miss early client updates.
            const persistenceOrigin = "persistence:leveldb";
            const retentionDocName = persistedDocNameForLeveldb(docName);

            ydoc.on("update", (update: Uint8Array, origin: unknown) => {
              if (origin === persistenceOrigin) return;
              if (!shouldPersist(docName)) return;
              if (retentionManager?.isPurging(retentionDocName)) return;
              void enqueue(retentionDocName, async () => {
                await hashedLdb.storeUpdate(docName, update);
              });
              void retentionManager?.markSeen(retentionDocName);
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
          },
          writeState: async (docName: string, ydoc: any) => {
            // Compact updates on last client disconnect to keep DB size bounded.
            if (!shouldPersist(docName)) return;
            const retentionDocName = persistedDocNameForLeveldb(docName);
            if (retentionManager?.isPurging(retentionDocName)) return;

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

    if (persistenceBackend !== "file") {
      tombstoneSweepCursor = 0;
      return {
        expiredTombstonesRemoved: expiredDocKeys.length,
        tombstonesProcessed: 0,
        docBlobsDeleted: 0,
        errors,
      };
    }

    const docKeys = tombstones
      .entries()
      .map(([docKey]) => docKey)
      .sort((a, b) => a.localeCompare(b));

    const SWEEP_DOC_LIMIT = 100;
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
  };

  const triggerTombstoneSweep = async (): Promise<TombstoneSweepResult> => {
    if (tombstoneSweepInFlight) return await tombstoneSweepInFlight;
    tombstoneSweepInFlight = sweepTombstonesOnce().finally(() => {
      tombstoneSweepInFlight = null;
    });
    return await tombstoneSweepInFlight;
  };

  const wss = new WebSocketServer({ noServer: true });

  const server = http.createServer((req, res) => {
    void (async () => {
      if (!req.url) {
        res.writeHead(400).end();
        return;
      }

      const url = new URL(req.url, `http://${req.headers.host ?? "localhost"}`);

      if (req.method === "GET" && url.pathname === "/readyz") {
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

      if (req.method === "GET" && url.pathname === "/healthz") {
        const snapshot = connectionTracker.snapshot();
        sendJson(res, 200, {
          status: "ok",
          uptimeSec: Math.round(process.uptime()),
          connections: snapshot,
          backend: persistenceBackend ?? config.persistence.backend,
          encryptionEnabled: config.persistence.encryption.mode !== "off",
          tombstonesCount: tombstones.count(),
        });
        return;
      }

      if (url.pathname.startsWith("/internal/")) {
        req.resume();

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
        if (provided !== config.internalAdminToken) {
          sendJson(res, 403, { error: "forbidden" });
          return;
        }

        if (req.method === "GET" && url.pathname === "/internal/stats") {
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

          sendJson(res, 200, {
            ok: true,
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
              topDocs,
            },
          });
          return;
        }

        if (req.method === "POST" && url.pathname === "/internal/retention/sweep") {
          if (retentionManager) {
            if (config.retention.ttlMs <= 0) {
              sendJson(res, 400, {
                error: "retention_disabled",
                message:
                  "Retention is disabled. Set SYNC_SERVER_RETENTION_TTL_MS to a positive integer (milliseconds).",
              });
              return;
            }

            const result = await retentionManager.sweep();
            sendJson(res, 200, { ok: true, ...result });
            return;
          }

          const result = await triggerTombstoneSweep();
          sendJson(res, 200, { ok: true, ...result });
          return;
        }

        if (req.method === "DELETE" && url.pathname.startsWith("/internal/docs/")) {
          const ip = pickIp(req, config.trustProxy);
          let docName: string;
          try {
            docName = decodeURIComponent(
              url.pathname.slice("/internal/docs/".length)
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
  });

  wss.on("connection", (ws, req) => {
    const ip = pickIp(req, config.trustProxy);
    const url = new URL(req.url ?? "/", "http://localhost");
    const docName = url.pathname.replace(/^\//, "");
    const persistedName = persistedDocNameForLeveldb(docName);
    const active = activeSocketsByDoc.get(docName) ?? new Set<WebSocket>();
    active.add(ws);
    activeSocketsByDoc.set(docName, active);
    const authCtx = (req as IncomingMessage & { auth?: unknown }).auth as
      | AuthContext
      | undefined;

    let messagesInWindow = 0;
    const messageWindow = setInterval(() => {
      messagesInWindow = 0;
    }, config.limits.messageWindowMs);
    messageWindow.unref();

    ws.on("message", () => {
      messagesInWindow += 1;
      if (messagesInWindow > config.limits.maxMessagesPerWindow) {
        logger.warn(
          { ip, docName, userId: authCtx?.userId },
          "ws_message_rate_limited"
        );
        ws.close(1013, "Rate limit exceeded");
      }
    });

    // Treat an active websocket session as document activity. y-websocket sends
    // periodic pings (30s) and receives pongs, which lets us refresh `lastSeenMs`
    // even for read-only sessions with no Yjs updates.
    ws.on("pong", () => {
      if (!shouldPersist(docName)) return;
      void retentionManager?.markSeen(persistedName);
    });

    ws.on("close", () => {
      clearInterval(messageWindow);
      connectionTracker.unregister(ip);
      docConnectionTracker.unregister(persistedName);
      if (leveldbDocNameHashingEnabled) {
        docConnectionTracker.unregister(docName);
      }
      const sockets = activeSocketsByDoc.get(docName);
      if (sockets) {
        sockets.delete(ws);
        if (sockets.size === 0) activeSocketsByDoc.delete(docName);
      }
      logger.info({ ip, docName, userId: authCtx?.userId }, "ws_connection_closed");
    });

    ws.on("error", (err) => {
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
      ws.close(1013, "Document is being purged");
      return;
    }

    if (shouldPersist(docName)) {
      void retentionManager?.markSeen(persistedName);
    }

    installYwsSecurity(ws, { docName, auth: authCtx, logger });
    setupWSConnection(ws, req, { gc: config.gc });
  });

  server.on("upgrade", (req, socket, head) => {
    try {
      const ip = pickIp(req, config.trustProxy);

      if (!connectionAttemptLimiter.consume(ip)) {
        sendUpgradeRejection(socket, 429, "Too Many Requests");
        return;
      }

      if (!req.url) {
        sendUpgradeRejection(socket, 400, "Missing URL");
        return;
      }

      const url = new URL(req.url, "http://localhost");
      const docName = url.pathname.replace(/^\//, "");
      if (!docName) {
        sendUpgradeRejection(socket, 400, "Missing document id");
        return;
      }

      const persistedName = persistedDocNameForLeveldb(docName);

      const purging =
        retentionManager?.isPurging(persistedName) ||
        (leveldbDocNameHashingEnabled && retentionManager?.isPurging(docName));
      if (purging) {
        sendUpgradeRejection(socket, 503, "Document is being purged");
        return;
      }

      const token = extractToken(req, url);
      let authCtx;
      try {
        authCtx = authenticateRequest(config.auth, token, docName);
      } catch (err) {
        if (err instanceof AuthError) {
          sendUpgradeRejection(socket, err.statusCode, err.message);
          return;
        }
        throw err;
      }

      const docKey = docKeyFromDocName(docName);
      if (tombstones.has(docKey)) {
        sendUpgradeRejection(socket, 410, "Gone");
        return;
      }

      const connResult = connectionTracker.tryRegister(ip);
      if (!connResult.ok) {
        sendUpgradeRejection(socket, 429, connResult.reason);
        return;
      }
      docConnectionTracker.register(persistedName);
      if (leveldbDocNameHashingEnabled) {
        docConnectionTracker.register(docName);
      }

      (req as IncomingMessage & { auth?: unknown }).auth = authCtx;

      try {
        wss.handleUpgrade(req, socket, head, (ws) => {
          wss.emit("connection", ws, req);
        });
      } catch (err) {
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

      if (config.retention.sweepIntervalMs > 0) {
        tombstoneSweepTimer = setInterval(() => {
          void triggerTombstoneSweep().catch((err) =>
            logger.error({ err }, "tombstone_sweep_failed")
          );
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
      const errors: unknown[] = [];

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

      if (errors.length > 0) {
        throw new AggregateError(errors, "sync-server shutdown failed");
      }
    },

    getHttpUrl() {
      const addr = server.address() as AddressInfo | null;
      const port = addr?.port ?? config.port;
      return `http://${config.host}:${port}`;
    },

    getWsUrl() {
      const addr = server.address() as AddressInfo | null;
      const port = addr?.port ?? config.port;
      return `ws://${config.host}:${port}`;
    },
  };

  return handle;
}
