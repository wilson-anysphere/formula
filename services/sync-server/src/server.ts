import http from "node:http";
import { promises as fs } from "node:fs";
import { createRequire } from "node:module";
import type { AddressInfo } from "node:net";
import type { Duplex } from "node:stream";

import type { Logger } from "pino";
import WebSocket, { WebSocketServer } from "ws";
import type { IncomingMessage } from "node:http";

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
import { ConnectionTracker, TokenBucketRateLimiter } from "./limits.js";
import {
  FilePersistence,
  migrateLegacyPlaintextFilesToEncryptedFormat,
} from "./persistence.js";
import {
  DocConnectionTracker,
  LeveldbRetentionManager,
  type LeveldbPersistenceLike,
} from "./retention.js";
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

export type SyncServerHandle = {
  start: () => Promise<{ port: number }>;
  stop: () => Promise<void>;
  getHttpUrl: () => string;
  getWsUrl: () => string;
};

type LeveldbPersistence = LeveldbPersistenceLike & {
  getYDoc: (docName: string) => Promise<any>;
  storeUpdate: (docName: string, update: Uint8Array) => Promise<unknown>;
  flushDocument: (docName: string) => Promise<void>;
  destroy: () => Promise<void>;
};

export type SyncServerCreateOptions = {
  createLeveldbPersistence?: (location: string) => LeveldbPersistence;
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

  // When a document is purged while clients are connected, y-websocket can
  // re-persist it via in-flight updates and `writeState()` on last disconnect.
  // Tombstoning prevents persistence until all connections are closed.
  const tombstonedDocs = new Set<string>();
  const tombstoneForceTimers = new Map<string, NodeJS.Timeout>();
  const tombstoneGraceTimers = new Map<string, NodeJS.Timeout>();
  const tombstoneGraceMs = 500;
  const tombstoneMaxMs = 30_000;

  const shouldPersist = (docName: string) => !tombstonedDocs.has(docName);

  const tombstoneDoc = (docName: string) => {
    tombstonedDocs.add(docName);

    const grace = tombstoneGraceTimers.get(docName);
    if (grace) {
      clearTimeout(grace);
      tombstoneGraceTimers.delete(docName);
    }

    if (!tombstoneForceTimers.has(docName)) {
      const forceTimer = setTimeout(() => {
        tombstonedDocs.delete(docName);
        tombstoneForceTimers.delete(docName);
        const pendingGrace = tombstoneGraceTimers.get(docName);
        if (pendingGrace) {
          clearTimeout(pendingGrace);
          tombstoneGraceTimers.delete(docName);
        }
        logger.warn({ docName }, "internal_doc_tombstone_force_cleared");
      }, tombstoneMaxMs);
      forceTimer.unref();
      tombstoneForceTimers.set(docName, forceTimer);
    }
  };

  const maybeClearTombstone = (docName: string) => {
    if (!tombstonedDocs.has(docName)) return;
    const active = activeSocketsByDoc.get(docName);
    if (active && active.size > 0) return;
    if (tombstoneGraceTimers.has(docName)) return;

    const graceTimer = setTimeout(() => {
      const stillActive = activeSocketsByDoc.get(docName);
      if (stillActive && stillActive.size > 0) {
        tombstoneGraceTimers.delete(docName);
        return;
      }

      tombstonedDocs.delete(docName);
      const force = tombstoneForceTimers.get(docName);
      if (force) {
        clearTimeout(force);
        tombstoneForceTimers.delete(docName);
      }
      tombstoneGraceTimers.delete(docName);
    }, tombstoneGraceMs);
    graceTimer.unref();
    tombstoneGraceTimers.set(docName, graceTimer);
  };

  let persistenceInitialized = false;
  let persistenceCleanup: (() => Promise<void>) | null = null;
  let persistenceBackend: "file" | "leveldb" | null = null;
  let clearPersistedDocument: ((docName: string) => Promise<void>) | null = null;

  let retentionManager: LeveldbRetentionManager | null = null;
  let retentionSweepTimer: NodeJS.Timeout | null = null;

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
      clearPersistedDocument = (docName: string) =>
        persistence.clearDocument(docName);
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
      const ldb = createLeveldbPersistence(config.dataDir);
      persistenceBackend = "leveldb";
      retentionManager = new LeveldbRetentionManager(
        ldb,
        docConnectionTracker,
        logger,
        config.retention.ttlMs
      );

      const queues = new Map<string, Promise<void>>();
      const enqueue = (docName: string, task: () => Promise<void>) => {
        const prev = queues.get(docName) ?? Promise.resolve();
        const next = prev
          .catch(() => {
            // Keep the queue alive even if a previous write failed.
          })
          .then(task);
        queues.set(docName, next);
        void next.finally(() => {
          if (queues.get(docName) === next) queues.delete(docName);
        });
        return next;
      };

      persistenceInitialized = true;
      persistenceCleanup = async () => {
        await Promise.allSettled([...queues.values()]);
        await ldb.destroy();
      };
      clearPersistedDocument = (docName: string) =>
        enqueue(docName, () => ldb.clearDocument(docName));

      setPersistence({
        provider: ldb,
        bindState: async (docName: string, ydoc: any) => {
          const persistenceOrigin = "persistence:leveldb";

          ydoc.on("update", (update: Uint8Array, origin: unknown) => {
            if (origin === persistenceOrigin) return;
            if (!shouldPersist(docName)) return;
            void enqueue(docName, async () => {
              await ldb.storeUpdate(docName, update);
            });
            void retentionManager?.markSeen(docName);
          });

          if (shouldPersist(docName)) {
            void retentionManager?.markSeen(docName, { force: true });

            const persistedYdoc = await ldb.getYDoc(docName);
            Y.applyUpdate(
              ydoc,
              Y.encodeStateAsUpdate(persistedYdoc),
              persistenceOrigin
            );
          }
        },
        writeState: async (docName: string) => {
          if (!shouldPersist(docName)) return;
          await enqueue(docName, () => ldb.flushDocument(docName));
          void retentionManager?.markFlushed(docName);
        },
      });

      maybeStartRetentionSweeper();
      logger.info({ dir: config.dataDir }, "persistence_leveldb_enabled");
      return;
    }

    const require = createRequire(import.meta.url);
    let LeveldbPersistenceCtor: new (location: string) => LeveldbPersistence;
    try {
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      ({ LeveldbPersistence: LeveldbPersistenceCtor } = require("y-leveldb") as {
        LeveldbPersistence: new (location: string) => LeveldbPersistence;
      });
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code === "MODULE_NOT_FOUND") {
        if ((process.env.NODE_ENV ?? "development") === "production") {
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
        clearPersistedDocument = (docName: string) =>
          persistence.clearDocument(docName);
        return;
      }

      throw err;
    }

    if (config.persistence.encryption.mode === "keyring") {
      throw new Error(
        "SYNC_SERVER_PERSISTENCE_ENCRYPTION=keyring is only supported with SYNC_SERVER_PERSISTENCE_BACKEND=file."
      );
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
        const ldb = new LeveldbPersistenceCtor(config.dataDir);
        persistenceBackend = "leveldb";
        retentionManager = new LeveldbRetentionManager(
          ldb,
          docConnectionTracker,
          logger,
          config.retention.ttlMs
        );

        const queues = new Map<string, Promise<void>>();
        const enqueue = (docName: string, task: () => Promise<void>) => {
          const prev = queues.get(docName) ?? Promise.resolve();
          const next = prev
            .catch(() => {
              // Keep the queue alive even if a previous write failed.
            })
            .then(task);
          queues.set(docName, next);
          void next.finally(() => {
            if (queues.get(docName) === next) queues.delete(docName);
          });
          return next;
        };

        persistenceInitialized = true;
        persistenceCleanup = async () => {
          await Promise.allSettled([...queues.values()]);
          await ldb.destroy();
        };
        clearPersistedDocument = (docName: string) =>
          enqueue(docName, () => ldb.clearDocument(docName));

        setPersistence({
          provider: ldb,
          bindState: async (docName: string, ydoc: any) => {
            // Important: `y-websocket` does not await `bindState()`. Attach the
           // update listener first so we don't miss early client updates.
           const persistenceOrigin = "persistence:leveldb";

            ydoc.on("update", (update: Uint8Array, origin: unknown) => {
              if (origin === persistenceOrigin) return;
              if (!shouldPersist(docName)) return;
              void enqueue(docName, async () => {
                await ldb.storeUpdate(docName, update);
              });
              void retentionManager?.markSeen(docName);
            });

            if (shouldPersist(docName)) {
              void retentionManager?.markSeen(docName, { force: true });
              const persistedYdoc = await ldb.getYDoc(docName);
              Y.applyUpdate(
                ydoc,
                Y.encodeStateAsUpdate(persistedYdoc),
                persistenceOrigin
              );
            }
          },
          writeState: async (docName: string) => {
            // Compact updates on last client disconnect to keep DB size bounded.
            if (!shouldPersist(docName)) return;
            await enqueue(docName, () => ldb.flushDocument(docName));
            void retentionManager?.markFlushed(docName);
          },
        });

        maybeStartRetentionSweeper();
        logger.info({ dir: config.dataDir }, "persistence_leveldb_enabled");
        return;
      } catch (err) {
        if (attempt < maxAttempts && isLockError(err)) {
          const delayMs = 50 * Math.pow(2, attempt - 1);
          logger.warn(
            { err, attempt, delayMs },
            "leveldb_locked_retrying_open"
          );
          await new Promise((r) => setTimeout(r, delayMs));
          continue;
        }
        throw err;
      }
    }
  };

  const sendJson = (
    res: http.ServerResponse,
    statusCode: number,
    body: any
  ) => {
    res.writeHead(statusCode, { "content-type": "application/json" });
    res.end(JSON.stringify(body));
  };

  const server = http.createServer((req, res) => {
    void (async () => {
      if (!req.url) {
        res.writeHead(400).end();
        return;
      }

      const url = new URL(req.url, `http://${req.headers.host ?? "localhost"}`);
      if (req.method === "GET" && url.pathname === "/healthz") {
        const snapshot = connectionTracker.snapshot();
        sendJson(res, 200, {
          status: "ok",
          uptimeSec: Math.round(process.uptime()),
          connections: snapshot,
        });
        return;
      }

      if (req.method === "POST" && url.pathname === "/internal/retention/sweep") {
        req.resume();

        if (!config.internalAdminToken) {
          sendJson(res, 404, { error: "not_found" });
          return;
        }

        const header = req.headers["x-internal-admin-token"];
        const provided = Array.isArray(header) ? header[0] : header;
        if (provided !== config.internalAdminToken) {
          sendJson(res, 403, { error: "forbidden" });
          return;
        }

        if (config.retention.ttlMs <= 0) {
          sendJson(res, 400, {
            error: "retention_disabled",
            message:
              "Retention is disabled. Set SYNC_SERVER_RETENTION_TTL_MS to a positive integer (milliseconds).",
          });
          return;
        }

        if (!retentionManager) {
          sendJson(res, 400, {
            error: "retention_not_supported",
            message:
              "Retention is only supported when SYNC_SERVER_PERSISTENCE_BACKEND=leveldb and y-leveldb is installed.",
          });
          return;
        }

        const result = await retentionManager.sweep();
        sendJson(res, 200, { ok: true, ...result });
        return;
      }

      const internalPrefix = "/internal/docs";
      if (
        url.pathname === internalPrefix ||
        url.pathname.startsWith(`${internalPrefix}/`)
      ) {
        const ip = pickIp(req, config.trustProxy);
        if (!config.internalAdminToken) {
          logger.warn(
            { ip, path: url.pathname, reason: "disabled" },
            "internal_doc_purge_rejected"
          );
          sendJson(res, 404, { error: "not_found" });
          return;
        }

        const header = req.headers["x-internal-admin-token"];
        const provided = Array.isArray(header) ? header[0] : header;
        if (provided !== config.internalAdminToken) {
          logger.warn(
            { ip, path: url.pathname, reason: "forbidden" },
            "internal_doc_purge_rejected"
          );
          sendJson(res, 403, { error: "forbidden" });
          return;
        }

        if (req.method === "DELETE") {
          let docName = "";
          if (url.pathname.startsWith(`${internalPrefix}/`)) {
            try {
              docName = decodeURIComponent(
                url.pathname.slice(`${internalPrefix}/`.length)
              );
            } catch {
              logger.warn(
                { ip, reason: "invalid_doc_name" },
                "internal_doc_purge_rejected"
              );
              sendJson(res, 400, { error: "bad_request" });
              return;
            }
          }
          if (!docName) {
            logger.warn(
              { ip, reason: "missing_doc_name" },
              "internal_doc_purge_rejected"
            );
            sendJson(res, 400, { error: "bad_request" });
            return;
          }

          logger.info({ ip, docName }, "internal_doc_purge_requested");

          tombstoneDoc(docName);

          const sockets = activeSocketsByDoc.get(docName);
          const socketCount = sockets?.size ?? 0;
          if (sockets) {
            for (const ws of Array.from(sockets)) ws.terminate();
          }

          if (!clearPersistedDocument || !persistenceBackend) {
            throw new Error("Persistence is not initialized");
          }

          await clearPersistedDocument(docName);
          maybeClearTombstone(docName);

          logger.info(
            { docName, backend: persistenceBackend, terminated: socketCount },
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

  const wss = new WebSocketServer({ noServer: true });

  wss.on("connection", (ws, req) => {
    const ip = pickIp(req, config.trustProxy);
    const url = new URL(req.url ?? "/", "http://localhost");
    const docName = url.pathname.replace(/^\//, "");
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

    ws.on("close", () => {
      clearInterval(messageWindow);
      connectionTracker.unregister(ip);
      docConnectionTracker.unregister(docName);
      const sockets = activeSocketsByDoc.get(docName);
      if (sockets) {
        sockets.delete(ws);
        if (sockets.size === 0) activeSocketsByDoc.delete(docName);
      }
      maybeClearTombstone(docName);
      logger.info(
        { ip, docName, userId: authCtx?.userId },
        "ws_connection_closed"
      );
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

    void retentionManager?.markSeen(docName);

    if (retentionManager?.isPurging(docName)) {
      logger.warn({ ip, docName }, "ws_connection_rejected_doc_purging");
      ws.close(1013, "Document is being purged");
      return;
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
      if (tombstonedDocs.has(docName)) {
        sendUpgradeRejection(socket, 503, "Document is being purged");
        logger.warn({ ip, docName }, "ws_connection_rejected_doc_tombstoned");
        return;
      }

      if (retentionManager?.isPurging(docName)) {
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

      const connResult = connectionTracker.tryRegister(ip);
      if (!connResult.ok) {
        sendUpgradeRejection(socket, 429, connResult.reason);
        return;
      }
      docConnectionTracker.register(docName);

      (req as IncomingMessage & { auth?: unknown }).auth = authCtx;

      try {
        wss.handleUpgrade(req, socket, head, (ws) => {
          wss.emit("connection", ws, req);
        });
      } catch (err) {
        connectionTracker.unregister(ip);
        docConnectionTracker.unregister(docName);
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
      await initPersistence();
      await new Promise<void>((resolve) => {
        server.listen(config.port, config.host, () => resolve());
      });
      const addr = server.address() as AddressInfo | null;
      const port = addr?.port ?? config.port;
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
      for (const ws of wss.clients) ws.terminate();

      await new Promise<void>((resolve) => wss.close(() => resolve()));
      await new Promise<void>((resolve, reject) =>
        server.close((err) => (err ? reject(err) : resolve()))
      );

      if (retentionSweepTimer) {
        clearInterval(retentionSweepTimer);
        retentionSweepTimer = null;
      }

      if (persistenceCleanup) {
        await persistenceCleanup();
        persistenceCleanup = null;
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
