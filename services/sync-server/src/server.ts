import http from "node:http";
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
import { AuthError, authenticateRequest, extractToken } from "./auth.js";
import { ConnectionTracker, TokenBucketRateLimiter } from "./limits.js";
import { FilePersistence } from "./persistence.js";
import { Y } from "./yjs.js";

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

export function createSyncServer(config: SyncServerConfig, logger: Logger) {
  const connectionTracker = new ConnectionTracker(
    config.limits.maxConnections,
    config.limits.maxConnectionsPerIp
  );
  const connectionAttemptLimiter = new TokenBucketRateLimiter(
    config.limits.maxConnAttemptsPerWindow,
    config.limits.connAttemptWindowMs
  );
  let persistenceInitialized = false;
  let persistenceCleanup: (() => Promise<void>) | null = null;

  const initPersistence = async () => {
    if (persistenceInitialized) return;

    if (config.persistence.backend === "file") {
      persistenceInitialized = true;
      persistenceCleanup = null;
      setPersistence(
        new FilePersistence(
          config.dataDir,
          logger,
          config.persistence.compactAfterUpdates
        )
      );
      logger.info({ dir: config.dataDir }, "persistence_file_enabled");
      return;
    }

    const require = createRequire(import.meta.url);
    let LeveldbPersistence: new (location: string) => any;
    try {
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      ({ LeveldbPersistence } = require("y-leveldb") as {
        LeveldbPersistence: new (location: string) => any;
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
        persistenceInitialized = true;
        persistenceCleanup = null;
        setPersistence(
          new FilePersistence(
            config.dataDir,
            logger,
            config.persistence.compactAfterUpdates
          )
        );
        return;
      }

      throw err;
    }

    const isLockError = (err: unknown) => {
      const msg =
        err instanceof Error ? err.message : typeof err === "string" ? err : "";
      return msg.toLowerCase().includes("lock");
    };

    const maxAttempts = 5;
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      try {
        const ldb = new LeveldbPersistence(config.dataDir);

        persistenceInitialized = true;
        persistenceCleanup = async () => {
          await ldb.destroy();
        };

        setPersistence({
          provider: ldb,
          bindState: async (docName: string, ydoc: any) => {
            // Important: `y-websocket` does not await `bindState()`. Attach the
            // update listener first so we don't miss early client updates.
            const persistenceOrigin = "persistence:leveldb";

            ydoc.on("update", (update: Uint8Array, origin: unknown) => {
              if (origin === persistenceOrigin) return;
              void ldb.storeUpdate(docName, update);
            });

            const persistedYdoc = await ldb.getYDoc(docName);
            Y.applyUpdate(
              ydoc,
              Y.encodeStateAsUpdate(persistedYdoc),
              persistenceOrigin
            );
          },
          writeState: async (docName: string) => {
            // Compact updates on last client disconnect to keep DB size bounded.
            await ldb.flushDocument(docName);
          },
        });

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

  const server = http.createServer((req, res) => {
    if (!req.url) {
      res.writeHead(400).end();
      return;
    }

    const url = new URL(req.url, `http://${req.headers.host ?? "localhost"}`);
    if (req.method === "GET" && url.pathname === "/healthz") {
      const snapshot = connectionTracker.snapshot();
      res.writeHead(200, { "content-type": "application/json" });
      res.end(
        JSON.stringify({
          status: "ok",
          uptimeSec: Math.round(process.uptime()),
          connections: snapshot,
        })
      );
      return;
    }

    res.writeHead(404, { "content-type": "application/json" });
    res.end(JSON.stringify({ error: "not_found" }));
  });

  const wss = new WebSocketServer({ noServer: true });

  wss.on("connection", (ws, req) => {
    const ip = pickIp(req, config.trustProxy);
    const url = new URL(req.url ?? "/", "http://localhost");
    const docName = url.pathname.replace(/^\//, "");
    const authCtx = (req as IncomingMessage & { auth?: unknown }).auth as
      | { userId: string; tokenType: string }
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
      logger.info(
        { ip, docName, userId: authCtx?.userId },
        "ws_connection_closed"
      );
    });

    ws.on("error", (err) => {
      logger.warn({ err, ip, docName }, "ws_connection_error");
    });

    logger.info(
      { ip, docName, userId: authCtx?.userId, tokenType: authCtx?.tokenType },
      "ws_connection_open"
    );

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

      (req as IncomingMessage & { auth?: unknown }).auth = authCtx;

      try {
        wss.handleUpgrade(req, socket, head, (ws) => {
          wss.emit("connection", ws, req);
        });
      } catch (err) {
        connectionTracker.unregister(ip);
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
