import http from "node:http";
import type { Duplex } from "node:stream";
import jwt from "jsonwebtoken";
import { WebSocketServer, type WebSocket, type RawData } from "ws";

export interface SyncServerConfig {
  port: number;
  syncTokenSecret: string;
}

export interface SyncServer {
  listen: () => Promise<number>;
  close: () => Promise<void>;
}

function unauthorized(socket: Duplex): void {
  socket.write("HTTP/1.1 401 Unauthorized\r\n\r\n");
  socket.destroy();
}

export function createSyncServer(config: SyncServerConfig): SyncServer {
  const server = http.createServer((req, res) => {
    if (req.method === "GET" && req.url === "/health") {
      res.writeHead(200, { "content-type": "application/json" });
      res.end(JSON.stringify({ status: "ok" }));
      return;
    }
    res.writeHead(404);
    res.end();
  });

  const wss = new WebSocketServer({ noServer: true });

  server.on("upgrade", (req, socket, head) => {
    try {
      const url = new URL(req.url ?? "/", "http://localhost");
      const docId = url.pathname.replace(/^\//, "");
      const token = url.searchParams.get("token");

      if (!docId || !token) return unauthorized(socket);

      const decoded = jwt.verify(token, config.syncTokenSecret, {
        audience: "formula-sync"
      }) as { sub?: string; docId?: string; orgId?: string; role?: string };

      if (decoded.docId !== docId) return unauthorized(socket);

      wss.handleUpgrade(req, socket, head, (ws: WebSocket) => {
        (ws as any).context = {
          userId: decoded.sub ?? null,
          docId,
          orgId: decoded.orgId ?? null,
          role: decoded.role ?? null
        };
        wss.emit("connection", ws, req);
      });
    } catch (_err) {
      unauthorized(socket);
    }
  });

  wss.on("connection", (ws: WebSocket) => {
    const context = (ws as any).context ?? {};
    ws.send(JSON.stringify({ type: "connected", ...context }));
    ws.on("message", (data: RawData) => {
      // Placeholder: this is a token-gated transport. The CRDT sync protocol
      // (Task 24) can replace this echo behavior.
      ws.send(data);
    });
  });

  return {
    listen: async () =>
      new Promise<number>((resolve) => {
        server.listen(config.port, () => {
          const address = server.address();
          if (address && typeof address === "object") resolve(address.port);
          else resolve(config.port);
        });
      }),
    close: async () =>
      new Promise<void>((resolve, reject) => {
        wss.clients.forEach((ws) => ws.close());
        wss.close();
        server.close((err) => {
          if (err) reject(err);
          else resolve();
        });
      })
  };
}
