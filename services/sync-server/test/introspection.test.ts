import assert from "node:assert/strict";
import http from "node:http";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import jwt from "jsonwebtoken";
import WebSocket from "ws";

import { getAvailablePort, startSyncServer } from "./test-helpers.ts";

const JWT_SECRET = "test-secret";
const JWT_AUDIENCE = "formula-sync";

function signJwt(payload: Record<string, unknown>): string {
  return jwt.sign(payload, JWT_SECRET, {
    algorithm: "HS256",
    audience: JWT_AUDIENCE
  });
}

function expectWebSocketUpgradeStatus(url: string, expectedStatusCode: number): Promise<void> {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(url);
    let finished = false;

    const finish = (cb: () => void) => {
      if (finished) return;
      finished = true;
      try {
        ws.terminate();
      } catch {
        // ignore
      }
      cb();
    };

    ws.on("open", () => {
      finish(() => reject(new Error("Expected WebSocket upgrade rejection")));
    });

    ws.on("unexpected-response", (_req, res) => {
      try {
        assert.equal(res.statusCode, expectedStatusCode);
        res.resume();
        finish(resolve);
      } catch (err) {
        finish(() => reject(err));
      }
    });

    ws.on("error", (err) => {
      if (finished) return;
      reject(err);
    });
  });
}

function waitForWebSocketOpen(ws: WebSocket): Promise<void> {
  return new Promise((resolve, reject) => {
    const onError = (err: unknown) => {
      ws.off("open", onOpen);
      reject(err);
    };
    const onOpen = () => {
      ws.off("error", onError);
      resolve();
    };
    ws.once("error", onError);
    ws.once("open", onOpen);
  });
}

test("sync-server rejects connections when introspection marks token inactive", async (t) => {
  const docName = `doc-${Math.random().toString(16).slice(2)}`;

  const okToken = signJwt({ sub: "u-ok", docId: docName, orgId: "o1", role: "editor", sessionId: "s-ok" });
  const revokedToken = signJwt({
    sub: "u-revoked",
    docId: docName,
    orgId: "o1",
    role: "editor",
    sessionId: "s-revoked"
  });
  const notMemberToken = signJwt({
    sub: "u-removed",
    docId: docName,
    orgId: "o1",
    role: "editor",
    sessionId: "s-removed"
  });

  const introspectionAdminToken = "introspection-admin-token";
  const calls: Array<{ token: string; docId: string }> = [];

  const introspectionPort = await getAvailablePort();
  const introspectionServer = http.createServer((req, res) => {
    void (async () => {
      if (req.method !== "POST" || req.url !== "/internal/sync/introspect") {
        res.writeHead(404).end();
        return;
      }

      const header = req.headers["x-internal-admin-token"];
      const provided =
        typeof header === "string" ? header : Array.isArray(header) ? header[0] : undefined;
      if (provided !== introspectionAdminToken) {
        res.writeHead(403, { "content-type": "application/json" });
        res.end(JSON.stringify({ ok: false, active: false, error: "forbidden" }));
        return;
      }

      const chunks: Buffer[] = [];
      for await (const chunk of req) {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      }
      const parsed = JSON.parse(Buffer.concat(chunks).toString("utf8")) as any;
      calls.push({ token: parsed.token as string, docId: parsed.docId as string });

      let body: any;
      if (parsed.token === okToken) {
        body = { ok: true, active: true, userId: "u-ok", orgId: "o1", role: "editor" };
      } else if (parsed.token === revokedToken) {
        body = {
          ok: false,
          active: false,
          error: "forbidden",
          reason: "session_revoked",
          userId: "u-revoked",
          orgId: "o1",
          role: "editor",
        };
      } else if (parsed.token === notMemberToken) {
        body = {
          ok: false,
          active: false,
          error: "forbidden",
          reason: "not_member",
          userId: "u-removed",
          orgId: "o1",
          role: "editor",
        };
      } else {
        body = { ok: false, active: false, error: "forbidden", reason: "unknown_token" };
      }

      res.writeHead(200, { "content-type": "application/json" });
      res.end(JSON.stringify(body));
    })().catch((err) => {
      res.writeHead(500, { "content-type": "application/json" });
      res.end(JSON.stringify({ error: "internal_error", message: String(err) }));
    });
  });

  await new Promise<void>((resolve) => {
    introspectionServer.listen(introspectionPort, "127.0.0.1", () => resolve());
  });
  t.after(async () => {
    await new Promise<void>((resolve, reject) =>
      introspectionServer.close((err) => (err ? reject(err) : resolve()))
    );
  });

  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-introspection-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const syncServer = await startSyncServer({
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: {
      SYNC_SERVER_INTROSPECTION_URL: `http://127.0.0.1:${introspectionPort}`,
      SYNC_SERVER_INTROSPECTION_TOKEN: introspectionAdminToken,
      SYNC_SERVER_INTROSPECTION_CACHE_TTL_MS: "0"
    }
  });
  t.after(async () => {
    await syncServer.stop();
  });

  // Happy path: introspection active=true allows websocket connection.
  const okWs = new WebSocket(`${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(okToken)}`);
  t.after(() => {
    okWs.terminate();
  });
  await waitForWebSocketOpen(okWs);
  okWs.close();

  await expectWebSocketUpgradeStatus(
    `${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(revokedToken)}`,
    401
  );

  await expectWebSocketUpgradeStatus(
    `${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(notMemberToken)}`,
    403
  );

  assert.ok(calls.length >= 3);
  assert.ok(calls.some((c) => c.token === revokedToken && c.docId === docName));
  assert.ok(calls.some((c) => c.token === notMemberToken && c.docId === docName));
});
