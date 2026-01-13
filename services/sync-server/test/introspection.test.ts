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
  const invalidClaimsToken = signJwt({
    sub: "u-invalid",
    docId: docName,
    orgId: "o1",
    role: "editor",
    sessionId: "s-invalid"
  });
  const mfaRequiredToken = signJwt({
    sub: "u-mfa-required",
    docId: docName,
    orgId: "o1",
    role: "editor",
    sessionId: "00000000-0000-0000-0000-000000000002"
  });
  const notOrgMemberToken = signJwt({
    sub: "u-not-org-member",
    docId: docName,
    orgId: "o1",
    role: "editor",
    sessionId: "00000000-0000-0000-0000-000000000003"
  });
  const forbiddenActiveToken = signJwt({
    sub: "u-forbidden-active",
    docId: docName,
    orgId: "o1",
    role: "editor",
    sessionId: "s-forbidden-active"
  });
  const apiKeyRevokedToken = signJwt({
    sub: "u-api-key-revoked",
    docId: docName,
    orgId: "o1",
    role: "editor",
    apiKeyId: "00000000-0000-0000-0000-000000000001"
  });
  const notMemberToken = signJwt({
    sub: "u-removed",
    docId: docName,
    orgId: "o1",
    role: "editor",
    sessionId: "s-removed"
  });

  const introspectionAdminToken = "introspection-admin-token";
  const calls: Array<{ token: string; docId: string; clientIp: unknown; userAgent: unknown }> = [];

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
      calls.push({
        token: parsed.token as string,
        docId: parsed.docId as string,
        clientIp: parsed.clientIp,
        userAgent: parsed.userAgent,
      });

      let body: any;
      let statusCode = 200;
      if (parsed.token === okToken) {
        // Exercise compatibility: allow `ok` responses without `active`.
        body = { ok: true, userId: "u-ok", orgId: "o1", role: "editor" };
      } else if (parsed.token === revokedToken) {
        // Exercise compatibility: when the API returns a 403 + `reason` without an
        // explicit `{ active/ok }` boolean, the sync-server should still map the
        // reason to a 401 rejection (revoked sessions are unauthorized).
        statusCode = 403;
        body = { reason: "session_revoked", error: "forbidden" };
      } else if (parsed.token === invalidClaimsToken) {
        statusCode = 403;
        body = { reason: "invalid_claims", error: "forbidden" };
      } else if (parsed.token === mfaRequiredToken) {
        body = { active: false, reason: "mfa_required" };
      } else if (parsed.token === notOrgMemberToken) {
        body = { active: false, reason: "not_org_member" };
      } else if (parsed.token === forbiddenActiveToken) {
        // Even if the body claims `{ active: true }`, a 403 response must be treated
        // as an inactive token.
        statusCode = 403;
        body = { active: true, userId: "u-forbidden-active", orgId: "o1", role: "editor", error: "forbidden" };
      } else if (parsed.token === apiKeyRevokedToken) {
        body = { active: false, reason: "api_key_revoked" };
      } else if (parsed.token === notMemberToken) {
        body = { active: false, reason: "not_member", userId: "u-removed", orgId: "o1", role: "editor" };
      } else {
        body = { active: false, reason: "unknown_token" };
      }

      res.writeHead(statusCode, { "content-type": "application/json" });
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
  const longUserAgent = `sync-server-test-${"a".repeat(2_000)}`;
  const okWs = new WebSocket(
    `${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(okToken)}`,
    {
      headers: {
        "user-agent": longUserAgent,
      },
    }
  );
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
    `${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(invalidClaimsToken)}`,
    401
  );

  await expectWebSocketUpgradeStatus(
    `${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(forbiddenActiveToken)}`,
    403
  );

  await expectWebSocketUpgradeStatus(
    `${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(apiKeyRevokedToken)}`,
    401
  );

  await expectWebSocketUpgradeStatus(
    `${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(mfaRequiredToken)}`,
    403
  );

  await expectWebSocketUpgradeStatus(
    `${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(notOrgMemberToken)}`,
    403
  );

  await expectWebSocketUpgradeStatus(
    `${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(notMemberToken)}`,
    403
  );

  assert.ok(calls.length >= 7);
  assert.ok(calls.some((c) => c.token === revokedToken && c.docId === docName));
  assert.ok(calls.some((c) => c.token === invalidClaimsToken && c.docId === docName));
  assert.ok(calls.some((c) => c.token === forbiddenActiveToken && c.docId === docName));
  assert.ok(calls.some((c) => c.token === apiKeyRevokedToken && c.docId === docName));
  assert.ok(calls.some((c) => c.token === mfaRequiredToken && c.docId === docName));
  assert.ok(calls.some((c) => c.token === notOrgMemberToken && c.docId === docName));
  assert.ok(calls.some((c) => c.token === notMemberToken && c.docId === docName));

  const okCall = calls.find((c) => c.token === okToken && c.docId === docName);
  assert.ok(okCall);
  assert.equal(typeof okCall.userAgent, "string");
  assert.equal(okCall.userAgent, longUserAgent.slice(0, 512));

  for (const call of calls) {
    assert.ok(typeof call.clientIp === "string" && call.clientIp.length > 0);
    assert.notEqual(call.clientIp, "unknown");
  }

  const metricsRes = await fetch(`${syncServer.httpUrl}/metrics`);
  assert.equal(metricsRes.status, 200);
  const metricsText = await metricsRes.text();
  const match = metricsText.match(
    /sync_server_introspection_request_duration_ms\{(?=[^}]*path="jwt_revalidation")(?=[^}]*result="ok")[^}]*\}\s+([0-9eE+\-.]+)/
  );
  assert.ok(
    match,
    `Expected sync_server_introspection_request_duration_ms{path="jwt_revalidation",result="ok"} in /metrics:\n${metricsText}`
  );
  assert.ok(Number(match[1]) > 0);
});
