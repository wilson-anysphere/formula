import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import jwt from "jsonwebtoken";
import WebSocket from "ws";

import { startSyncServer } from "./test-helpers.ts";

const JWT_SECRET = "test-secret";
const JWT_AUDIENCE = "formula-sync";

function signJwt(payload: Record<string, unknown>): string {
  return jwt.sign(payload, JWT_SECRET, {
    algorithm: "HS256",
    audience: JWT_AUDIENCE,
  });
}

async function wsUpgradeStatus(url: string): Promise<number> {
  return await new Promise<number>((resolve, reject) => {
    const ws = new WebSocket(url);
    let settled = false;

    const finish = (fn: () => void) => {
      if (settled) return;
      settled = true;
      fn();
    };

    ws.once("open", () => {
      finish(() => {
        ws.close();
        reject(new Error("Expected connection to be rejected"));
      });
    });

    ws.once("unexpected-response", (_req, res) => {
      finish(() => {
        res.resume();
        ws.terminate();
        resolve(res.statusCode ?? 0);
      });
    });

    ws.once("error", (err) => {
      finish(() => {
        // `ws` sometimes emits an error without `unexpected-response`. Try to
        // extract the HTTP status code if present.
        const match = String(err).match(/\b(\d{3})\b/);
        if (match) resolve(Number(match[1]));
        else reject(err);
      });
    });
  });
}

test("rejects JWT without sub when SYNC_SERVER_JWT_REQUIRE_SUB=1", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-jwt-claims-"));
  const server = await startSyncServer({
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: {
      SYNC_SERVER_JWT_REQUIRE_SUB: "1",
      SYNC_SERVER_JWT_REQUIRE_EXP: "0",
    },
  });

  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;
  const token = signJwt({ docId: docName, orgId: "o1", role: "editor" });
  const status = await wsUpgradeStatus(
    `${server.wsUrl}/${docName}?token=${encodeURIComponent(token)}`
  );
  assert.equal(status, 403);
});

test("rejects JWT without exp when SYNC_SERVER_JWT_REQUIRE_EXP=1", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-jwt-claims-"));
  const server = await startSyncServer({
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: {
      SYNC_SERVER_JWT_REQUIRE_SUB: "0",
      SYNC_SERVER_JWT_REQUIRE_EXP: "1",
    },
  });

  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;
  const token = signJwt({ sub: "u1", docId: docName, orgId: "o1", role: "editor" });
  const status = await wsUpgradeStatus(
    `${server.wsUrl}/${docName}?token=${encodeURIComponent(token)}`
  );
  assert.equal(status, 401);
});

