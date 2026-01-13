import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import { startSyncServer } from "./test-helpers.ts";

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

function expectWebSocketUpgradeStatus(
  url: string,
  expectedStatusCode: number,
  opts?: { origin?: string }
): Promise<string> {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(
      url,
      opts?.origin ? { headers: { Origin: opts.origin } } : undefined
    );
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
      res.setEncoding("utf8");
      let body = "";
      res.on("data", (chunk) => {
        body += chunk;
      });
      res.on("end", () => {
        try {
          assert.equal(res.statusCode, expectedStatusCode);
          finish(() => resolve(body));
        } catch (err) {
          finish(() => reject(err));
        }
      });
    });

    ws.on("error", (err) => {
      if (finished) return;
      reject(err);
    });
  });
}

test("rejects websocket upgrades with disallowed Origin when SYNC_SERVER_ALLOWED_ORIGINS is set", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-origin-allowlist-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_ALLOWED_ORIGINS: "https://allowed.example.com",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const url = `${server.wsUrl}/origin-doc?token=test-token`;
  await expectWebSocketUpgradeStatus(url, 403, { origin: "https://evil.example.com" });

  const metricsRes = await fetch(`${server.httpUrl}/metrics`);
  assert.equal(metricsRes.status, 200);
  const body = await metricsRes.text();

  const match = body.match(
    /sync_server_ws_connections_rejected_total\{[^}]*reason="origin_not_allowed"[^}]*\}\s+([0-9.]+)/
  );
  assert.ok(match, "expected metrics to include origin_not_allowed rejection reason");
  assert.ok(Number(match[1]) >= 1, `expected rejection counter >= 1 (got ${match[1]})`);
});

test("accepts websocket upgrades with allowed Origin when SYNC_SERVER_ALLOWED_ORIGINS is set", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-origin-allowed-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_ALLOWED_ORIGINS: "https://allowed.example.com",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const url = `${server.wsUrl}/origin-doc?token=test-token`;
  const ws = new WebSocket(url, { headers: { Origin: "https://allowed.example.com" } });
  t.after(() => ws.terminate());
  await waitForWebSocketOpen(ws);
});

