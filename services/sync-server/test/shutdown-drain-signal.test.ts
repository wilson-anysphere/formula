import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import { startSyncServer, waitForCondition } from "./test-helpers.ts";

function expectWebSocketUpgradeStatus(
  url: string,
  expectedStatusCode: number,
  expectedBodySubstring?: string
): Promise<void> {
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
      res.setEncoding("utf8");
      let body = "";
      res.on("data", (chunk) => {
        body += chunk;
      });
      res.on("end", () => {
        try {
          assert.equal(res.statusCode, expectedStatusCode);
          if (expectedBodySubstring) {
            assert.ok(body.includes(expectedBodySubstring));
          }
          finish(resolve);
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

test("SIGTERM triggers drain mode before shutting down existing websockets (grace period)", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-shutdown-drain-signal-"));

  const graceMs = 500;

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_SHUTDOWN_GRACE_MS: String(graceMs),
    },
  });

  let stopPromise: Promise<void> | null = null;
  t.after(async () => {
    try {
      if (stopPromise) {
        await stopPromise;
      } else {
        await server.stop();
      }
    } catch {
      // ignore
    }
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const ws = new WebSocket(`${server.wsUrl}/drain-doc?token=test-token`);
  let wsClosed = false;
  const wsClosedPromise = new Promise<void>((resolve) => {
    ws.once("close", () => {
      wsClosed = true;
      resolve();
    });
    ws.once("error", () => {
      wsClosed = true;
      resolve();
    });
  });
  t.after(() => {
    try {
      ws.terminate();
    } catch {
      // ignore
    }
  });

  await new Promise<void>((resolve, reject) => {
    ws.once("open", () => resolve());
    ws.once("error", reject);
  });

  const startedAt = Date.now();
  stopPromise = server.stop();

  // Socket should remain open briefly (we're allowing existing clients during grace).
  await new Promise((r) => setTimeout(r, 50));
  assert.equal(ws.readyState, WebSocket.OPEN);

  // Readiness should fail during drain so load balancers stop routing new work.
  await waitForCondition(async () => {
    const res = await fetch(`${server.httpUrl}/readyz`);
    if (res.status !== 503) return false;
    const json = (await res.json()) as { reason?: unknown };
    return json.reason === "draining";
  }, 2_000);

  // New websocket upgrades should be rejected while draining.
  await expectWebSocketUpgradeStatus(`${server.wsUrl}/new-doc?token=test-token`, 503, "draining");

  // Shutdown should take at least the grace period (since we keep the socket open).
  await Promise.race([
    stopPromise,
    new Promise<never>((_, reject) => {
      const timeout = setTimeout(() => reject(new Error("Timed out waiting for shutdown")), 10_000);
      timeout.unref();
    }),
  ]);

  const elapsed = Date.now() - startedAt;
  assert.ok(elapsed >= graceMs - 50, `expected >= ${graceMs}ms, got ${elapsed}ms`);

  await wsClosedPromise;
  assert.equal(wsClosed, true);
});

