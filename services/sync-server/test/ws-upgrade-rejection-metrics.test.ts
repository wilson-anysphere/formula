import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import { startSyncServer, waitForCondition } from "./test-helpers.ts";

function getRejectionMetricValue(metricsText: string, reason: string): number | null {
  const match = metricsText.match(
    new RegExp(
      `^sync_server_ws_connections_rejected_total\\{[^}]*reason="${reason}"[^}]*\\}\\s+([0-9.]+)$`,
      "m"
    )
  );
  if (!match) return null;
  const value = Number(match[1]);
  return Number.isFinite(value) ? value : null;
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

test("ws upgrade rejection metrics include pre-initialized reasons and increment on doc_id_too_long", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-upgrade-metrics-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const metricsBeforeRes = await fetch(`${server.httpUrl}/metrics`);
  assert.equal(metricsBeforeRes.ok, true);
  const metricsBefore = await metricsBeforeRes.text();

  const expectedReasons = [
    "rate_limit",
    "auth_failure",
    "draining",
    "tombstone",
    "retention_purging",
    "max_connections_per_doc",
    "doc_id_too_long",
    "missing_doc_id",
    "method_not_allowed",
    "origin_not_allowed",
  ];
  for (const reason of expectedReasons) {
    const value = getRejectionMetricValue(metricsBefore, reason);
    assert.equal(value, 0, `expected ${reason} pre-initialized at 0`);
  }

  const tooLongDocName = "a".repeat(1025);
  await expectWebSocketUpgradeStatus(
    `${server.wsUrl}/${tooLongDocName}?token=test-token`,
    414
  );

  await waitForCondition(
    async () => {
      const res = await fetch(`${server.httpUrl}/metrics`);
      if (!res.ok) return false;
      const body = await res.text();
      const value = getRejectionMetricValue(body, "doc_id_too_long");
      return typeof value === "number" && value >= 1;
    },
    5_000,
    100
  );
});
