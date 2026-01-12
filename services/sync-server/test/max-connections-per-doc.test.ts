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

test("enforces SYNC_SERVER_MAX_CONNECTIONS_PER_DOC during websocket upgrade", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-max-per-doc-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_MAX_CONNECTIONS_PER_DOC: "2",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "max-per-doc";
  const url = `${server.wsUrl}/${docName}?token=test-token`;

  const ws1 = new WebSocket(url);
  t.after(() => ws1.terminate());
  await waitForWebSocketOpen(ws1);

  const ws2 = new WebSocket(url);
  t.after(() => ws2.terminate());
  await waitForWebSocketOpen(ws2);

  await expectWebSocketUpgradeStatus(url, 429, "max_connections_per_doc_exceeded");

  // A distinct document should still accept new connections.
  const otherUrl = `${server.wsUrl}/other-doc?token=test-token`;
  const wsOther = new WebSocket(otherUrl);
  t.after(() => wsOther.terminate());
  await waitForWebSocketOpen(wsOther);
});

