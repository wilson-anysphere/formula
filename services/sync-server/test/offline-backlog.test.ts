import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

import {
  startSyncServer,
  waitForCondition,
  waitForProviderSync,
} from "./test-helpers.ts";

function setCellValue(doc: Y.Doc, cellKey: string, value: unknown): void {
  doc.transact(() => {
    const cells = doc.getMap<unknown>("cells");
    let cell = cells.get(cellKey);
    if (!(cell instanceof Y.Map)) {
      cell = new Y.Map();
      cells.set(cellKey, cell);
    }
    (cell as Y.Map<unknown>).set("value", value);
    (cell as Y.Map<unknown>).set("modifiedBy", "user-a");
  });
}

async function waitForProviderStatus(
  provider: any,
  expected: "connected" | "disconnected",
  timeoutMs: number
): Promise<void> {
  const wsconnected = Boolean(provider?.wsconnected);
  if (expected === "connected" ? wsconnected : !wsconnected) return;

  await new Promise<void>((resolve, reject) => {
    const timeout = setTimeout(() => {
      provider.off?.("status", handler);
      reject(new Error(`Timed out waiting for provider status "${expected}"`));
    }, timeoutMs);
    timeout.unref();

    const handler = (event: any) => {
      const status = typeof event === "string" ? event : event?.status;
      if (status !== expected) return;
      clearTimeout(timeout);
      provider.off?.("status", handler);
      resolve();
    };

    provider.on?.("status", handler);
  });
}

test(
  "flushes 1000+ offline cell updates on reconnect under default message limits",
  async (t) => {
    const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-offline-"));
    t.after(async () => {
      await rm(dataDir, { recursive: true, force: true });
    });

    const server = await startSyncServer({
      dataDir,
      auth: { mode: "opaque", token: "test-token" },
      env: {
        // Ensure encryption-related env vars don't leak into this test.
        SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
        SYNC_SERVER_PERSISTENCE_ENCRYPTION_KEY_B64: "",
        SYNC_SERVER_ENCRYPTION_KEYRING_JSON: "",
        SYNC_SERVER_ENCRYPTION_KEYRING_PATH: "",

        // Ensure message limit env vars don't leak into this test; empty string
        // means "use default" (see envInt in src/config.ts).
        SYNC_SERVER_MAX_MESSAGES_PER_WINDOW: "",
        SYNC_SERVER_MESSAGE_WINDOW_MS: "",
        SYNC_SERVER_MAX_MESSAGE_BYTES: "",
        SYNC_SERVER_MAX_MESSAGES_PER_DOC_WINDOW: "",
        SYNC_SERVER_DOC_MESSAGE_WINDOW_MS: "",
      },
    });
    t.after(async () => {
      await server.stop();
    });

    const docName = "offline-backlog-doc";
    const docA = new Y.Doc();
    const docB = new Y.Doc();

    const providerA = new WebsocketProvider(server.wsUrl, docName, docA, {
      WebSocketPolyfill: WebSocket,
      disableBc: true,
      params: { token: "test-token" },
    });
    const providerB = new WebsocketProvider(server.wsUrl, docName, docB, {
      WebSocketPolyfill: WebSocket,
      disableBc: true,
      params: { token: "test-token" },
    });

    t.after(() => {
      providerA.destroy();
      providerB.destroy();
      docA.destroy();
      docB.destroy();
    });

    const offlineOps = 1_100;
    const updates = Array.from({ length: offlineOps }, (_, i) => {
      const key = `Sheet1:${i}:0`;
      const value = `v${i}`;
      return { key, value };
    });

    const closeCodes: number[] = [];
    const abortController = new AbortController();

    const recordClose = (code: unknown) => {
      if (typeof code !== "number") return;
      closeCodes.push(code);
      if (code === 1009 || code === 1013) {
        abortController.abort(
          new Error(`WebSocket closed with code ${code} during offline flush`)
        );
      }
    };

    const attachProviderWsCloseListener = () => {
      const ws = (providerA as any).ws as WebSocket | undefined;
      if (!ws) return;
      const marker = "__offlineBacklogCloseListenerInstalled";
      if ((ws as any)[marker]) return;
      (ws as any)[marker] = true;
      ws.on("close", (code) => recordClose(code));
    };

    // y-websocket emits both `connection-close` and a raw `ws` close event; capture both.
    (providerA as any).on?.("connection-close", (event: any) =>
      recordClose(event?.code)
    );
    (providerA as any).on?.("status", (event: any) => {
      const status = typeof event === "string" ? event : event?.status;
      if (status === "connected") attachProviderWsCloseListener();
    });

    await waitForProviderSync(providerA);
    await waitForProviderSync(providerB);
    attachProviderWsCloseListener();

    // Simulate client A going offline while B stays connected.
    (providerA as any).disconnect?.();
    await waitForProviderStatus(providerA, "disconnected", 10_000);

    // Apply a large offline backlog (cell-ish updates) on A's local Y.Doc.
    for (const { key, value } of updates) {
      setCellValue(docA, key, value);
    }
    assert.equal(docA.getMap("cells").size, offlineOps);

    // Reconnect A and ensure the backlog flush doesn't trip server default limits.
    (providerA as any).connect?.();
    await waitForProviderStatus(providerA, "connected", 10_000);

    await waitForCondition(
      () => {
        const cellsB = docB.getMap("cells");
        if (cellsB.size !== offlineOps) return false;
        for (const { key, value } of updates) {
          const cell = cellsB.get(key);
          if (!(cell instanceof Y.Map)) return false;
          if ((cell as Y.Map<unknown>).get("value") !== value) return false;
        }
        return true;
      },
      20_000,
      50,
      abortController.signal
    );

    assert.equal(docB.getMap("cells").size, offlineOps);
    assert.ok(
      !closeCodes.includes(1009),
      `Unexpected 1009 (message too big) close codes: ${closeCodes.join(",")}`
    );
    assert.ok(
      !closeCodes.includes(1013),
      `Unexpected 1013 (rate limit) close codes: ${closeCodes.join(",")}`
    );
  }
);
