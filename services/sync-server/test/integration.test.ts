import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import type { AddressInfo } from "node:net";
import net from "node:net";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import WebSocket from "ws";
import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";

async function waitForCondition(
  condition: () => boolean | Promise<boolean>,
  timeoutMs: number,
  intervalMs: number = 25
): Promise<void> {
  const start = Date.now();
  while (Date.now() - start <= timeoutMs) {
    if (await condition()) return;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for condition");
}

async function getAvailablePort(): Promise<number> {
  return await new Promise((resolve, reject) => {
    const server = net.createServer();
    server.unref();
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address() as AddressInfo;
      const port = address.port;
      server.close(() => resolve(port));
    });
  });
}

async function waitForServerReady(baseUrl: string): Promise<void> {
  await waitForCondition(async () => {
    try {
      const res = await fetch(`${baseUrl}/healthz`);
      return res.ok;
    } catch {
      return false;
    }
  }, 10_000);
}

function waitForProviderSync(provider: {
  on: (event: string, cb: (...args: any[]) => void) => void;
  off: (event: string, cb: (...args: any[]) => void) => void;
}): Promise<void> {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      provider.off("sync", handler);
      reject(new Error("Timed out waiting for provider sync"));
    }, 10_000);
    timeout.unref();

    const handler = (isSynced: boolean) => {
      if (!isSynced) return;
      clearTimeout(timeout);
      provider.off("sync", handler);
      resolve();
    };
    provider.on("sync", handler);
  });
}

test("syncs between two clients and persists across restart", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const httpUrl = `http://127.0.0.1:${port}`;
  const wsUrl = `ws://127.0.0.1:${port}`;

  const serviceDir = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    ".."
  );
  const entry = path.join(serviceDir, "src", "index.ts");

  let serverProcess: ReturnType<typeof spawn> | null = null;
  let stdout = "";
  let stderr = "";

  const startServer = async () => {
    const child = spawn(process.execPath, ["--import", "tsx", entry], {
      cwd: serviceDir,
      env: {
        ...process.env,
        NODE_ENV: "test",
        LOG_LEVEL: "silent",
        SYNC_SERVER_HOST: "127.0.0.1",
        SYNC_SERVER_PORT: String(port),
        SYNC_SERVER_DATA_DIR: dataDir,
        SYNC_SERVER_AUTH_TOKEN: "test-token",
        SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES: "10"
      },
      stdio: ["ignore", "pipe", "pipe"],
    });

    child.stdout?.on("data", (d) => {
      stdout += d.toString();
      stdout = stdout.slice(-10_000);
    });
    child.stderr?.on("data", (d) => {
      stderr += d.toString();
      stderr = stderr.slice(-10_000);
    });

    serverProcess = child;

    try {
      await waitForServerReady(httpUrl);
    } catch (err) {
      child.kill("SIGTERM");
      throw new Error(
        `Server failed to start: ${String(err)}\nstdout:\n${stdout}\nstderr:\n${stderr}`
      );
    }
  };

  const stopServer = async () => {
    const child = serverProcess;
    if (!child) return;
    serverProcess = null;
    child.kill("SIGTERM");
    await new Promise<void>((resolve) => {
      child.once("exit", () => resolve());
    });
  };

  t.after(async () => {
    await stopServer();
  });

  await startServer();

  const docName = "test-doc";

  const doc1 = new Y.Doc();
  const doc2 = new Y.Doc();

  const provider1 = new WebsocketProvider(wsUrl, docName, doc1, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  const provider2 = new WebsocketProvider(wsUrl, docName, doc2, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider1.destroy();
    provider2.destroy();
    doc1.destroy();
    doc2.destroy();
  });

  await waitForProviderSync(provider1);
  await waitForProviderSync(provider2);

  doc1.getText("t").insert(0, "hello");

  await waitForCondition(() => doc2.getText("t").toString() === "hello", 10_000);
  assert.equal(doc2.getText("t").toString(), "hello");

  provider1.destroy();
  provider2.destroy();
  doc1.destroy();
  doc2.destroy();

  // Give the server a moment to persist state after the last client disconnects.
  await new Promise((r) => setTimeout(r, 250));

  await stopServer();
  await startServer();

  const doc3 = new Y.Doc();
  const provider3 = new WebsocketProvider(wsUrl, docName, doc3, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider3.destroy();
    doc3.destroy();
  });

  await waitForProviderSync(provider3);
  await waitForCondition(() => doc3.getText("t").toString() === "hello", 10_000);
  assert.equal(doc3.getText("t").toString(), "hello");
});
