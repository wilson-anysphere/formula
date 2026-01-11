import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import net from "node:net";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";
import { randomUUID } from "node:crypto";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { createUndoService } from "../packages/collab/undo/index.js";
import { createCollabSession } from "../packages/collab/session/src/index.ts";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";

async function waitForCondition(condition, timeoutMs, intervalMs = 25) {
  const start = Date.now();
  while (Date.now() - start <= timeoutMs) {
    if (await condition()) return;
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for condition");
}

async function getAvailablePort() {
  return await new Promise((resolve, reject) => {
    const server = net.createServer();
    server.unref();
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = /** @type {{ port: number }} */ (server.address());
      const port = address.port;
      server.close(() => resolve(port));
    });
  });
}

async function waitForServerReady(baseUrl) {
  await waitForCondition(async () => {
    try {
      const res = await fetch(`${baseUrl}/healthz`);
      return res.ok;
    } catch {
      return false;
    }
  }, 10_000);
}

function waitForProviderSync(provider) {
  // y-websocket sets `.synced` once the initial sync completes.
  if (provider?.synced) return Promise.resolve();

  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      provider.off("sync", handler);
      reject(new Error("Timed out waiting for provider sync"));
    }, 10_000);
    timeout.unref();

    const handler = (isSynced) => {
      if (!isSynced) return;
      clearTimeout(timeout);
      provider.off("sync", handler);
      resolve();
    };
    provider.on("sync", handler);
  });
}

async function waitForCell(documentController, sheetId, coord, expected) {
  await waitForCondition(() => {
    const cell = documentController.getCell(sheetId, coord);
    return (cell.value ?? null) === (expected.value ?? null) && (cell.formula ?? null) === (expected.formula ?? null);
  }, 10_000);
}

test("sync-server + collab-session + Yjs↔DocumentController binder: sync, undo, persistence", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-e2e-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const httpUrl = `http://127.0.0.1:${port}`;
  const wsUrl = `ws://127.0.0.1:${port}`;

  const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const serviceDir = path.join(repoRoot, "services", "sync-server");
  const entry = path.join(serviceDir, "src", "index.ts");

  // Resolve deps from the sync-server package so this test doesn't need to add them to the root package.json.
  const requireFromSyncServer = createRequire(path.join(serviceDir, "package.json"));
  const yWebsocketDir = path.dirname(requireFromSyncServer.resolve("y-websocket/package.json"));

  // Import y-websocket's ESM entrypoint so it shares the same `yjs` module instance as the test (ESM),
  // avoiding the "Yjs was already imported" constructor-mismatch warning.
  const [{ WebsocketProvider }, wsModule] = await Promise.all([
    import(pathToFileURL(path.join(yWebsocketDir, "src", "y-websocket.js")).href),
    import(pathToFileURL(requireFromSyncServer.resolve("ws")).href),
  ]);
  const WebSocket = wsModule.default ?? wsModule.WebSocket ?? wsModule;

  let serverProcess = null;
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
        SYNC_SERVER_PERSISTENCE_BACKEND: "file",
        SYNC_SERVER_AUTH_TOKEN: "test-token",
        SYNC_SERVER_PERSIST_COMPACT_AFTER_UPDATES: "10",
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
    await new Promise((resolve) => {
      child.once("exit", () => resolve());
    });
  };

  t.after(async () => {
    await stopServer();
  });

  const createClient = ({ wsUrl: wsBaseUrl, docId, token, user, activeSheet }) => {
    const ydoc = new Y.Doc();
    const provider = new WebsocketProvider(wsBaseUrl, docId, ydoc, {
      WebSocketPolyfill: WebSocket,
      disableBc: true,
      params: { token },
    });

    const session = createCollabSession({
      doc: ydoc,
      provider,
      presence: { user, activeSheet, throttleMs: 0 },
      defaultSheetId: activeSheet,
    });

    const undo = createUndoService({ mode: "collab", doc: ydoc, scope: session.cells });
    const documentController = new DocumentController();
    const binder = bindYjsToDocumentController({
      ydoc,
      documentController,
      undoService: undo,
      defaultSheetId: activeSheet,
      userId: user.id,
    });

    let destroyed = false;
    const destroy = () => {
      if (destroyed) return;
      destroyed = true;
      binder.destroy();
      session.destroy();
      ydoc.destroy();
    };

    return { ydoc, provider, session, undo, documentController, binder, destroy };
  };

  await startServer();

  const docId = `e2e-${randomUUID()}`;

  const clientA = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-a", name: "User A", color: "#ff0000" },
    activeSheet: "Sheet1",
  });
  const clientB = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-b", name: "User B", color: "#00ff00" },
    activeSheet: "Sheet1",
  });

  t.after(() => {
    clientA.destroy();
    clientB.destroy();
  });

  await waitForProviderSync(clientA.provider);
  await waitForProviderSync(clientB.provider);

  // --- Edits propagate A -> B ---
  clientA.session.setCellFormula("Sheet1:0:1", "=1+1");
  clientA.documentController.setCellValue("Sheet1", "A1", "hello");

  await waitForCell(clientB.documentController, "Sheet1", "B1", { value: null, formula: "=1+1" });
  await waitForCell(clientB.documentController, "Sheet1", "A1", { value: "hello", formula: null });
  assert.equal(clientB.session.getCell("Sheet1:0:1")?.formula, "=1+1");
  assert.equal(clientB.session.getCell("Sheet1:0:0")?.value, "hello");

  // --- Edits propagate B -> A ---
  clientB.documentController.setCellValue("Sheet1", "C1", 123);
  await waitForCell(clientA.documentController, "Sheet1", "C1", { value: 123, formula: null });
  assert.equal(clientA.session.getCell("Sheet1:0:2")?.value, 123);

  // --- Undo only affects local-origin changes ---
  clientA.undo.undo();

  await waitForCell(clientA.documentController, "Sheet1", "A1", { value: null, formula: null });
  await waitForCell(clientB.documentController, "Sheet1", "A1", { value: null, formula: null });

  assert.equal(clientA.documentController.getCell("Sheet1", "B1").formula, "=1+1");
  assert.equal(clientB.documentController.getCell("Sheet1", "B1").formula, "=1+1");

  assert.equal(clientA.documentController.getCell("Sheet1", "C1").value, 123);
  assert.equal(clientB.documentController.getCell("Sheet1", "C1").value, 123);

  // Tear down clients and restart the server, keeping the same data directory.
  clientA.destroy();
  clientB.destroy();

  // Give the server a moment to persist state after the last client disconnects.
  await new Promise((r) => setTimeout(r, 500));

  await stopServer();
  await startServer();

  const clientC = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-c", name: "User C", color: "#0000ff" },
    activeSheet: "Sheet1",
  });

  t.after(() => {
    clientC.destroy();
  });

  await waitForProviderSync(clientC.provider);

  // --- Hydration from persisted sync-server state ---
  await waitForCell(clientC.documentController, "Sheet1", "B1", { value: null, formula: "=1+1" });
  await waitForCell(clientC.documentController, "Sheet1", "C1", { value: 123, formula: null });
  await waitForCell(clientC.documentController, "Sheet1", "A1", { value: null, formula: null });
});
