import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { access, mkdtemp, rm } from "node:fs/promises";
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

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === "ENOENT") return false;
    throw err;
  }
}

async function expectWsUpgradeStatus(
  url: string,
  expectedStatusCode: number
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
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

test("internal endpoints are disabled without admin token", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const res = await fetch(`${server.httpUrl}/internal/stats`);
  assert.equal(res.status, 404);
});

test("internal endpoints require x-internal-admin-token", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const missingHeader = await fetch(`${server.httpUrl}/internal/stats`);
  assert.equal(missingHeader.status, 403);

  const wrongHeader = await fetch(`${server.httpUrl}/internal/stats`, {
    headers: { "x-internal-admin-token": "wrong-token" },
  });
  assert.equal(wrongHeader.status, 403);

  const ok = await fetch(`${server.httpUrl}/internal/stats`, {
    headers: { "x-internal-admin-token": "admin-token" },
  });
  assert.equal(ok.status, 200);
  const body = (await ok.json()) as { tombstonesCount?: unknown };
  assert.equal(typeof body.tombstonesCount, "number");
});

test("purge creates tombstone and prevents doc resurrection", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "purge-doc";
  const doc = new Y.Doc();
  const provider = new WebsocketProvider(server.wsUrl, docName, doc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider.destroy();
    doc.destroy();
  });

  await waitForProviderSync(provider);
  doc.getText("t").insert(0, "goodbye");

  // Give the provider a moment to send the update before disconnecting.
  await new Promise((r) => setTimeout(r, 50));

  provider.destroy();
  doc.destroy();

  const docKey = createHash("sha256").update(docName).digest("hex");
  const persistedPath = path.join(dataDir, `${docKey}.yjs`);

  await waitForCondition(() => fileExists(persistedPath), 10_000);
  assert.equal(await fileExists(persistedPath), true);

  const purgeRes = await fetch(
    `${server.httpUrl}/internal/docs/${encodeURIComponent(docName)}`,
    {
      method: "DELETE",
      headers: { "x-internal-admin-token": "admin-token" },
    }
  );
  assert.equal(purgeRes.status, 200);
  assert.deepEqual(await purgeRes.json(), { ok: true });

  await waitForCondition(async () => !(await fileExists(persistedPath)), 10_000);
  assert.equal(await fileExists(persistedPath), false);

  await expectWsUpgradeStatus(
    `${server.wsUrl}/${encodeURIComponent(docName)}?token=test-token`,
    410
  );
});

test("purge decodes url-encoded doc ids", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  // y-websocket uses the room name as the URL path verbatim; slashes are valid
  // and must be handled by the internal purge endpoint via URL decoding.
  const docName = "purge/doc";

  const doc = new Y.Doc();
  const provider = new WebsocketProvider(server.wsUrl, docName, doc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    provider.destroy();
    doc.destroy();
  });

  await waitForProviderSync(provider);
  doc.getText("t").insert(0, "goodbye");
  await new Promise((r) => setTimeout(r, 50));

  provider.destroy();
  doc.destroy();

  const docKey = createHash("sha256").update(docName).digest("hex");
  const persistedPath = path.join(dataDir, `${docKey}.yjs`);
  await waitForCondition(() => fileExists(persistedPath), 10_000);

  const purgeRes = await fetch(
    `${server.httpUrl}/internal/docs/${encodeURIComponent(docName)}`,
    {
      method: "DELETE",
      headers: { "x-internal-admin-token": "admin-token" },
    }
  );
  assert.equal(purgeRes.status, 200);
  assert.deepEqual(await purgeRes.json(), { ok: true });

  await waitForCondition(async () => !(await fileExists(persistedPath)), 10_000);

  await expectWsUpgradeStatus(`${server.wsUrl}/${docName}?token=test-token`, 410);
});
