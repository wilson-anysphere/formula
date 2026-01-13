import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import { startSyncServer, waitForCondition } from "./test-helpers.ts";
import { Y } from "./yjs-interop.ts";

function encodeVarUint(value: number): Uint8Array {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new Error("Invalid varUint value");
  }
  const bytes: number[] = [];
  let v = value;
  while (v > 0x7f) {
    bytes.push(0x80 | (v % 0x80));
    v = Math.floor(v / 0x80);
  }
  bytes.push(v);
  return Uint8Array.from(bytes);
}

function concatUint8Arrays(arrays: Uint8Array[]): Uint8Array {
  const total = arrays.reduce((sum, arr) => sum + arr.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const arr of arrays) {
    out.set(arr, offset);
    offset += arr.length;
  }
  return out;
}

function parsePromCounterValue(body: string, metricName: string): number {
  for (const rawLine of body.split("\n")) {
    const line = rawLine.trim();
    if (line.length === 0 || line.startsWith("#")) continue;
    if (!line.startsWith(metricName)) continue;

    // Avoid accidentally matching e.g. `${metricName}_created`.
    const next = line.charAt(metricName.length);
    if (next && next !== "{" && next !== " " && next !== "\t") continue;

    const parts = line.split(/\s+/);
    if (parts.length < 2) {
      throw new Error(`Malformed Prometheus sample line: ${line}`);
    }
    const value = Number(parts[1]);
    if (!Number.isFinite(value)) {
      throw new Error(`Invalid Prometheus sample value: ${parts[1]} (line=${line})`);
    }
    return value;
  }
  throw new Error(`Missing Prometheus metric: ${metricName}`);
}

test("increments Prometheus counter on reserved-root mutation guard rejection", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-reserved-root-metrics-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
      SYNC_SERVER_RESERVED_ROOT_GUARD_ENABLED: "true",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const metricName = "sync_server_ws_reserved_root_mutations_total";

  const getCounterValue = async (): Promise<number> => {
    const res = await fetch(`${server.httpUrl}/metrics`);
    assert.equal(res.status, 200);
    return parsePromCounterValue(await res.text(), metricName);
  };

  const before = await getCounterValue();

  const docName = "reserved-root-metrics-doc";
  const ws = new WebSocket(`${server.wsUrl}/${docName}?token=test-token`);
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

  const close = new Promise<{ code: number; reason: Buffer }>((resolve) => {
    ws.once("close", (code, reason) => resolve({ code, reason }));
  });

  const attackerDoc = new Y.Doc();
  attackerDoc.getMap("versions").set("v1", new Y.Map());
  const update = Y.encodeStateAsUpdate(attackerDoc);

  const message = concatUint8Arrays([
    encodeVarUint(0), // sync outer message
    encodeVarUint(2), // Update
    encodeVarUint(update.length),
    update,
  ]);
  ws.send(Buffer.from(message));

  const { code, reason } = await close;
  assert.equal(code, 1008);
  assert.equal(reason.toString("utf8"), "reserved root mutation");

  await waitForCondition(async () => (await getCounterValue()) === before + 1, 5_000);
  assert.equal(await getCounterValue(), before + 1);
});

