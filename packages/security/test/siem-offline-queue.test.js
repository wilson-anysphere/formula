import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { mkdtemp, readFile, readdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { NodeFsOfflineAuditQueue } from "../siem/queue/node_fs.js";
import { IndexedDbOfflineAuditQueue } from "../siem/queue/indexeddb.js";

import { indexedDB, IDBKeyRange } from "fake-indexeddb";

function makeEvent({ secret = "supersecret", eventType = "document.opened" } = {}) {
  return {
    id: randomUUID(),
    timestamp: "2025-01-01T00:00:00.000Z",
    orgId: "org_1",
    eventType,
    details: { token: secret },
  };
}

async function listFiles(dirPath) {
  const entries = await readdir(dirPath, { withFileTypes: true });
  return entries.filter((entry) => entry.isFile()).map((entry) => entry.name);
}

test("NodeFsOfflineAuditQueue redacts before persistence and rotates segments", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-"));
  const queue = new NodeFsOfflineAuditQueue({
    dirPath: dir,
    maxSegmentBytes: 200,
    flushBatchSize: 2,
    redactionOptions: {},
  });

  const event = makeEvent();
  await queue.enqueue(event);

  const segmentsDir = path.join(dir, "segments");
  const filesAfterOne = await listFiles(segmentsDir);
  const openFileName = filesAfterOne.find((name) => name.endsWith(".open.jsonl"));
  assert.ok(openFileName, `expected an open segment file, got: ${filesAfterOne.join(", ")}`);

  const raw = await readFile(path.join(segmentsDir, openFileName), "utf8");
  assert.ok(raw.includes("[REDACTED]"));
  assert.ok(!raw.includes(event.details.token));

  // Force rotation by size.
  await queue.enqueue(makeEvent({ secret: "s2" }));
  await queue.enqueue(makeEvent({ secret: "s3" }));

  const files = await listFiles(segmentsDir);
  assert.ok(files.some((name) => name.endsWith(".jsonl") && !name.includes(".open.") && !name.includes(".inflight.")));
});

test("NodeFsOfflineAuditQueue resumes mid-flush without resending acked events", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-crash-"));
  const queue = new NodeFsOfflineAuditQueue({
    dirPath: dir,
    maxSegmentBytes: 1024 * 1024,
    flushBatchSize: 2,
  });

  const events = Array.from({ length: 5 }, () => makeEvent());
  for (const event of events) await queue.enqueue(event);

  const firstAttemptSent = [];
  const firstAttemptAcked = [];
  let callCount = 0;
  const flakyExporter = {
    async sendBatch(batch) {
      callCount += 1;
      firstAttemptSent.push(...batch.map((evt) => evt.id));
      if (callCount === 2) throw new Error("simulated send failure");
      firstAttemptAcked.push(...batch.map((evt) => evt.id));
    },
  };

  await assert.rejects(queue.flushToExporter(flakyExporter), /simulated send failure/);

  const segmentsDir = path.join(dir, "segments");
  const afterFailure = await listFiles(segmentsDir);
  assert.ok(afterFailure.some((name) => name.endsWith(".inflight.jsonl")), "expected inflight segment after failure");
  assert.ok(afterFailure.some((name) => name.endsWith(".cursor.json")), "expected cursor after partial flush");

  // "Restart" by creating a fresh instance pointed at the same dir.
  const restartedQueue = new NodeFsOfflineAuditQueue({ dirPath: dir, flushBatchSize: 2 });
  const secondAttemptSent = [];
  const exporter = {
    async sendBatch(batch) {
      secondAttemptSent.push(...batch.map((evt) => evt.id));
    },
  };

  const result = await restartedQueue.flushToExporter(exporter);
  assert.equal(result.sent, 3);

  for (const id of firstAttemptAcked) {
    assert.ok(!secondAttemptSent.includes(id), `did not expect acked event to be resent: ${id}`);
  }
  assert.deepEqual([...firstAttemptAcked, ...secondAttemptSent].sort(), events.map((evt) => evt.id).sort());

  assert.deepEqual(await restartedQueue.readAll(), []);
});

test("NodeFsOfflineAuditQueue enforces maxBytes backpressure", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-capacity-"));
  const queue = new NodeFsOfflineAuditQueue({
    dirPath: dir,
    maxBytes: 450,
    maxSegmentBytes: 1024 * 1024,
  });

  await queue.enqueue(makeEvent({ secret: "a" }));
  await queue.enqueue(makeEvent({ secret: "b" }));

  await assert.rejects(queue.enqueue(makeEvent({ secret: "c" })), (error) => error?.code === "EQUEUEFULL");
});

test("IndexedDbOfflineAuditQueue redacts before persistence and flushes batches", async () => {
  globalThis.indexedDB = indexedDB;
  globalThis.IDBKeyRange = IDBKeyRange;

  const dbName = `siem-idb-${randomUUID()}`;
  const queue = new IndexedDbOfflineAuditQueue({ dbName, flushBatchSize: 2, maxBytes: 10_000 });

  const event = makeEvent();
  await queue.enqueue(event);
  const stored = await queue.readAll();
  assert.equal(stored.length, 1);
  assert.equal(stored[0].details.token, "[REDACTED]");

  const extra = [makeEvent({ secret: "x2" }), makeEvent({ secret: "x3" })];
  for (const evt of extra) await queue.enqueue(evt);

  const sentIds = [];
  const exporter = {
    async sendBatch(batch) {
      sentIds.push(...batch.map((evt) => evt.id));
    },
  };

  const result = await queue.flushToExporter(exporter);
  assert.equal(result.sent, 3);
  assert.deepEqual(sentIds.sort(), [event, ...extra].map((evt) => evt.id).sort());
  assert.deepEqual(await queue.readAll(), []);
});

test("IndexedDbOfflineAuditQueue enforces maxBytes backpressure", async () => {
  globalThis.indexedDB = indexedDB;
  globalThis.IDBKeyRange = IDBKeyRange;

  const dbName = `siem-idb-cap-${randomUUID()}`;
  const queue = new IndexedDbOfflineAuditQueue({ dbName, maxBytes: 200 });

  await queue.enqueue(makeEvent({ secret: "a" }));
  await assert.rejects(queue.enqueue(makeEvent({ secret: "b" })), (error) => error?.code === "EQUEUEFULL");
});
