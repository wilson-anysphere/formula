import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { spawn } from "node:child_process";
import { appendFile, mkdtemp, readFile, readdir, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { createAuditEvent } from "../../audit-core/index.js";
import { NodeFsOfflineAuditQueue } from "../siem/queue/node_fs.js";
import { IndexedDbOfflineAuditQueue } from "../siem/queue/indexeddb.js";
import { OfflineAuditQueue } from "../siem/offlineQueue.js";

import { indexedDB, IDBKeyRange } from "fake-indexeddb";

function makeEvent({ secret = "supersecret", eventType = "document.opened" } = {}) {
  return createAuditEvent({
    id: randomUUID(),
    timestamp: "2025-01-01T00:00:00.000Z",
    eventType,
    actor: { type: "user", id: "user_1" },
    context: { orgId: "org_1" },
    success: true,
    details: { token: secret }
  });
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function listFiles(dirPath) {
  const entries = await readdir(dirPath, { withFileTypes: true });
  return entries.filter((entry) => entry.isFile()).map((entry) => entry.name);
}

test("NodeFsOfflineAuditQueue redacts before persistence and rotates segments", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-"));
  const queue = new NodeFsOfflineAuditQueue({
    dirPath: dir,
    // Must be larger than a single serialized audit event so we can observe an
    // `.open.jsonl` segment before rotation.
    maxSegmentBytes: 300,
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
    // Allow two events to enqueue, then reject the third.
    maxBytes: 700,
    maxSegmentBytes: 1024 * 1024,
  });

  await queue.enqueue(makeEvent({ secret: "a" }));
  await queue.enqueue(makeEvent({ secret: "b" }));

  await assert.rejects(queue.enqueue(makeEvent({ secret: "c" })), (error) => error?.code === "EQUEUEFULL");
});

test("NodeFsOfflineAuditQueue ignores acked segments when enforcing maxBytes", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-acked-gc-"));
  const queue = new NodeFsOfflineAuditQueue({
    dirPath: dir,
    maxBytes: 300,
    maxSegmentBytes: 1024 * 1024,
  });

  await queue.ensureDir();
  const segmentsDir = path.join(dir, "segments");
  const baseName = `segment-${Date.now()}-acked`;
  const ackedPath = path.join(segmentsDir, `${baseName}.acked.jsonl`);
  const cursorPath = path.join(segmentsDir, `${baseName}.cursor.json`);

  // Create a big acked file that would otherwise fill the queue and block new writes.
  await writeFile(ackedPath, "x".repeat(500), "utf8");
  await writeFile(cursorPath, JSON.stringify({ acked: 1 }), "utf8");

  await queue.enqueue(makeEvent({ secret: "after-acked" }));

  const files = await listFiles(segmentsDir);
  assert.ok(!files.includes(path.basename(ackedPath)));
  assert.ok(!files.includes(path.basename(cursorPath)));
});

test("NodeFsOfflineAuditQueue rotates segments by maxSegmentAgeMs", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-age-"));
  const queue = new NodeFsOfflineAuditQueue({
    dirPath: dir,
    maxSegmentAgeMs: 1,
    maxSegmentBytes: 1024 * 1024,
  });

  await queue.enqueue(makeEvent({ secret: "age1" }));
  await sleep(5);
  await queue.enqueue(makeEvent({ secret: "age2" }));

  const segmentsDir = path.join(dir, "segments");
  const files = await listFiles(segmentsDir);
  const openSegments = files.filter((name) => name.endsWith(".open.jsonl"));
  const pendingSegments = files.filter(
    (name) => name.endsWith(".jsonl") && !name.includes(".open.") && !name.includes(".inflight.") && !name.includes(".acked.")
  );

  assert.equal(openSegments.length, 1);
  assert.equal(pendingSegments.length, 1);
});

test("NodeFsOfflineAuditQueue reclaims enqueue lock from a crashed writer", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-enqueue-lock-"));

  const child = spawn(process.execPath, ["-e", "process.exit(0)"], { stdio: "ignore" });
  const pid = child.pid;
  await new Promise((resolve, reject) => {
    child.on("exit", resolve);
    child.on("error", reject);
  });

  await writeFile(path.join(dir, "queue.enqueue.lock"), JSON.stringify({ pid, createdAt: Date.now() }), "utf8");

  const queue = new NodeFsOfflineAuditQueue({ dirPath: dir, maxBytes: 1024 * 1024 });
  await queue.enqueue(makeEvent({ secret: "lock1" }));
});

test("NodeFsOfflineAuditQueue reclaims flush lock from a crashed flusher", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-flush-lock-"));
  const queue = new NodeFsOfflineAuditQueue({ dirPath: dir, flushBatchSize: 10 });

  const child = spawn(process.execPath, ["-e", "process.exit(0)"], { stdio: "ignore" });
  const pid = child.pid;
  await new Promise((resolve, reject) => {
    child.on("exit", resolve);
    child.on("error", reject);
  });

  await writeFile(path.join(dir, "queue.flush.lock"), JSON.stringify({ pid, createdAt: Date.now() }), "utf8");
  await queue.enqueue(makeEvent({ secret: "flush1" }));

  const sent = [];
  const exporter = {
    async sendBatch(batch) {
      sent.push(...batch.map((evt) => evt.id));
    },
  };

  const result = await queue.flushToExporter(exporter);
  assert.equal(result.sent, 1);
  assert.equal(sent.length, 1);
});

test("NodeFsOfflineAuditQueue prevents concurrent flushers from duplicating sends", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-concurrent-flush-"));
  const queueA = new NodeFsOfflineAuditQueue({ dirPath: dir, flushBatchSize: 10 });
  const queueB = new NodeFsOfflineAuditQueue({ dirPath: dir, flushBatchSize: 10 });

  const events = [makeEvent({ secret: "cf1" }), makeEvent({ secret: "cf2" }), makeEvent({ secret: "cf3" })];
  for (const event of events) await queueA.enqueue(event);

  let unblock;
  const blockPromise = new Promise((resolve) => {
    unblock = resolve;
  });

  let firstBatchStarted = false;
  const sentA = [];
  const exporterA = {
    async sendBatch(batch) {
      sentA.push(...batch.map((evt) => evt.id));
      firstBatchStarted = true;
      await blockPromise;
    },
  };

  const flushA = queueA.flushToExporter(exporterA);
  while (!firstBatchStarted) await sleep(1);

  const sentB = [];
  const exporterB = {
    async sendBatch(batch) {
      sentB.push(...batch.map((evt) => evt.id));
    },
  };

  let flushBDone = false;
  const flushB = queueB.flushToExporter(exporterB).then((result) => {
    flushBDone = true;
    return result;
  });

  await sleep(25);
  assert.equal(flushBDone, false, "expected second flusher to wait on the lock");
  unblock();

  const [resultA, resultB] = await Promise.all([flushA, flushB]);
  assert.equal(resultA.sent, events.length);
  assert.equal(resultB.sent, 0);
  assert.deepEqual(sentA.sort(), events.map((evt) => evt.id).sort());
  assert.deepEqual(sentB, []);
});

test("NodeFsOfflineAuditQueue does not lose events when flushing an open segment with a partial tail record", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-open-segment-"));
  const queue = new NodeFsOfflineAuditQueue({
    dirPath: dir,
    flushBatchSize: 10,
  });

  await queue.ensureDir();

  const segmentsDir = path.join(dir, "segments");
  const baseName = `segment-${Date.now()}-partialtail`;
  const openPath = path.join(segmentsDir, `${baseName}.open.jsonl`);

  const event1 = {
    id: randomUUID(),
    timestamp: "2025-01-01T00:00:00.000Z",
    orgId: "org_1",
    eventType: "document.opened",
    details: { token: "[REDACTED]" },
  };
  const event2 = {
    id: randomUUID(),
    timestamp: "2025-01-01T00:00:01.000Z",
    orgId: "org_1",
    eventType: "document.modified",
    details: { token: "[REDACTED]" },
  };

  const event2Json = JSON.stringify(event2);
  const partialEvent2 = event2Json.slice(0, -2); // missing closing braces to simulate an in-progress write
  const remainderEvent2 = event2Json.slice(-2) + "\n";

  await writeFile(openPath, `${JSON.stringify(event1)}\n${partialEvent2}`, "utf8");

  const firstFlushSent = [];
  let appended = false;
  const exporter = {
    async sendBatch(batch) {
      firstFlushSent.push(...batch.map((evt) => evt.id));
      if (!appended) {
        appended = true;
        await appendFile(openPath, remainderEvent2, "utf8");
      }
    },
  };

  const firstResult = await queue.flushToExporter(exporter);
  assert.equal(firstResult.sent, 1);
  assert.deepEqual(firstFlushSent, [event1.id]);

  const pending = await queue.readAll();
  assert.equal(pending.length, 1);
  assert.equal(pending[0].id, event2.id);

  const secondFlushSent = [];
  const exporter2 = {
    async sendBatch(batch) {
      secondFlushSent.push(...batch.map((evt) => evt.id));
    },
  };

  const secondResult = await queue.flushToExporter(exporter2);
  assert.equal(secondResult.sent, 1);
  assert.deepEqual(secondFlushSent, [event2.id]);

  assert.deepEqual(await queue.readAll(), []);
});

test("NodeFsOfflineAuditQueue finalizes orphaned open segments after export", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-orphan-open-"));
  const queue = new NodeFsOfflineAuditQueue({ dirPath: dir, flushBatchSize: 10 });
  await queue.ensureDir();

  const segmentsDir = path.join(dir, "segments");
  const baseName = `segment-${Date.now()}-orphaned`;
  const openPath = path.join(segmentsDir, `${baseName}.open.jsonl`);
  const lockPath = path.join(segmentsDir, `${baseName}.open.lock`);

  const event = {
    id: randomUUID(),
    timestamp: "2025-01-01T00:00:00.000Z",
    orgId: "org_1",
    eventType: "document.opened",
    details: { token: "[REDACTED]" },
  };

  await writeFile(openPath, `${JSON.stringify(event)}\n`, "utf8");
  await writeFile(lockPath, JSON.stringify({ pid: 999999, createdAt: Date.now() }), "utf8");

  const sentIds = [];
  const exporter = {
    async sendBatch(batch) {
      sentIds.push(...batch.map((evt) => evt.id));
    },
  };

  const result = await queue.flushToExporter(exporter);
  assert.equal(result.sent, 1);
  assert.deepEqual(sentIds, [event.id]);

  const remainingFiles = await listFiles(segmentsDir);
  assert.equal(remainingFiles.length, 0);
});

test("NodeFsOfflineAuditQueue cleans up lockless open segments once stale", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-lockless-open-"));
  const queue = new NodeFsOfflineAuditQueue({ dirPath: dir, flushBatchSize: 10, orphanOpenSegmentStaleMs: 0 });
  await queue.ensureDir();

  const segmentsDir = path.join(dir, "segments");
  const baseName = `segment-${Date.now()}-lockless`;
  const openPath = path.join(segmentsDir, `${baseName}.open.jsonl`);

  const event = {
    id: randomUUID(),
    timestamp: "2025-01-01T00:00:00.000Z",
    orgId: "org_1",
    eventType: "document.opened",
    details: { token: "[REDACTED]" },
  };

  await writeFile(openPath, `${JSON.stringify(event)}\n`, "utf8");

  const sentIds = [];
  const exporter = {
    async sendBatch(batch) {
      sentIds.push(...batch.map((evt) => evt.id));
    },
  };

  const result = await queue.flushToExporter(exporter);
  assert.equal(result.sent, 1);
  assert.deepEqual(sentIds, [event.id]);

  const remainingFiles = await listFiles(segmentsDir);
  assert.equal(remainingFiles.length, 0);
});

test("NodeFsOfflineAuditQueue removes stale cursor tmp files during flush", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-tmp-cleanup-"));
  const queue = new NodeFsOfflineAuditQueue({ dirPath: dir, flushBatchSize: 10 });
  await queue.ensureDir();

  const segmentsDir = path.join(dir, "segments");
  const tmpPath = path.join(segmentsDir, "segment-0-dead.cursor.json.tmp");
  await writeFile(tmpPath, "{\"acked\": 1", "utf8");

  await queue.enqueue(makeEvent({ secret: "tmp1" }));
  const exporter = { async sendBatch() {} };
  await queue.flushToExporter(exporter);

  const files = await listFiles(segmentsDir);
  assert.ok(!files.some((name) => name.endsWith(".tmp")), `expected tmp files to be removed, got: ${files.join(", ")}`);
});

test("NodeFsOfflineAuditQueue removes dangling open lock files without matching segments", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-lock-cleanup-"));
  const queue = new NodeFsOfflineAuditQueue({ dirPath: dir, orphanOpenSegmentStaleMs: 0 });
  await queue.ensureDir();

  const segmentsDir = path.join(dir, "segments");
  const lockName = `segment-${Date.now()}-dangling.open.lock`;
  await writeFile(path.join(segmentsDir, lockName), JSON.stringify({ pid: 999999, createdAt: Date.now() }), "utf8");

  await sleep(5);
  const exporter = { async sendBatch() {} };
  await queue.flushToExporter(exporter);

  const files = await listFiles(segmentsDir);
  assert.ok(!files.includes(lockName), `expected dangling lock to be removed, got: ${files.join(", ")}`);
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
  // Allow one event to enqueue, then reject the next.
  const queue = new IndexedDbOfflineAuditQueue({ dbName, maxBytes: 300 });

  await queue.enqueue(makeEvent({ secret: "a" }));
  await assert.rejects(queue.enqueue(makeEvent({ secret: "b" })), (error) => error?.code === "EQUEUEFULL");
});

test("IndexedDbOfflineAuditQueue prevents concurrent flushers from duplicating sends", async () => {
  globalThis.indexedDB = indexedDB;
  globalThis.IDBKeyRange = IDBKeyRange;

  const dbName = `siem-idb-lock-${randomUUID()}`;
  const queueA = new IndexedDbOfflineAuditQueue({ dbName, flushBatchSize: 10, flushLockTimeoutMs: 2_000 });
  const events = [makeEvent({ secret: "a1" }), makeEvent({ secret: "a2" }), makeEvent({ secret: "a3" })];
  for (const event of events) await queueA.enqueue(event);

  let unblock;
  const blockPromise = new Promise((resolve) => {
    unblock = resolve;
  });

  let firstBatchStarted = false;
  const sentA = [];
  const exporterA = {
    async sendBatch(batch) {
      sentA.push(...batch.map((evt) => evt.id));
      firstBatchStarted = true;
      await blockPromise;
    },
  };

  const flushA = queueA.flushToExporter(exporterA);
  while (!firstBatchStarted) await sleep(1);

  const queueB = new IndexedDbOfflineAuditQueue({ dbName, flushBatchSize: 10, flushLockTimeoutMs: 2_000 });
  const sentB = [];
  const exporterB = {
    async sendBatch(batch) {
      sentB.push(...batch.map((evt) => evt.id));
    },
  };

  let flushBDone = false;
  const flushB = queueB.flushToExporter(exporterB).then((result) => {
    flushBDone = true;
    return result;
  });

  await sleep(25);
  assert.equal(flushBDone, false, "expected second flusher to wait for the lock");
  unblock();

  const [resultA, resultB] = await Promise.all([flushA, flushB]);
  assert.equal(resultA.sent, 3);
  assert.equal(resultB.sent, 0);
  assert.deepEqual(sentA.sort(), events.map((evt) => evt.id).sort());
  assert.deepEqual(sentB, []);
});

test("IndexedDbOfflineAuditQueue reclaims inflight records after a crash", async () => {
  globalThis.indexedDB = indexedDB;
  globalThis.IDBKeyRange = IDBKeyRange;

  const dbName = `siem-idb-crash-${randomUUID()}`;
  const queue = new IndexedDbOfflineAuditQueue({ dbName, flushBatchSize: 2 });
  const events = [makeEvent({ secret: "c1" }), makeEvent({ secret: "c2" }), makeEvent({ secret: "c3" })];
  for (const event of events) await queue.enqueue(event);

  // Simulate a crash after claiming a batch (records are now inflight), before send/ack.
  await queue._claimBatch();

  const restarted = new IndexedDbOfflineAuditQueue({ dbName, flushBatchSize: 2 });
  const sentIds = [];
  const exporter = {
    async sendBatch(batch) {
      sentIds.push(...batch.map((evt) => evt.id));
    },
  };

  const result = await restarted.flushToExporter(exporter);
  assert.equal(result.sent, 3);
  assert.deepEqual(sentIds.sort(), events.map((evt) => evt.id).sort());
  assert.deepEqual(await restarted.readAll(), []);
});

test("OfflineAuditQueue selects Node FS backend when dirPath is provided", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "siem-queue-wrapper-fs-"));
  const queue = new OfflineAuditQueue({ dirPath: dir, flushBatchSize: 10 });

  const event = makeEvent();
  await queue.enqueue(event);

  const sent = [];
  const exporter = {
    async sendBatch(batch) {
      sent.push(...batch);
    },
  };

  const result = await queue.flushToExporter(exporter);
  assert.equal(result.sent, 1);
  assert.equal(sent.length, 1);
  assert.equal(sent[0].details.token, "[REDACTED]");
});

test("OfflineAuditQueue selects IndexedDB backend when indexedDB is available", async () => {
  globalThis.indexedDB = indexedDB;
  globalThis.IDBKeyRange = IDBKeyRange;

  const dbName = `siem-wrapper-idb-${randomUUID()}`;
  const queue = new OfflineAuditQueue({ dbName, flushBatchSize: 10 });

  const event = makeEvent();
  await queue.enqueue(event);

  const sent = [];
  const exporter = {
    async sendBatch(batch) {
      sent.push(...batch);
    },
  };

  const result = await queue.flushToExporter(exporter);
  assert.equal(result.sent, 1);
  assert.equal(sent.length, 1);
  assert.equal(sent[0].details.token, "[REDACTED]");
});
