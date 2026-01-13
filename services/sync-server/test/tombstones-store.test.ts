import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import { createLogger } from "../src/logger.js";
import { TombstoneStore } from "../src/tombstones.js";

function parseJsonLines(raw: string): any[] {
  return raw
    .trim()
    .split("\n")
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}

test("tombstone store persists using append-only tombstones.log", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-tombstones-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const logger = createLogger("silent");
  const store = new TombstoneStore(dataDir, logger);
  await store.init();

  await store.set("a", 1);
  await store.set("b", 2);
  assert.equal(store.count(), 2);

  const logPath = path.join(dataDir, "tombstones.log");
  const raw = await readFile(logPath, "utf8");
  const records = parseJsonLines(raw);
  assert.equal(records.length, 2);
  assert.deepEqual(
    records.map((r) => r.op),
    ["set", "set"]
  );
  if (process.platform !== "win32") {
    const st = await stat(logPath);
    assert.equal(st.mode & 0o777, 0o600);
  }

  const store2 = new TombstoneStore(dataDir, logger);
  await store2.init();
  assert.equal(store2.has("a"), true);
  assert.equal(store2.has("b"), true);
  assert.deepEqual(store2.entries().sort(([a], [b]) => a.localeCompare(b)), [
    ["a", { deletedAtMs: 1 }],
    ["b", { deletedAtMs: 2 }],
  ]);
});

test("tombstone store delete persists via log replay", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-tombstones-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const logger = createLogger("silent");
  const store = new TombstoneStore(dataDir, logger);
  await store.init();

  await store.set("a", 123);
  await store.delete("a");
  assert.equal(store.has("a"), false);

  const raw = await readFile(path.join(dataDir, "tombstones.log"), "utf8");
  const records = parseJsonLines(raw);
  assert.equal(records.length, 2);
  assert.deepEqual(records[0], { op: "set", docKey: "a", deletedAtMs: 123 });
  assert.deepEqual(records[1], { op: "delete", docKey: "a" });

  const store2 = new TombstoneStore(dataDir, logger);
  await store2.init();
  assert.equal(store2.has("a"), false);
});

test("sweepExpired removes expired tombstones and persists removals", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-tombstones-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const logger = createLogger("silent");
  const store = new TombstoneStore(dataDir, logger);
  await store.init();

  await store.set("expired", 0);
  await store.set("fresh", 9_000);

  const { expiredDocKeys } = await store.sweepExpired(5_000, 10_000);
  assert.deepEqual(expiredDocKeys, ["expired"]);
  assert.equal(store.has("expired"), false);
  assert.equal(store.has("fresh"), true);

  const store2 = new TombstoneStore(dataDir, logger);
  await store2.init();
  assert.equal(store2.has("expired"), false);
  assert.equal(store2.has("fresh"), true);
});

test("migrates legacy tombstones.json to the new snapshot + log format", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-tombstones-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const legacyPath = path.join(dataDir, "tombstones.json");
  await writeFile(
    legacyPath,
    `${JSON.stringify({ schemaVersion: 1, tombstones: { a: { deletedAtMs: 123 } } })}\n`,
    { mode: 0o600 }
  );

  const logger = createLogger("silent");
  const store = new TombstoneStore(dataDir, logger);
  await store.init();
  assert.equal(store.has("a"), true);

  const snapshotPath = path.join(dataDir, "tombstones.snapshot.json");
  const snapshotRaw = await readFile(snapshotPath, "utf8");
  const snapshot = JSON.parse(snapshotRaw) as any;
  assert.equal(snapshot.schemaVersion, 2);
  assert.deepEqual(snapshot.tombstones, { a: { deletedAtMs: 123 } });

  const entries = await readdir(dataDir);
  assert.equal(entries.includes("tombstones.json"), false);
  assert.ok(entries.some((name) => name.startsWith("tombstones.json.bak")));

  // A restart should load from snapshot even without a log present.
  const store2 = new TombstoneStore(dataDir, logger);
  await store2.init();
  assert.equal(store2.has("a"), true);
});

