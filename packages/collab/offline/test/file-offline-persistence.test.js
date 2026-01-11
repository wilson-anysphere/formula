import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, stat, appendFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import * as Y from "yjs";

import { attachOfflinePersistence } from "../src/index.node.ts";

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

test("attachOfflinePersistence(file) restores Yjs state across restarts", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "formula-collab-offline-file-"));
  const filePath = path.join(dir, "doc.yjslog");
  const guid = `formula-collab-offline-file-${crypto.randomUUID()}`;

  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  {
    const doc = new Y.Doc({ guid });
    const persistence = attachOfflinePersistence(doc, { mode: "file", filePath });
    await persistence.whenLoaded();

    doc.getMap("cells").set("Sheet1:0:0", "hello");

    // Give the persistence layer a moment to flush.
    await sleep(10);

    persistence.destroy();
    doc.destroy();
  }

  {
    const doc = new Y.Doc({ guid });
    const persistence = attachOfflinePersistence(doc, { mode: "file", filePath });
    await persistence.whenLoaded();

    assert.equal(doc.getMap("cells").get("Sheet1:0:0"), "hello");

    await persistence.clear();
    persistence.destroy();
    doc.destroy();
  }

  {
    const doc = new Y.Doc({ guid });
    const persistence = attachOfflinePersistence(doc, { mode: "file", filePath });
    await persistence.whenLoaded();

    assert.equal(doc.getMap("cells").has("Sheet1:0:0"), false);

    persistence.destroy();
    doc.destroy();
  }
});

test("attachOfflinePersistence(file) truncates partial tail records on load", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "formula-collab-offline-file-corrupt-"));
  const filePath = path.join(dir, "doc.yjslog");
  const guid = `formula-collab-offline-file-corrupt-${crypto.randomUUID()}`;

  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  // Write a valid update first.
  {
    const doc = new Y.Doc({ guid });
    const persistence = attachOfflinePersistence(doc, { mode: "file", filePath });
    await persistence.whenLoaded();

    doc.getMap("cells").set("Sheet1:0:0", 123);
    await sleep(10);

    persistence.destroy();
    doc.destroy();
  }

  const sizeBeforeCorrupt = (await stat(filePath)).size;

  // Append an incomplete record: length prefix says 10 bytes, but we only write 3.
  await appendFile(filePath, Buffer.from([0, 0, 0, 10, 1, 2, 3]));
  const sizeWithCorruptTail = (await stat(filePath)).size;
  assert.ok(sizeWithCorruptTail > sizeBeforeCorrupt);

  // Restart and ensure the corrupt tail is ignored + truncated.
  {
    const doc = new Y.Doc({ guid });
    const persistence = attachOfflinePersistence(doc, { mode: "file", filePath });
    await persistence.whenLoaded();

    assert.equal(doc.getMap("cells").get("Sheet1:0:0"), 123);

    const sizeAfterRepair = (await stat(filePath)).size;
    assert.equal(sizeAfterRepair, sizeBeforeCorrupt);

    persistence.destroy();
    doc.destroy();
  }
});
