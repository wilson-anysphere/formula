import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import { createLogger } from "../src/logger.js";
import { FilePersistence } from "../src/persistence.js";
import { Y } from "../src/yjs.js";

test("FilePersistence backpressure disables a doc when the per-doc queue depth is exceeded", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-file-backpressure-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const originalAppendFile = fs.appendFile.bind(fs);
  let appendCalls = 0;
  (fs as any).appendFile = async () => {
    appendCalls += 1;
    // Never resolve: simulate a stalled filesystem / blocked write.
    await new Promise<void>(() => {});
  };
  t.after(() => {
    (fs as any).appendFile = originalAppendFile;
  });

  let overloadCalls = 0;
  let overloadScope: string | null = null;
  let overloadDoc: string | null = null;

  const logger = createLogger("silent");
  const persistence = new FilePersistence(
    dataDir,
    logger,
    10,
    { mode: "off" },
    () => true,
    {
      maxQueueDepthPerDoc: 1,
      maxQueueDepthTotal: 0,
      onOverload: (docName, scope) => {
        overloadCalls += 1;
        overloadDoc = docName;
        overloadScope = scope;
      },
    }
  );

  const docName = "test-doc";
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  persistence.bindState(docName, doc);

  doc.getText("t").insert(0, "a");
  // Give the async persistence task time to reach the (patched) `fs.appendFile` call.
  // On some Node/platform combinations, this can take more than a single tick because the
  // persistence path awaits an async `fs.mkdir(...)` first.
  const start = Date.now();
  while (appendCalls === 0 && Date.now() - start < 1000) {
    // Use a timer tick (instead of only `setImmediate`) to give async fs work time to complete
    // on slower/contended environments.
    await new Promise((r) => setTimeout(r, 1));
  }
  assert.equal(appendCalls, 1);

  // Second update would exceed queue depth=1, so the doc is disabled + overload callback fired.
  doc.getText("t").insert(1, "b");
  assert.equal(overloadCalls, 1);
  assert.equal(overloadDoc, docName);
  assert.equal(overloadScope, "doc");
  assert.equal(persistence.isDocDisabled(docName), true);

  // Further updates should be ignored (no further append attempts).
  doc.getText("t").insert(2, "c");
  await new Promise((r) => setImmediate(r));
  assert.equal(appendCalls, 1);
});
