import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../../document/documentController.js";
import { MockEngine } from "../../document/engine.js";

import { enqueueApplyForDocument } from "../applyQueue.ts";

test("enqueueApplyForDocument continues after a rejected task", async () => {
  const doc = new DocumentController({ engine: new MockEngine() });

  /** @type {string[]} */
  const log = [];

  const first = enqueueApplyForDocument(doc, async () => {
    log.push("first");
    throw new Error("boom");
  });

  const second = enqueueApplyForDocument(doc, async () => {
    log.push("second");
    return 42;
  });

  await assert.rejects(first, (err) => err?.message === "boom");
  assert.equal(await second, 42);
  assert.deepEqual(log, ["first", "second"]);
});

