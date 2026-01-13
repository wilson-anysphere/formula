import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { NodeFileBinaryStorage } from "../src/store/nodeFileBinaryStorage.js";

test("NodeFileBinaryStorage remove is safe when file is missing", async () => {
  const dir = path.join(os.tmpdir(), `formula-ai-rag-${randomUUID()}`);
  const filePath = path.join(dir, "db.sqlite");
  const storage = new NodeFileBinaryStorage(filePath);

  try {
    await assert.doesNotReject(storage.remove());
    assert.equal(await storage.load(), null);

    const bytes = new Uint8Array([9, 8, 7]);
    await storage.save(bytes);
    assert.deepEqual(Array.from((await storage.load()) ?? []), Array.from(bytes));

    await storage.remove();
    assert.equal(await storage.load(), null);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

