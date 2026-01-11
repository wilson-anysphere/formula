import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import { FileCollabPersistence } from "@formula/collab-persistence/file";
import { createCollabSession } from "../src/index.ts";

test("CollabSession local file persistence round-trip (restart)", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "collab-file-persistence-"));
  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  const docId = `doc-${randomUUID()}`;

  {
    const persistence = new FileCollabPersistence(dir, { compactAfterUpdates: 5 });
    const session = createCollabSession({ docId, persistence });

    await session.whenLocalPersistenceLoaded();

    session.setCellValue("Sheet1:0:0", 123);
    session.setCellFormula("Sheet1:0:1", "=1+1");

    await session.flushLocalPersistence();
    session.destroy();
    session.doc.destroy();
    await persistence.flush(docId);
  }

  {
    const persistence = new FileCollabPersistence(dir, { compactAfterUpdates: 5 });
    const session = createCollabSession({ docId, persistence });

    await session.whenLocalPersistenceLoaded();

    assert.equal((await session.getCell("Sheet1:0:0"))?.value, 123);
    assert.equal((await session.getCell("Sheet1:0:1"))?.formula, "=1+1");

    session.destroy();
    session.doc.destroy();
    await persistence.flush(docId);
  }
});
