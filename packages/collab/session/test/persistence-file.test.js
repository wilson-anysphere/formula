import test from "node:test";
import assert from "node:assert/strict";
import { appendFile, mkdtemp, readdir, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import { FileCollabPersistence } from "@formula/collab-persistence/file";
import { createCollabSession } from "../src/index.ts";
import { KeyRing } from "../../../security/crypto/keyring.js";

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

test("FileCollabPersistence truncates partial tail records on load (crash recovery)", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "collab-file-persistence-truncate-"));
  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  const docId = `doc-${randomUUID()}`;

  let filePath = "";
  let sizeBeforeCorrupt = 0;

  // Write a valid update first.
  {
    const persistence = new FileCollabPersistence(dir, { compactAfterUpdates: 1_000 });
    const session = createCollabSession({
      docId,
      persistence,
      schema: { autoInit: false },
    });

    await session.whenLocalPersistenceLoaded();
    await session.setCellValue("Sheet1:0:0", 123);
    await session.flushLocalPersistence();

    session.destroy();
    session.doc.destroy();
    await persistence.flush(docId);
  }

  const yjsFiles = (await readdir(dir)).filter((name) => name.endsWith(".yjs"));
  assert.equal(yjsFiles.length, 1);
  filePath = path.join(dir, yjsFiles[0]);
  sizeBeforeCorrupt = (await stat(filePath)).size;

  // Append an incomplete record: length prefix says 10 bytes, but we only write 3.
  await appendFile(filePath, Buffer.from([0, 0, 0, 10, 1, 2, 3]));
  const sizeWithCorruptTail = (await stat(filePath)).size;
  assert.ok(sizeWithCorruptTail > sizeBeforeCorrupt);

  // Restart and ensure the corrupt tail is ignored + truncated before future appends.
  {
    const persistence = new FileCollabPersistence(dir, { compactAfterUpdates: 1_000 });
    const session = createCollabSession({
      docId,
      persistence,
      schema: { autoInit: false },
    });

    await session.whenLocalPersistenceLoaded();
    assert.equal((await session.getCell("Sheet1:0:0"))?.value, 123);

    const sizeAfterRepair = (await stat(filePath)).size;
    assert.equal(sizeAfterRepair, sizeBeforeCorrupt);

    session.destroy();
    session.doc.destroy();
    await persistence.flush(docId);
  }
});

test("FileCollabPersistence (encrypted) truncates partial tail records on load", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "collab-file-persistence-truncate-encrypted-"));
  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  const docId = `doc-${randomUUID()}`;
  const keyRing1 = KeyRing.create();
  const keyRing2 = KeyRing.fromJSON(keyRing1.toJSON());

  let filePath = "";
  let sizeBeforeCorrupt = 0;

  {
    const persistence = new FileCollabPersistence(dir, {
      compactAfterUpdates: 1_000,
      keyRing: keyRing1,
    });
    const session = createCollabSession({
      docId,
      persistence,
      schema: { autoInit: false },
    });

    await session.whenLocalPersistenceLoaded();
    await session.setCellValue("Sheet1:0:0", "secret");
    await session.flushLocalPersistence();

    session.destroy();
    session.doc.destroy();
    await persistence.flush(docId);
  }

  const yjsFiles = (await readdir(dir)).filter((name) => name.endsWith(".yjs"));
  assert.equal(yjsFiles.length, 1);
  filePath = path.join(dir, yjsFiles[0]);

  const header = await readFile(filePath);
  assert.equal(header.subarray(0, 8).toString("ascii"), "FMLYJS01");

  sizeBeforeCorrupt = (await stat(filePath)).size;

  // Append an incomplete encrypted record.
  await appendFile(filePath, Buffer.from([0, 0, 0, 50, 1, 2, 3]));
  const sizeWithCorruptTail = (await stat(filePath)).size;
  assert.ok(sizeWithCorruptTail > sizeBeforeCorrupt);

  {
    const persistence = new FileCollabPersistence(dir, {
      compactAfterUpdates: 1_000,
      keyRing: keyRing2,
    });
    const session = createCollabSession({
      docId,
      persistence,
      schema: { autoInit: false },
    });

    await session.whenLocalPersistenceLoaded();
    assert.equal((await session.getCell("Sheet1:0:0"))?.value, "secret");

    const sizeAfterRepair = (await stat(filePath)).size;
    assert.equal(sizeAfterRepair, sizeBeforeCorrupt);

    session.destroy();
    session.doc.destroy();
    await persistence.flush(docId);
  }
});
