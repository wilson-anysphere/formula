import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import * as Y from "yjs";

import { FileCollabPersistence } from "@formula/collab-persistence/file";
import { createCollabSession } from "../src/index.ts";

test("CollabSession schema init waits for local persistence load (avoids extra default sheet)", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "collab-session-offline-schema-"));
  const docId = `doc-${randomUUID()}`;

  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  // Seed local persistence storage with a workbook that already has a sheet id
  // that is not the default "Sheet1".
  {
    const persistence = new FileCollabPersistence(dir, { compactAfterUpdates: 5 });
    const session = createCollabSession({ docId, persistence, schema: { autoInit: false } });
    await session.whenLocalPersistenceLoaded();

    const sheets = session.doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Persisted");
    sheet.set("name", "Persisted");
    sheets.push([sheet]);

    await session.flushLocalPersistence();
    session.destroy();
    session.doc.destroy();
    await persistence.flush(docId);
  }

  // Now construct a CollabSession with schema auto-init enabled (default) and
  // ensure it doesn't eagerly create the default sheet before persistence load.
  const persistence = new FileCollabPersistence(dir, { compactAfterUpdates: 5 });
  const session = createCollabSession({
    docId,
    doc: new Y.Doc({ guid: docId }),
    persistence,
  });

  t.after(async () => {
    session.destroy();
    session.doc.destroy();
    await persistence.flush(docId);
  });

  await session.whenLocalPersistenceLoaded();

  const ids = session.sheets.toArray().map((s) => String(s.get("id") ?? ""));
  assert.deepEqual(ids, ["Persisted"]);
});
