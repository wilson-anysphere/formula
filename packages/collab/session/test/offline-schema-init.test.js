import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import * as Y from "yjs";

import { attachOfflinePersistence } from "@formula/collab-offline";
import { createCollabSession } from "../src/index.ts";

test("CollabSession schema init waits for offline load (avoids extra default sheet)", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "collab-session-offline-schema-"));
  const filePath = path.join(dir, "doc.yjslog");

  t.after(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  // Seed offline storage with a workbook that already has a sheet id that is
  // not the default "Sheet1".
  {
    const doc = new Y.Doc();
    const persistence = attachOfflinePersistence(doc, { mode: "file", filePath });
    await persistence.whenLoaded();

    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Persisted");
    sheet.set("name", "Persisted");
    sheets.push([sheet]);

    persistence.destroy();
    doc.destroy();
  }

  // Now construct a CollabSession with schema auto-init enabled (default) and
  // ensure it doesn't eagerly create the default sheet before offline load.
  const session = createCollabSession({
    doc: new Y.Doc(),
    offline: { mode: "file", filePath, autoLoad: false },
  });

  t.after(() => {
    session.destroy();
    session.doc.destroy();
  });

  await session.offline?.whenLoaded();

  const ids = session.sheets.toArray().map((s) => String(s.get("id") ?? ""));
  assert.deepEqual(ids, ["Persisted"]);
});

