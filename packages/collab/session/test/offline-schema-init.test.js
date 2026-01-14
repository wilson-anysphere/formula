import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { randomUUID } from "node:crypto";

import * as Y from "yjs";

import { FileCollabPersistence } from "@formula/collab-persistence/file";
import { createCollabSession } from "../src/index.ts";

/**
 * Best-effort temp directory cleanup.
 *
 * `FileCollabPersistence` can still have in-flight async writes when the test completes, which
 * occasionally causes `rm({ recursive: true })` to throw `ENOTEMPTY`. Retry briefly to avoid
 * flaky failures.
 *
 * @param {string} target
 */
async function rmWithRetries(target) {
  const retryable = new Set(["ENOTEMPTY", "EBUSY", "EPERM"]);
  const maxAttempts = 8;
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    try {
      await rm(target, { recursive: true, force: true });
      return;
    } catch (err) {
      const code = /** @type {any} */ (err)?.code;
      if (!retryable.has(code) || attempt === maxAttempts - 1) throw err;
      await new Promise((r) => setTimeout(r, 25 * (attempt + 1)));
    }
  }
}

test("CollabSession schema init waits for local persistence load (avoids extra default sheet)", async (t) => {
  const dir = await mkdtemp(path.join(tmpdir(), "collab-session-offline-schema-"));
  const docId = `doc-${randomUUID()}`;

  t.after(async () => {
    await rmWithRetries(dir);
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
