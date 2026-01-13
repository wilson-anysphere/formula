import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { createCollabVersioning } from "../src/index.ts";

test("CollabVersioning.destroy unsubscribes from Yjs document updates", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  // Instantiate workbook roots before VersionManager attaches listeners so the
  // test only measures the effects of *updates*.
  const cells = doc.getMap("cells");

  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
  });

  assert.equal(versioning.manager.dirty, false);

  versioning.destroy();

  // Workbook edits after destroy should not mark the old manager dirty.
  cells.set("Sheet1:0:0", "alpha");
  assert.equal(versioning.manager.dirty, false);

  // A destroyed manager should not attempt to snapshot since it never becomes dirty.
  assert.equal(await versioning.manager.maybeSnapshot(), null);
});

test("CollabVersioning create/destroy cycles do not accumulate dirty listeners", (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  const cells = doc.getMap("cells");

  const versioning1 = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
  });
  const manager1 = versioning1.manager;
  versioning1.destroy();

  const versioning2 = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
  });
  t.after(() => versioning2.destroy());
  const manager2 = versioning2.manager;

  cells.set("Sheet1:0:0", "alpha");

  assert.equal(manager1.dirty, false);
  assert.equal(manager2.dirty, true);
});

