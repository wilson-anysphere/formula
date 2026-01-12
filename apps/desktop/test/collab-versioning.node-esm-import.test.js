import assert from "node:assert/strict";
import test from "node:test";

import { createCollabSession } from "@formula/collab-session";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
import { createCollabVersioning as createVersioningFromTs } from "../../../packages/collab/versioning/src/index.ts";

test("collab-versioning is importable under Node ESM when executing TS sources directly", async () => {
  const mod = await import("@formula/collab-versioning");

  assert.equal(typeof mod.createCollabVersioning, "function");
  assert.equal(typeof createVersioningFromTs, "function");

  const session = createCollabSession();
  const versioning = mod.createCollabVersioning({ session, autoStart: false });
  assert.equal(typeof versioning.listVersions, "function");

  versioning.destroy();
  session.destroy();
  session.doc.destroy();
});
