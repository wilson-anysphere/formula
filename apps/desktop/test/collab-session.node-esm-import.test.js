import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
import { createCollabSession as createSessionFromTs } from "../../../packages/collab/session/src/index.ts";

test("collab-session is importable under Node ESM when executing TS sources directly", async () => {
  const mod = await import("@formula/collab-session");

  assert.equal(typeof mod.createCollabSession, "function");
  assert.equal(typeof mod.createSession, "function");
  assert.equal(typeof createSessionFromTs, "function");

  const session = mod.createCollabSession();
  assert.ok(session?.doc);

  session.destroy?.();
  session.doc.destroy?.();
});
