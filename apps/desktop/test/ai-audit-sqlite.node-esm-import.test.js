import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
import { SqliteAIAuditStore as StoreFromTs } from "../../../packages/ai-audit/src/sqlite.node.ts";

test("ai-audit/sqlite is importable under Node ESM when executing TS sources directly", async () => {
  const mod = await import("@formula/ai-audit/sqlite");

  assert.equal(typeof mod.SqliteAIAuditStore, "function");
  assert.equal(typeof mod.createSqliteAIAuditStoreNode, "function");
  assert.equal(typeof mod.locateSqlJsFileNode, "function");
  assert.equal(typeof StoreFromTs, "function");
});
