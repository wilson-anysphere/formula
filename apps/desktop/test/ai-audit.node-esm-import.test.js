import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
import { AIAuditRecorder as RecorderFromTs } from "../../../packages/ai-audit/src/index.node.ts";

test("ai-audit is importable under Node ESM when executing TS sources (strip-types)", async () => {
  // This test lives under apps/desktop so `@formula/ai-audit` resolves via
  // apps/desktop/node_modules (where it's declared as a dependency).
  const mod = await import("@formula/ai-audit");

  assert.equal(typeof mod.AIAuditRecorder, "function");
  assert.equal(typeof mod.MemoryAIAuditStore, "function");
  assert.equal(typeof RecorderFromTs, "function");

  const store = new mod.MemoryAIAuditStore();
  const recorder = new mod.AIAuditRecorder({
    store,
    session_id: "s1",
    mode: "chat",
    input: { workbook_id: "w1" },
    model: "test-model",
  });

  recorder.recordToolCall({ name: "read_range", parameters: { range: "A1:A1" } });
  await recorder.finalize();

  const entries = await store.listEntries({ session_id: "s1" });
  assert.equal(entries.length, 1);
  assert.equal(entries[0]?.session_id, "s1");
});

