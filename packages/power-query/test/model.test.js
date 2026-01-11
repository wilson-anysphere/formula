import assert from "node:assert/strict";
import test from "node:test";

test("model typedef module parses", async () => {
  const mod = await import("../src/model.js");
  assert.equal(typeof mod, "object");
});
