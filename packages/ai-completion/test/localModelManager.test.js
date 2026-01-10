import assert from "node:assert/strict";
import test from "node:test";

import { LocalModelManager } from "../src/localModelManager.js";

test("LocalModelManager.initialize checks health and pulls missing models", async () => {
  const calls = [];
  const ollamaClient = {
    async health() {
      calls.push(["health"]);
      return true;
    },
    async hasModel(name) {
      calls.push(["hasModel", name]);
      return false;
    },
    async pullModel(name) {
      calls.push(["pullModel", name]);
    },
    async generate() {
      throw new Error("not used");
    },
  };

  const manager = new LocalModelManager({
    ollamaClient,
    requiredModels: ["formula-completion"],
  });

  await manager.initialize();

  assert.deepEqual(calls, [
    ["health"],
    ["hasModel", "formula-completion"],
    ["pullModel", "formula-completion"],
  ]);
});

test("LocalModelManager.complete caches generate results (LRU)", async () => {
  let generateCalls = 0;
  const ollamaClient = {
    async health() {
      return true;
    },
    async hasModel() {
      return true;
    },
    async pullModel() {},
    async generate() {
      generateCalls++;
      return { response: "COMPLETION" };
    },
  };

  const manager = new LocalModelManager({
    ollamaClient,
    requiredModels: ["formula-completion"],
    cacheSize: 10,
  });

  const first = await manager.complete("prompt", { model: "formula-completion" });
  const second = await manager.complete("prompt", { model: "formula-completion" });

  assert.equal(first, "COMPLETION");
  assert.equal(second, "COMPLETION");
  assert.equal(generateCalls, 1, "Expected generate to be called once due to caching");
});
