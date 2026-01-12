import assert from "node:assert/strict";
import test from "node:test";

// Include explicit `.ts` import specifiers so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
import {
  MarketplaceClient,
  WebExtensionManager,
  normalizeMarketplaceBaseUrl as normalizeFromIndex,
} from "../packages/extension-marketplace/src/index.ts";
import { normalizeMarketplaceBaseUrl as normalizeFromClient } from "../packages/extension-marketplace/src/MarketplaceClient.ts";

test("extension-marketplace TS sources are importable under Node ESM (strip-types)", () => {
  assert.equal(typeof MarketplaceClient, "function");
  assert.equal(typeof WebExtensionManager, "function");
  assert.equal(typeof normalizeFromIndex, "function");
  assert.equal(typeof normalizeFromClient, "function");

  assert.equal(normalizeFromIndex(""), "/api");
  assert.equal(normalizeFromClient(""), "/api");
  assert.equal(normalizeFromIndex("/api/"), "/api");
  assert.equal(normalizeFromClient("/api/"), "/api");
  assert.equal(normalizeFromIndex("https://marketplace.formula.app"), "https://marketplace.formula.app/api");
  assert.equal(normalizeFromClient("https://marketplace.formula.app"), "https://marketplace.formula.app/api");
});

