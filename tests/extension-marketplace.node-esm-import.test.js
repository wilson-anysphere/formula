import assert from "node:assert/strict";
import test from "node:test";

// Include explicit `.ts` import specifiers so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
//
// Note: `@formula/marketplace-shared` is a workspace package (at `./shared/`) that can be
// missing from cached/stale `node_modules` installs. The repo's node:test runner installs
// an ESM loader that resolves missing `@formula/*` imports directly from workspace source
// so this suite can still run in those environments.

test(
  "extension-marketplace MarketplaceClient TS source is importable under Node ESM when executing TS sources directly",
  async () => {
    const { MarketplaceClient, normalizeMarketplaceBaseUrl: normalizeFromClient } = await import(
      "../packages/extension-marketplace/src/MarketplaceClient.ts"
    );

    assert.equal(typeof MarketplaceClient, "function");
    assert.equal(typeof normalizeFromClient, "function");

    assert.equal(normalizeFromClient(""), "/api");
    assert.equal(normalizeFromClient("/api/"), "/api");
    assert.equal(normalizeFromClient("https://marketplace.formula.app"), "https://marketplace.formula.app/api");
  },
);

test(
  "extension-marketplace full TS sources are importable under Node ESM when executing TS sources directly",
  async () => {
    const { MarketplaceClient, WebExtensionManager, normalizeMarketplaceBaseUrl: normalizeFromIndex } = await import(
      "../packages/extension-marketplace/src/index.ts"
    );
    const { normalizeMarketplaceBaseUrl: normalizeFromClient } = await import(
      "../packages/extension-marketplace/src/MarketplaceClient.ts"
    );

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
  },
);
