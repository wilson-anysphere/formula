import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import test from "node:test";

// Include explicit `.ts` import specifiers so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
//
// Note: `@formula/marketplace-shared` is a workspace package that can be missing in
// some cached/stale installs (agent sandboxes, CI caches). When it's not resolvable,
// importing `WebExtensionManager` will fail with `ERR_MODULE_NOT_FOUND`.
//
// Check resolvability relative to the `packages/extension-marketplace` package boundary
// (where pnpm links workspace deps), rather than relative to this test file.
const require = createRequire(import.meta.url);
const extensionMarketplaceDir = fileURLToPath(new URL("../packages/extension-marketplace", import.meta.url));
let hasWorkspaceDeps = true;
for (const specifier of ["@formula/marketplace-shared/package.json", "@formula/extension-host/package.json"]) {
  try {
    require.resolve(specifier, { paths: [extensionMarketplaceDir] });
  } catch {
    hasWorkspaceDeps = false;
    break;
  }
}

test("extension-marketplace TS sources are importable under Node ESM (strip-types)", { skip: !hasWorkspaceDeps }, async () => {
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
});
