import { build, stop } from "esbuild";
import { fileURLToPath } from "node:url";
import { afterAll, expect, test } from "vitest";

afterAll(() => {
  // Prevent esbuild's long-lived service process from keeping Vitest alive.
  stop();
});

test("branches browser entrypoint bundles for the browser (no Node built-ins)", async () => {
  const entry = fileURLToPath(new URL("../../../../packages/versioning/branches/src/browser.js", import.meta.url));

  await expect(
    build({
      entryPoints: [entry],
      bundle: true,
      platform: "browser",
      format: "esm",
      write: false,
      logLevel: "silent",
      // YjsBranchStore uses a Node-only `node:zlib` fallback behind a runtime check.
      // Mark it external so this smoke test focuses on *accidental* Node-only imports
      // (e.g. SQLiteBranchStore pulling `node:fs`/`node:path` into the bundle).
      external: ["node:zlib"],
    }),
  ).resolves.toBeDefined();
});

