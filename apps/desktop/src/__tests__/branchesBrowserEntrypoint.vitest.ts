import { build, stop } from "esbuild";
import { fileURLToPath } from "node:url";
import { afterAll, expect, test } from "vitest";

afterAll(() => {
  // Prevent esbuild's long-lived service process from keeping Vitest alive.
  stop();
});

test("branches browser entrypoint bundles for the browser (no Node built-ins)", async () => {
  const entry = fileURLToPath(new URL("../../../../packages/versioning/branches/src/browser.js", import.meta.url));

  const result = await build({
    entryPoints: [entry],
    bundle: true,
    platform: "browser",
    format: "esm",
    write: false,
    logLevel: "silent",
  });

  // If the browser entrypoint accidentally pulled in SQLiteBranchStore, the bundle would
  // contain Node built-ins like `node:fs`/`node:path` (and bundling would often fail).
  const bundled = result.outputFiles.map((f) => f.text).join("\n");

  // No Node built-ins should leak into the browser bundle.
  const nodeSpecifiers = Array.from(new Set(bundled.match(/node:[A-Za-z0-9_.\\/\\-]+/g) ?? [])).sort();
  expect(nodeSpecifiers).toEqual([]);

  expect(bundled).not.toContain("SQLiteBranchStore");
  expect(bundled).not.toContain("node:fs");
  expect(bundled).not.toContain("node:path");
  expect(bundled).not.toContain("node:module");
});
