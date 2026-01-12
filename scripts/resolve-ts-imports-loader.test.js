import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import path from "node:path";
import test from "node:test";
import { pathToFileURL } from "node:url";

// Include an explicit `.ts` specifier so `scripts/run-node-tests.mjs` can skip this
// suite when TypeScript execution isn't available (no transpile loader and no
// `--experimental-strip-types` support).
import { valueFromBar } from "./__fixtures__/resolve-ts-imports/foo.ts";
import { valueFromBarExtensionless } from "./__fixtures__/resolve-ts-imports/foo-extensionless.ts";

test("node:test runner resolves bundler-style + extensionless TS specifiers", () => {
  assert.equal(valueFromBar(), 42);
  assert.equal(valueFromBarExtensionless(), 42);
});

function supportsTypeStripping() {
  const probe = spawnSync(process.execPath, ["--experimental-strip-types", "-e", "process.exit(0)"], {
    stdio: "ignore",
  });
  return probe.status === 0;
}

test(
  "resolve-ts-imports-loader works under --experimental-strip-types (no TypeScript dependency)",
  { skip: !supportsTypeStripping() },
  () => {
    const repoRoot = path.resolve(new URL(".", import.meta.url).pathname, "..");
    const loaderUrl = pathToFileURL(path.join(repoRoot, "scripts", "resolve-ts-imports-loader.mjs")).href;
    const child = spawnSync(
      process.execPath,
      [
        "--no-warnings",
        "--experimental-strip-types",
        "--loader",
        loaderUrl,
        "--input-type=module",
        "-e",
        [
          'import { valueFromBar } from "./scripts/__fixtures__/resolve-ts-imports/foo.ts";',
          'import { valueFromBarExtensionless } from "./scripts/__fixtures__/resolve-ts-imports/foo-extensionless.ts";',
          "if (valueFromBar() !== 42) process.exit(1);",
          "if (valueFromBarExtensionless() !== 42) process.exit(1);",
        ].join("\n"),
      ],
      { cwd: repoRoot, encoding: "utf8" },
    );

    assert.equal(
      child.status,
      0,
      `child node process failed (exit ${child.status})\nstdout:\n${child.stdout}\nstderr:\n${child.stderr}`,
    );
  },
);
