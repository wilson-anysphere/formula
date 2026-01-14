import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { rmSync, statSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

import { resolve as resolveTsImportsLoader } from "./resolve-ts-imports-loader.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

// Include an explicit `.ts` specifier so `scripts/run-node-tests.mjs` can skip this
// suite when TypeScript execution isn't available (no transpile loader and no
// built-in TypeScript support).
import { valueFromBar } from "./__fixtures__/resolve-ts-imports/foo.ts";
import { valueFromBarExtensionless } from "./__fixtures__/resolve-ts-imports/foo-extensionless.ts";
import { valueFromDirImport } from "./__fixtures__/resolve-ts-imports/foo-dir-import.ts";
import { valueFromBarJsx } from "./__fixtures__/resolve-ts-imports/foo-jsx.ts";

test("node:test runner resolves bundler-style + extensionless + directory TS specifiers", () => {
  assert.equal(valueFromBar(), 42);
  assert.equal(valueFromBarExtensionless(), 42);
  assert.equal(valueFromDirImport(), 42);
  assert.equal(valueFromBarJsx(), 42);
});

test("resolve-ts-imports-loader resolves @formula/* workspace packages when default resolution fails", async () => {
  const miss = new Error("ERR_MODULE_NOT_FOUND");
  /** @type {any} */ (miss).code = "ERR_MODULE_NOT_FOUND";
  const failingResolve = async () => {
    throw miss;
  };

  const resolvedCollab = await resolveTsImportsLoader("@formula/collab-session?raw", { parentURL: import.meta.url }, failingResolve);
  assert.equal(resolvedCollab.shortCircuit, true);
  assert.ok(typeof resolvedCollab.url === "string" && resolvedCollab.url.includes("?raw"));

  const collabUrl = new URL(resolvedCollab.url);
  collabUrl.search = "";
  collabUrl.hash = "";
  const collabPath = fileURLToPath(collabUrl);
  assert.ok(collabPath.startsWith(repoRoot), "expected resolved file to be within the repo");
  assert.ok(
    collabPath.includes(path.join("packages", "collab", "session")),
    "expected resolved file to be under packages/collab/session",
  );
  assert.ok(statSync(collabPath).isFile(), "expected resolved workspace entrypoint to exist");

  // `@formula/marketplace-shared` is a workspace package backed by the repo `shared/` directory and
  // does not use an `exports` map. Ensure we can still resolve deep imports (e.g. `.mjs` entrypoints)
  // when the workspace link is missing from `node_modules`.
  const resolvedShared = await resolveTsImportsLoader(
    "@formula/marketplace-shared/extension-package/v2-browser.mjs#test",
    { parentURL: import.meta.url },
    failingResolve,
  );
  assert.equal(resolvedShared.shortCircuit, true);
  assert.ok(typeof resolvedShared.url === "string" && resolvedShared.url.includes("#test"));

  const sharedUrl = new URL(resolvedShared.url);
  sharedUrl.search = "";
  sharedUrl.hash = "";
  const sharedPath = fileURLToPath(sharedUrl);
  assert.ok(sharedPath.startsWith(repoRoot), "expected resolved file to be within the repo");
  assert.ok(
    sharedPath.includes(path.join("shared", "extension-package", "v2-browser.mjs")),
    "expected resolved file to be under shared/extension-package",
  );
  assert.ok(statSync(sharedPath).isFile(), "expected resolved workspace deep import to exist");
});

function getBuiltInTypeScriptSupport() {
  const flagProbe = spawnSync(process.execPath, ["--experimental-strip-types", "-e", "process.exit(0)"], {
    stdio: "ignore",
  });
  if (flagProbe.status === 0) {
    return { enabled: true, args: ["--experimental-strip-types"] };
  }

  const tmpFile = path.join(os.tmpdir(), `formula-strip-types-probe.${process.pid}.${Date.now()}.ts`);
  try {
    writeFileSync(
      tmpFile,
      [
        "export const x: number = 1;",
        "if (x !== 1) throw new Error('strip-types probe failed');",
        "",
      ].join("\n"),
      "utf8",
    );
    const fileUrl = pathToFileURL(tmpFile).href;
    const nativeProbe = spawnSync(process.execPath, ["--input-type=module", "-e", `import ${JSON.stringify(fileUrl)};`], {
      stdio: "ignore",
    });
    if (nativeProbe.status === 0) {
      return { enabled: true, args: [] };
    }
  } catch {
    // ignore
  } finally {
    rmSync(tmpFile, { force: true });
  }
  return { enabled: false, args: [] };
}

const builtInTypeScript = getBuiltInTypeScriptSupport();

function supportsRegister() {
  try {
    return typeof require("node:module")?.register === "function";
  } catch {
    return false;
  }
}

/**
 * @param {string} loaderUrl
 * @returns {string[]}
 */
function resolveNodeLoaderArgs(loaderUrl) {
  const allowedFlags =
    process.allowedNodeEnvironmentFlags && typeof process.allowedNodeEnvironmentFlags.has === "function"
      ? process.allowedNodeEnvironmentFlags
      : new Set();

  // Prefer the newer `register()` mechanism when available.
  if (supportsRegister() && allowedFlags.has("--import")) {
    const registerScript = `import { register } from "node:module"; register(${JSON.stringify(loaderUrl)});`;
    const dataUrl = `data:text/javascript;base64,${Buffer.from(registerScript, "utf8").toString("base64")}`;
    return ["--import", dataUrl];
  }

  if (allowedFlags.has("--loader")) return ["--loader", loaderUrl];
  if (allowedFlags.has("--experimental-loader")) return [`--experimental-loader=${loaderUrl}`];
  return [];
}

test(
  "resolve-ts-imports-loader works under Node built-in TypeScript execution (no TypeScript dependency)",
  { skip: !builtInTypeScript.enabled },
  () => {
    const loaderUrl = pathToFileURL(path.join(repoRoot, "scripts", "resolve-ts-imports-loader.mjs")).href;
    const child = spawnSync(
      process.execPath,
      [
        "--no-warnings",
        ...builtInTypeScript.args,
        ...resolveNodeLoaderArgs(loaderUrl),
        "--input-type=module",
        "-e",
        [
          'import { valueFromBar } from "./scripts/__fixtures__/resolve-ts-imports/foo.ts";',
          'import { valueFromBarExtensionless } from "./scripts/__fixtures__/resolve-ts-imports/foo-extensionless.ts";',
          'import { valueFromDirImport } from "./scripts/__fixtures__/resolve-ts-imports/foo-dir-import.ts";',
          'import { valueFromBarJsx } from "./scripts/__fixtures__/resolve-ts-imports/foo-jsx.ts";',
          "if (valueFromBar() !== 42) process.exit(1);",
          "if (valueFromBarExtensionless() !== 42) process.exit(1);",
          "if (valueFromDirImport() !== 42) process.exit(1);",
          "if (valueFromBarJsx() !== 42) process.exit(1);",
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
