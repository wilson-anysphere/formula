import { spawn, spawnSync } from "node:child_process";
import { rmSync, writeFileSync } from "node:fs";
import { readdir, readFile } from "node:fs/promises";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

/**
 * Node's `--test` runner started detecting TypeScript test files (`*.test.ts`) once
 * TypeScript stripping support landed. `apps/desktop` uses Vitest for `.test.ts`
 * suites, while `test:node` is intended to run only `node:test` suites written in
 * JavaScript.
 *
 * Run an explicit list of test files so `pnpm -C apps/desktop test:node` stays
 * stable across Node.js versions.
 */

const desktopRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const repoRoot = path.resolve(desktopRoot, "../..");
const require = createRequire(import.meta.url);

const testDir = path.normalize(fileURLToPath(new URL("../test/", import.meta.url)));
const clipboardTestDir = path.normalize(fileURLToPath(new URL("../src/clipboard/__tests__/", import.meta.url)));

/** @type {string[]} */
const files = [];
await collectTests(testDir, files);
await collectTests(clipboardTestDir, files);
files.sort((a, b) => a.localeCompare(b));

const tsLoaderArgs = resolveTypeScriptLoaderArgs();
const builtInTypeScript = getBuiltInTypeScriptSupport();
const canExecuteTypeScript = tsLoaderArgs.length > 0 || builtInTypeScript.enabled;

// Node's built-in "strip types" support can execute `.ts` modules, but does not support
// `.tsx` (JSX) without a real transpile loader.
const canExecuteTsx = tsLoaderArgs.length > 0;

let runnableFiles = files;
let typeScriptFilteredCount = 0;
let typeScriptTsxFilteredCount = 0;
if (!canExecuteTypeScript) {
  runnableFiles = await filterTypeScriptImportTests(files, ["ts", "tsx"]);
  typeScriptFilteredCount = files.length - runnableFiles.length;
} else if (!canExecuteTsx) {
  runnableFiles = await filterTypeScriptImportTests(files, ["tsx"]);
  typeScriptTsxFilteredCount = files.length - runnableFiles.length;
}

if (runnableFiles.length !== files.length) {
  const skipped = files.length - runnableFiles.length;
  /** @type {string[]} */
  const reasons = [];
  if (typeScriptFilteredCount > 0) {
    reasons.push(`${typeScriptFilteredCount} import TypeScript modules (TypeScript execution not available)`);
  }
  if (typeScriptTsxFilteredCount > 0) {
    reasons.push(`${typeScriptTsxFilteredCount} import .tsx modules (TSX execution not available)`);
  }
  const suffix = reasons.length > 0 ? ` (${reasons.join("; ")})` : "";
  console.log(`Skipping ${skipped} node:test file(s) that can't run in this environment${suffix}.`);
}

if (runnableFiles.length === 0) {
  console.log("No node:test files found.");
  process.exit(0);
}

const baseNodeArgs = ["--no-warnings"];
if (tsLoaderArgs.length > 0) {
  baseNodeArgs.push(...tsLoaderArgs);
} else if (builtInTypeScript.enabled) {
  baseNodeArgs.push(...builtInTypeScript.args);
  const loaderUrl = pathToFileURL(path.join(repoRoot, "scripts", "resolve-ts-imports-loader.mjs")).href;
  baseNodeArgs.push(...resolveNodeLoaderArgs(loaderUrl));
}

// Keep node:test parallelism conservative; some suites start background services and
// in CI/agent environments we can hit process/thread limits if too many test files
// run in parallel. Allow opting into higher parallelism via FORMULA_NODE_TEST_CONCURRENCY.
const parsedConcurrency = Number.parseInt(process.env.FORMULA_NODE_TEST_CONCURRENCY ?? "", 10);
const concurrency = Number.isFinite(parsedConcurrency) && parsedConcurrency > 0 ? parsedConcurrency : 1;
const nodeArgs = [...baseNodeArgs, `--test-concurrency=${concurrency}`, "--test", ...runnableFiles];
const child = spawn(process.execPath, nodeArgs, { stdio: "inherit" });
child.on("exit", (code, signal) => {
  if (signal) {
    console.error(`node:test exited with signal ${signal}`);
    process.exit(1);
  }
  process.exit(code ?? 1);
});

/**
 * @param {string} dir
 * @param {string[]} out
 * @returns {Promise<void>}
 */
async function collectTests(dir, out) {
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".test.js")) continue;
    out.push(path.join(dir, entry.name));
  }
}

/**
 * @param {string[]} files
 * @param {("ts" | "tsx")[]} extensions
 */
async function filterTypeScriptImportTests(files, extensions = ["ts", "tsx"]) {
  /** @type {string[]} */
  const out = [];
  const extGroup = extensions.join("|");
  const tsImportRe = new RegExp(
    `from\\s+["'][^"']+\\.(${extGroup})["']|import\\(\\s*["'][^"']+\\.(${extGroup})["']\\s*\\)`,
  );
  for (const file of files) {
    const text = await readFile(file, "utf8").catch(() => "");
    if (tsImportRe.test(text)) continue;
    out.push(file);
  }
  return out;
}

function resolveTypeScriptLoaderArgs() {
  try {
    require.resolve("typescript", { paths: [desktopRoot, repoRoot] });
  } catch {
    return [];
  }

  const loaderUrl = pathToFileURL(path.join(repoRoot, "scripts", "resolve-ts-loader.mjs")).href;
  return resolveNodeLoaderArgs(loaderUrl);
}

function getBuiltInTypeScriptSupport() {
  // Prefer explicit flag support when available (older Node versions require it).
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

/**
 * Resolve Node CLI flags to install an ESM loader.
 *
 * Prefer the newer `register()` API when available (via `--import`), since Node is
 * actively deprecating/removing the older `--experimental-loader` mechanism.
 *
 * @param {string} loaderUrl absolute file:// URL
 * @returns {string[]}
 */
function resolveNodeLoaderArgs(loaderUrl) {
  const allowedFlags =
    process.allowedNodeEnvironmentFlags && typeof process.allowedNodeEnvironmentFlags.has === "function"
      ? process.allowedNodeEnvironmentFlags
      : new Set();

  // `--import` exists before `module.register()` did, so gate on both.
  let supportsRegister = false;
  try {
    supportsRegister = typeof require("node:module")?.register === "function";
  } catch {
    supportsRegister = false;
  }

  if (supportsRegister && allowedFlags.has("--import")) {
    const registerScript = `import { register } from \"node:module\"; register(${JSON.stringify(loaderUrl)});`;
    const dataUrl = `data:text/javascript;base64,${Buffer.from(registerScript, "utf8").toString("base64")}`;
    return ["--import", dataUrl];
  }

  if (allowedFlags.has("--loader")) return ["--loader", loaderUrl];
  if (allowedFlags.has("--experimental-loader")) return [`--experimental-loader=${loaderUrl}`];
  return [];
}
