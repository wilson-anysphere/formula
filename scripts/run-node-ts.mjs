import { spawn, spawnSync } from "node:child_process";
import { rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(import.meta.url);

const rawArgs = process.argv.slice(2);
if (rawArgs[0] === "--") rawArgs.shift();

const entry = rawArgs[0];
const entryArgs = rawArgs.slice(1);

if (!entry) {
  console.error("Usage: node scripts/run-node-ts.mjs <entrypoint> [...args]");
  process.exit(2);
}

// Resolve relative entrypoints from the caller's cwd (not the repo root), but always
// execute the child process with `cwd=repoRoot` so workspace resolution behaves
// consistently.
const entryPath = path.isAbsolute(entry) ? entry : path.resolve(process.cwd(), entry);

const baseNodeArgs = ["--no-warnings"];
const tsLoaderArgs = resolveTypeScriptLoaderArgs();
if (tsLoaderArgs.length > 0) {
  baseNodeArgs.push(...tsLoaderArgs);
} else {
  const builtInTypeScript = getBuiltInTypeScriptSupport();
  if (!builtInTypeScript.enabled) {
    console.error(
      "TypeScript execution is not available in this environment.\n" +
        "- Install dependencies (to enable the TypeScript transpile loader), or\n" +
        "- Use a Node version with built-in TypeScript support.",
    );
    process.exit(1);
  }

  baseNodeArgs.push(...builtInTypeScript.args);

  // When using Node's built-in TS support, also install the `.js` -> `.ts` resolver
  // loader so bundler-style imports work (TypeScript ESM convention).
  const loaderUrl = pathToFileURL(path.join(repoRoot, "scripts", "resolve-ts-imports-loader.mjs")).href;
  baseNodeArgs.push(...resolveNodeLoaderArgs(loaderUrl));
}

const nodeArgs = [...baseNodeArgs, entryPath, ...entryArgs];
const child = spawn(process.execPath, nodeArgs, { stdio: "inherit", cwd: repoRoot });

child.on("exit", (code, signal) => {
  if (signal) {
    console.error(`node exited with signal ${signal}`);
    process.exit(1);
  }
  process.exit(code ?? 1);
});

child.on("error", (err) => {
  console.error(err);
  process.exit(1);
});

function resolveTypeScriptLoaderArgs() {
  // When `typescript` is available, prefer a real TS->JS transpile loader over Node's
  // "strip-only" TS support:
  // - strip-only mode rejects TS runtime features like parameter properties and enums
  // - `.tsx` (JSX) requires transpilation
  try {
    require.resolve("typescript", { paths: [repoRoot] });
  } catch {
    return [];
  }

  const loaderUrl = new URL("./resolve-ts-loader.mjs", import.meta.url).href;
  return resolveNodeLoaderArgs(loaderUrl);
}

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
    const registerScript = `import { register } from "node:module"; register(${JSON.stringify(loaderUrl)});`;
    const dataUrl = `data:text/javascript;base64,${Buffer.from(registerScript, "utf8").toString("base64")}`;
    return ["--import", dataUrl];
  }

  if (allowedFlags.has("--loader")) return ["--loader", loaderUrl];
  if (allowedFlags.has("--experimental-loader")) return [`--experimental-loader=${loaderUrl}`];
  return [];
}
