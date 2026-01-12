import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { normalizeVitestArgs } from "./vitestArgs.mjs";

let args = normalizeVitestArgs(process.argv.slice(2));

// Support `pnpm test:vitest -- run ...` (common muscle memory). We always default
// to run-once mode, so a leading `run` subcommand is redundant.
let mode = "run";
if (args[0] === "run") {
  args = args.slice(1);
} else if (args[0] === "watch" || args[0] === "dev") {
  // Allow `pnpm test:vitest watch` / `pnpm test:vitest dev` for local debugging.
  mode = "watch";
  args = args.slice(1);
}

// Prefer an explicit path to the local vitest binary so this wrapper works even
// when invoked outside of `pnpm` (where PATH might not include `node_modules/.bin`).
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

// `pnpm -C packages/foo test packages/foo/src/file.test.ts` is a common pattern when
// copy/pasting repo-rooted paths into a package-scoped test run. Normalize those
// repo-rooted paths to be relative to the current working directory so Vitest can
// resolve them correctly.
const cwdRelative = path.relative(repoRoot, process.cwd());
const isCwdInsideRepo =
  cwdRelative &&
  cwdRelative !== "." &&
  !cwdRelative.startsWith("..") &&
  !path.isAbsolute(cwdRelative);
if (isCwdInsideRepo) {
  const cwdPosix = cwdRelative.split(path.sep).join("/");
  const prefixPosix = `${cwdPosix}/`;
  const prefixWin = `${cwdPosix.replaceAll("/", "\\")}\\`;
  args = args.map((arg) => {
    if (typeof arg !== "string") return arg;
    if (arg.startsWith(prefixPosix)) return arg.slice(prefixPosix.length);
    if (arg.startsWith(prefixWin)) return arg.slice(prefixWin.length);
    return arg;
  });
}

const vitestBinName = process.platform === "win32" ? "vitest.cmd" : "vitest";
// Prefer the calling package's local binary (so workspace packages can pin their
// own Vitest major versions), falling back to the repo root binary.
const cwdVitestBin = path.join(process.cwd(), "node_modules", ".bin", vitestBinName);
const repoVitestBin = path.join(repoRoot, "node_modules", ".bin", vitestBinName);
const vitestCmd = existsSync(cwdVitestBin) ? cwdVitestBin : existsSync(repoVitestBin) ? repoVitestBin : "vitest";

const baseArgs = mode === "watch" ? ["--watch"] : ["--run"];

const child = spawn(vitestCmd, [...baseArgs, ...args], {
  stdio: "inherit",
  // On Windows, `.cmd` shims in PATH are easiest to resolve via a shell.
  shell: process.platform === "win32",
});

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }
  process.exit(code ?? 1);
});

child.on("error", (err) => {
  console.error(err);
  process.exit(1);
});
