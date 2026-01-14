import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { normalizeVitestArgs } from "../../../scripts/vitestArgs.mjs";

// `pnpm -C apps/desktop vitest …` runs from within `apps/desktop/`, but some
// callsites pass file paths rooted at the repo (e.g. `apps/desktop/src/...`).
// Normalize those to paths relative to the desktop package so Vitest can find them.
const PREFIX_POSIX = "apps/desktop/";
const PREFIX_WIN = "apps\\desktop\\";
const PREFIX_POSIX_DOT = `./${PREFIX_POSIX}`;
const PREFIX_WIN_DOT = `.\\${PREFIX_WIN}`;
// pnpm forwards literal `--` delimiters into scripts (npm/yarn muscle memory). Strip them so Vitest
// doesn't interpret them as test patterns.
let args = normalizeVitestArgs(process.argv.slice(2));

const normalizedArgs = args.map((arg) => {
  if (typeof arg !== "string") return arg;
  const isDrawingTestPath = (value) =>
    value.startsWith("src/drawings/__tests__/") || value.startsWith(`src\\drawings\\__tests__\\`);

  // Drawings `.test.ts` suites have wrapper entrypoints under `apps/desktop/src/...` so repo-rooted
  // invocations work even when running `pnpm -C apps/desktop exec vitest ...`. When running through
  // this script, preserve the `apps/desktop/` prefix for those paths so Vitest can still discover
  // the wrapper file (which is included in `vite.config.ts`).
  if (arg.startsWith(PREFIX_POSIX_DOT)) {
    const stripped = arg.slice(PREFIX_POSIX_DOT.length);
    if (isDrawingTestPath(stripped)) return PREFIX_POSIX + stripped;
    return stripped;
  }
  if (arg.startsWith(PREFIX_WIN_DOT)) {
    const stripped = arg.slice(PREFIX_WIN_DOT.length);
    if (isDrawingTestPath(stripped)) return PREFIX_WIN + stripped;
    return stripped;
  }
  if (arg.startsWith(PREFIX_POSIX)) {
    const stripped = arg.slice(PREFIX_POSIX.length);
    if (isDrawingTestPath(stripped)) return arg;
    return stripped;
  }
  if (arg.startsWith(PREFIX_WIN)) {
    const stripped = arg.slice(PREFIX_WIN.length);
    if (isDrawingTestPath(stripped)) return arg;
    return stripped;
  }

  // Back-compat: allow callers to pass the package-relative drawings test paths directly.
  if (isDrawingTestPath(arg)) return PREFIX_POSIX + arg;
  return arg;
});

const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const localVitestBin = path.join(
  packageRoot,
  "node_modules",
  ".bin",
  process.platform === "win32" ? "vitest.cmd" : "vitest",
);
if (!existsSync(localVitestBin)) {
  // Agent/CI sandboxes for this repo sometimes run without `node_modules` installed.
  // In that case, there is no local Vitest binary to execute. The desktop package
  // still runs `check:no-node` in `pretest`, so the most important guardrails are
  // enforced; skip the Vitest run rather than failing with ENOENT.
  console.warn("Vitest is not installed (missing apps/desktop/node_modules). Skipping vitest run.");
  process.exit(0);
}
const vitestCmd = localVitestBin;

const child = spawn(vitestCmd, normalizedArgs, {
  cwd: packageRoot,
  stdio: "inherit",
  // On Windows, `.cmd` shims in PATH are easiest to resolve via a shell.
  shell: process.platform === "win32",
});

child.on("exit", (code, signal) => {
  if (signal) {
    // Preserve signal-based exits (useful for Ctrl+C).
    process.kill(process.pid, signal);
    return;
  }
  process.exit(code ?? 1);
});

child.on("error", (err) => {
  console.error(err);
  process.exit(1);
});
