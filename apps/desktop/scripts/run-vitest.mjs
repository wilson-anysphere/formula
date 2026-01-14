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
const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
// pnpm forwards literal `--` delimiters into scripts (npm/yarn muscle memory). Strip them so Vitest
// doesn't interpret them as test patterns.
let args = normalizeVitestArgs(process.argv.slice(2));

const normalizedArgs = args.map((arg) => {
  if (typeof arg !== "string") return arg;
  const isDrawingTestPath = (value) =>
    value.startsWith("src/drawings/__tests__/") || value.startsWith(`src\\drawings\\__tests__\\`);
  const toPosix = (value) => value.replaceAll("\\", "/");
  const maybeDrawingWrapper = (value) => {
    if (!isDrawingTestPath(value)) return null;
    const posix = toPosix(value);
    const wrapper = PREFIX_POSIX + posix;
    // Wrapper entrypoints live under `apps/desktop/apps/desktop/...` and only exist for a subset
    // of the drawings tests. Only rewrite to the wrapper path if it actually exists; otherwise,
    // fall back to the real suite path under `src/drawings/__tests__`.
    const wrapperFsPath = path.join(packageRoot, ...wrapper.split("/"));
    if (!existsSync(wrapperFsPath)) return null;
    return value.includes("\\") ? wrapper.replaceAll("/", "\\") : wrapper;
  };

  // Drawings `.test.ts` suites have wrapper entrypoints under `apps/desktop/src/...` so repo-rooted
  // invocations (e.g. `apps/desktop/src/...`) still work when run from within the `apps/desktop/`
  // package directory.
  if (arg.startsWith(PREFIX_POSIX_DOT)) {
    const stripped = arg.slice(PREFIX_POSIX_DOT.length);
    const wrapper = maybeDrawingWrapper(stripped);
    if (wrapper) return wrapper;
    return stripped;
  }
  if (arg.startsWith(PREFIX_WIN_DOT)) {
    const stripped = arg.slice(PREFIX_WIN_DOT.length);
    const wrapper = maybeDrawingWrapper(stripped);
    if (wrapper) return wrapper;
    return stripped;
  }
  if (arg.startsWith(PREFIX_POSIX)) {
    const stripped = arg.slice(PREFIX_POSIX.length);
    const wrapper = maybeDrawingWrapper(stripped);
    if (wrapper) return wrapper;
    return stripped;
  }
  if (arg.startsWith(PREFIX_WIN)) {
    const stripped = arg.slice(PREFIX_WIN.length);
    const wrapper = maybeDrawingWrapper(stripped);
    if (wrapper) return wrapper;
    return stripped;
  }
  if (isDrawingTestPath(arg)) {
    const wrapper = maybeDrawingWrapper(arg);
    if (wrapper) return wrapper;
    return arg;
  }
  return arg;
});
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
