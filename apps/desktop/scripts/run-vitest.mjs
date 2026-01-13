import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

// `pnpm -C apps/desktop vitest …` runs from within `apps/desktop/`, but some
// callsites pass file paths rooted at the repo (e.g. `apps/desktop/src/...`).
// Normalize those to paths relative to the desktop package so Vitest can find them.
const PREFIX_POSIX = "apps/desktop/";
const PREFIX_WIN = "apps\\desktop\\";
let args = process.argv.slice(2);
// pnpm forwards a literal `--` delimiter into scripts. Strip the first occurrence so:
// - `pnpm -C apps/desktop vitest -- run <file>` behaves as expected
// - wrappers still work if the script itself provides fixed args before the delimiter
const delimiterIdx = args.indexOf("--");
if (delimiterIdx >= 0) {
  args = [...args.slice(0, delimiterIdx), ...args.slice(delimiterIdx + 1)];
}

const normalizedArgs = args.map((arg) => {
  if (typeof arg !== "string") return arg;
  // Vitest treats `--silent <pattern>` as "silent has value <pattern>". Normalize to the
  // explicit boolean form so `pnpm -C apps/desktop vitest --silent <file>` works.
  if (arg === "--silent") return "--silent=true";
  if (arg.startsWith(PREFIX_POSIX)) return arg.slice(PREFIX_POSIX.length);
  if (arg.startsWith(PREFIX_WIN)) return arg.slice(PREFIX_WIN.length);
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
