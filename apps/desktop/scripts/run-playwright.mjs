import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

// `pnpm -C apps/desktop test:e2e -- ...` forwards a literal `"--"` argument.
// Playwright treats `--` as an option-parsing sentinel, so any subsequent flags (e.g. `--grep`)
// are interpreted as positional file patterns instead of options. Strip it so developers can
// pass Playwright CLI flags through `pnpm test:e2e -- ...` reliably.
let args = process.argv.slice(2);
const delimiterIdx = args.indexOf("--");
if (delimiterIdx >= 0) {
  args = [...args.slice(0, delimiterIdx), ...args.slice(delimiterIdx + 1)];
}

// Some callsites pass file paths rooted at the repo (e.g. `apps/desktop/tests/e2e/...`).
// Normalize those to paths relative to the desktop package so Playwright can find them.
const PREFIX_POSIX = "apps/desktop/";
const PREFIX_WIN = "apps\\desktop\\";
const normalizedArgs = args.map((arg) => {
  if (typeof arg !== "string") return arg;
  if (arg.startsWith(PREFIX_POSIX)) return arg.slice(PREFIX_POSIX.length);
  if (arg.startsWith(PREFIX_WIN)) return arg.slice(PREFIX_WIN.length);
  return arg;
});

const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const localPlaywrightBin = path.join(
  packageRoot,
  "node_modules",
  ".bin",
  process.platform === "win32" ? "playwright.cmd" : "playwright",
);
const playwrightCmd = existsSync(localPlaywrightBin) ? localPlaywrightBin : "playwright";

const child = spawn(playwrightCmd, ["test", ...normalizedArgs], {
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
