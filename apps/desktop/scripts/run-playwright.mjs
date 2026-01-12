import { spawn } from "node:child_process";

// `pnpm -C apps/desktop test:e2e -- ...` currently forwards a literal `"--"` argument.
// Playwright treats `--` as an option-parsing sentinel, so any subsequent flags (e.g. `--grep`)
// are interpreted as positional file patterns instead of options. Strip it so developers can
// pass Playwright CLI flags through `pnpm test:e2e -- ...` reliably.
let args = process.argv.slice(2);
if (args[0] === "--") args = args.slice(1);

// Some callsites pass file paths rooted at the repo (e.g. `apps/desktop/tests/e2e/...`).
// Normalize those to paths relative to the desktop package so Playwright can find them.
const PREFIX = "apps/desktop/";
const normalizedArgs = args.map((arg) => {
  if (typeof arg !== "string") return arg;
  if (arg.startsWith(PREFIX)) return arg.slice(PREFIX.length);
  return arg;
});

const child = spawn("playwright", ["test", ...normalizedArgs], {
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

