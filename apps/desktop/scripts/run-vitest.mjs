import { spawn } from "node:child_process";

// `pnpm -C apps/desktop vitest …` runs from within `apps/desktop/`, but some
// callsites pass file paths rooted at the repo (e.g. `apps/desktop/src/...`).
// Normalize those to paths relative to the desktop package so Vitest can find them.
const PREFIX = "apps/desktop/";
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
  if (arg.startsWith(PREFIX)) return arg.slice(PREFIX.length);
  return arg;
});

const child = spawn("vitest", normalizedArgs, {
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
