import { spawn } from "node:child_process";

// pnpm forwards script args without requiring a `--` delimiter, but if callers *do*
// include one (npm/yarn muscle memory), pnpm forwards the literal `--` through to the
// script. Vitest treats a bare `--` as a test pattern, which can accidentally cause the
// full suite to run. Strip it so `pnpm test:vitest -- <file>` behaves as expected.
let args = process.argv.slice(2);
if (args[0] === "--") args = args.slice(1);

const child = spawn("vitest", ["run", ...args], {
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

