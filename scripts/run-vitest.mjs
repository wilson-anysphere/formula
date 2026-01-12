import { spawn } from "node:child_process";
import { normalizeVitestArgs } from "./vitestArgs.mjs";

const args = normalizeVitestArgs(process.argv.slice(2));

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
