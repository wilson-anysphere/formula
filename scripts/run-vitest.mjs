import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { normalizeVitestArgs } from "./vitestArgs.mjs";

const args = normalizeVitestArgs(process.argv.slice(2));

// Prefer an explicit path to the local vitest binary so this wrapper works even
// when invoked outside of `pnpm` (where PATH might not include `node_modules/.bin`).
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const localVitestBin = path.join(
  repoRoot,
  "node_modules",
  ".bin",
  process.platform === "win32" ? "vitest.cmd" : "vitest",
);
const vitestCmd = existsSync(localVitestBin) ? localVitestBin : "vitest";

const child = spawn(vitestCmd, ["run", ...args], {
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
