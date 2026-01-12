import { spawn } from "node:child_process";
import { readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const testDirUrl = new URL("../test/", import.meta.url);
const testDir = path.normalize(fileURLToPath(testDirUrl));

const entries = await readdir(testDir, { withFileTypes: true });
const files = entries
  .filter((entry) => entry.isFile() && entry.name.endsWith(".test.js"))
  .map((entry) => path.join(testDir, entry.name))
  .sort((a, b) => a.localeCompare(b));

if (files.length === 0) {
  console.log("No node:test files found.");
  process.exit(0);
}

// Keep node:test parallelism conservative; some suites start background services.
const nodeArgs = ["--no-warnings", "--test-concurrency=2", "--test", ...files];
const child = spawn(process.execPath, nodeArgs, { stdio: "inherit" });

child.on("exit", (code, signal) => {
  if (signal) {
    console.error(`node:test exited with signal ${signal}`);
    process.exit(1);
  }
  process.exit(code ?? 1);
});
