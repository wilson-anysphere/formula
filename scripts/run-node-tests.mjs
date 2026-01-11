import { spawn, spawnSync } from "node:child_process";
import { readdir } from "node:fs/promises";
import path from "node:path";

const repoRoot = path.resolve(new URL(".", import.meta.url).pathname, "..");

/**
 * @param {string} dir
 * @param {string[]} out
 * @returns {Promise<void>}
 */
async function collect(dir, out) {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    // Skip node_modules and other generated output.
    if (entry.name === "node_modules" || entry.name === "dist" || entry.name === "coverage") continue;

    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collect(fullPath, out);
      continue;
    }

    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".test.js")) continue;
    out.push(fullPath);
  }
}

/** @type {string[]} */
const testFiles = [];
await collect(repoRoot, testFiles);
testFiles.sort();

if (testFiles.length === 0) {
  console.log("No node:test files found.");
  process.exit(0);
}

const nodeArgs = ["--no-warnings"];
if (supportsTypeStripping()) nodeArgs.push("--experimental-strip-types");
nodeArgs.push("--test", ...testFiles);

const child = spawn(process.execPath, nodeArgs, {
  stdio: "inherit",
});

child.on("exit", (code, signal) => {
  if (signal) {
    console.error(`node:test exited with signal ${signal}`);
    process.exit(1);
  }
  process.exit(code ?? 1);
});

function supportsTypeStripping() {
  const probe = spawnSync(process.execPath, ["--experimental-strip-types", "-e", "process.exit(0)"], {
    stdio: "ignore",
  });
  return probe.status === 0;
}
