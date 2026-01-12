import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);

// `pnpm test:node --help` should be fast; avoid running the full repo scan just to
// print usage information.
const wantsHelp = args.includes("--help") || args.includes("-h");

/**
 * @param {string} scriptRelativePath
 * @param {string[]} scriptArgs
 * @returns {Promise<{ code: number, signal: NodeJS.Signals | null }>}
 */
async function runNodeScript(scriptRelativePath, scriptArgs = []) {
  const scriptPath = path.join(repoRoot, scriptRelativePath);
  const child = spawn(process.execPath, [scriptPath, ...scriptArgs], { stdio: "inherit" });
  return await new Promise((resolve, reject) => {
    child.on("exit", (code, signal) => {
      resolve({ code: code ?? 1, signal: /** @type {NodeJS.Signals | null} */ (signal) });
    });
    child.on("error", reject);
  });
}

try {
  if (!wantsHelp) {
    const policy = await runNodeScript("scripts/check-cursor-ai-policy.mjs");
    if (policy.signal) {
      process.kill(process.pid, policy.signal);
      // Unreachable, but keeps typecheckers happy.
      process.exit(1);
    }
    if (policy.code !== 0) process.exit(policy.code);
  }

  const tests = await runNodeScript("scripts/run-node-tests.mjs", args);
  if (tests.signal) {
    process.kill(process.pid, tests.signal);
    process.exit(1);
  }
  process.exit(tests.code);
} catch (err) {
  console.error(err);
  process.exit(1);
}

