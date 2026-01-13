import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const desktopDir = path.resolve(__dirname, "..");

function run(cmd, args, { cwd = desktopDir, env = process.env, shell = process.platform === "win32" } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, {
      cwd,
      env,
      stdio: "inherit",
      // Allow running pnpm on Windows without needing `.cmd` suffixes.
      shell,
    });
    child.on("error", reject);
    child.on("exit", (code, signal) => {
      if (signal) {
        reject(new Error(`${cmd} exited with signal ${signal}`));
        return;
      }
      resolve(code ?? 1);
    });
  });
}

async function main() {
  // Keep this script fast and cross-platform. We want to analyze the Vite/Rollup
  // bundle, so we invoke the build steps directly instead of going through `pnpm build`
  // (which would also run the workspace `prebuild` hook to rebuild the Rust/WASM engine).
  const env = { ...process.env, VITE_BUNDLE_ANALYZE: "1" };

  const pyodideCode = await run(process.execPath, ["scripts/ensure-pyodide-assets.mjs"], { env });
  if (pyodideCode !== 0) process.exit(pyodideCode);

  const viteCode = await run("vite", ["build"], { env });
  process.exit(viteCode);
}

main().catch((err) => {
  console.error("[build:analyze] ERROR:", err);
  process.exit(1);
});
