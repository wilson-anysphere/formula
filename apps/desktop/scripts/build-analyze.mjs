import { spawn } from "node:child_process";
import { existsSync, statSync } from "node:fs";
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

function usage() {
  return [
    "Build the desktop Vite bundle with rollup-plugin-visualizer enabled.",
    "",
    "Usage:",
    "  pnpm -C apps/desktop build:analyze",
    "  pnpm -C apps/desktop build:analyze:sourcemap",
    "",
    "Options:",
    "  --sourcemap   Enable Rollup sourcemaps for more accurate per-module attribution.",
    "  --help, -h    Show this help output.",
    "",
    "Any additional args are forwarded to `vite build`.",
  ].join("\n");
}

/**
 * @param {string[]} argv
 */
function parseArgs(argv) {
  let args = argv.slice();
  // pnpm forwards a literal `--` delimiter into scripts. Strip the first occurrence so
  // `pnpm build:analyze -- --mode development` behaves as expected.
  const delimiterIdx = args.indexOf("--");
  if (delimiterIdx >= 0) {
    args = [...args.slice(0, delimiterIdx), ...args.slice(delimiterIdx + 1)];
  }

  /** @type {{ sourcemap: boolean, help: boolean, viteArgs: string[] }} */
  const out = { sourcemap: false, help: false, viteArgs: [] };
  for (const arg of args) {
    if (!arg) continue;
    if (arg === "--help" || arg === "-h") {
      out.help = true;
      continue;
    }
    if (arg === "--sourcemap") {
      out.sourcemap = true;
      continue;
    }
    out.viteArgs.push(arg);
  }
  return out;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    console.log(usage());
    process.exit(0);
  }

  // Keep this script cross-platform and resilient. We set the env var and run the
  // Vite build directly (mirroring the `build` script body) to generate bundle stats,
  // without relying on the workspace `prebuild` hook.
  const env = {
    ...process.env,
    VITE_BUNDLE_ANALYZE: "1",
    ...(args.sourcemap ? { VITE_BUNDLE_ANALYZE_SOURCEMAP: "1" } : {}),
  };

  const pyodideCode = await run(process.execPath, ["scripts/ensure-pyodide-assets.mjs"], { env });
  if (pyodideCode !== 0) process.exit(pyodideCode);

  const viteCode = await run("vite", ["build", ...args.viteArgs], { env });
  if (viteCode !== 0) process.exit(viteCode);

  const distDir = path.join(desktopDir, "dist");
  const expected = ["bundle-stats.html", "bundle-stats.json"];
  if (env.VITE_BUNDLE_ANALYZE_SOURCEMAP === "1") {
    expected.push("bundle-stats-sourcemap.html");
  }

  /** @type {string[]} */
  const missing = [];
  for (const file of expected) {
    const absPath = path.join(distDir, file);
    if (!existsSync(absPath)) {
      missing.push(file);
      continue;
    }
    try {
      if (statSync(absPath).size === 0) missing.push(file);
    } catch {
      missing.push(file);
    }
  }

  if (missing.length > 0) {
    console.error("[build:analyze] Expected bundle analysis report files were not generated:");
    for (const file of missing) console.error(`- ${path.join("dist", file)}`);
    console.error("");
    console.error(
      "[build:analyze] This usually means rollup-plugin-visualizer failed to load. Ensure dependencies are installed (pnpm install).",
    );
    process.exit(1);
  }

  process.exit(0);
}

main().catch((err) => {
  console.error("[build:analyze] ERROR:", err);
  process.exit(1);
});
