import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const desktopDir = path.resolve(__dirname, "..");
const repoRoot = path.resolve(desktopDir, "../..");

function run(cmd, args, { cwd = repoRoot, env = process.env } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, {
      cwd,
      env,
      stdio: "inherit",
      // Allow running pnpm/cargo on Windows without needing `.cmd` suffixes.
      shell: process.platform === "win32",
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

function desktopBinaryPath() {
  const exe = process.platform === "win32" ? "formula-desktop.exe" : "formula-desktop";
  return path.join(repoRoot, "target", "release", exe);
}

async function main() {
  console.log("[coi-check] Building desktop frontend (Vite)...");
  const buildFrontendCode = await run("pnpm", ["build"], { cwd: desktopDir });
  if (buildFrontendCode !== 0) process.exit(buildFrontendCode);

  console.log("[coi-check] Building Tauri desktop binary (release)...");
  const buildDesktopCode = await run("cargo", [
    "build",
    "-p",
    "formula-desktop-tauri",
    "--features",
    "desktop",
    "--bin",
    "formula-desktop",
    "--release",
  ]);
  if (buildDesktopCode !== 0) process.exit(buildDesktopCode);

  const binary = desktopBinaryPath();
  console.log(`[coi-check] Running packaged app check: ${binary}`);

  const args = ["--cross-origin-isolation-check"];

  let runCmd = binary;
  let runArgs = args;
  if (process.platform === "linux") {
    // Ensure we have a virtual display in CI/headless environments.
    runCmd = path.join(repoRoot, "scripts", "xvfb-run-safe.sh");
    runArgs = [binary, ...args];
  }

  const runCode = await run(runCmd, runArgs);
  if (runCode !== 0) {
    console.error(
      `[coi-check] FAILED: packaged Tauri build is not cross-origin isolated (exit code ${runCode}).`,
    );
  } else {
    console.log("[coi-check] OK: packaged Tauri build is cross-origin isolated.");
  }
  process.exit(runCode);
}

main().catch((err) => {
  console.error("[coi-check] ERROR:", err);
  process.exit(1);
});

