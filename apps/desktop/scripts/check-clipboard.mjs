import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const desktopDir = path.resolve(__dirname, "..");
const repoRoot = path.resolve(desktopDir, "../..");

const DESKTOP_TAURI_PACKAGE = "formula-desktop-tauri";
const DESKTOP_BINARY_NAME = "formula-desktop";

function cargoTargetDir() {
  // Respect `CARGO_TARGET_DIR` if set, since some developer/CI environments override it
  // for caching. Cargo interprets relative paths relative to the working directory used
  // for `cargo build`, which this script sets to `repoRoot`.
  const targetDir = process.env.CARGO_TARGET_DIR;
  if (!targetDir || targetDir.trim() === "") return path.join(repoRoot, "target");
  return path.isAbsolute(targetDir) ? targetDir : path.join(repoRoot, targetDir);
}

function run(
  cmd,
  args,
  { cwd = repoRoot, env = process.env, shell = process.platform === "win32" } = {},
) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, {
      cwd,
      env,
      stdio: "inherit",
      // Allow running pnpm/cargo on Windows without needing `.cmd` suffixes.
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

function desktopBinaryPath() {
  const exe = process.platform === "win32" ? `${DESKTOP_BINARY_NAME}.exe` : DESKTOP_BINARY_NAME;
  return path.join(cargoTargetDir(), "release", exe);
}

async function cargoBuildDesktopBinary() {
  const cargoArgs = [
    "build",
    "-p",
    DESKTOP_TAURI_PACKAGE,
    "--features",
    "desktop",
    "--bin",
    DESKTOP_BINARY_NAME,
    "--release",
  ];

  // Prefer the repo's agent-safe wrapper on macOS/Linux.
  if (process.platform !== "win32") {
    return await run("bash", ["scripts/cargo_agent.sh", ...cargoArgs]);
  }

  // On Windows, we *try* bash first (Git Bash, MSYS2, etc). If it's not available,
  // fall back to `cargo` so the script still works in plain Windows terminals.
  try {
    return await run("bash", ["scripts/cargo_agent.sh", ...cargoArgs], { shell: false });
  } catch (err) {
    if (err && typeof err === "object" && "code" in err && err.code === "ENOENT") {
      console.warn("[clipboard-check] `bash` not found; falling back to `cargo`.");
      return await run("cargo", cargoArgs);
    }
    throw err;
  }
}

async function main() {
  console.log("[clipboard-check] Building desktop frontend (Vite)...");
  const buildFrontendCode = await run("pnpm", ["build"], { cwd: desktopDir });
  if (buildFrontendCode !== 0) process.exit(buildFrontendCode);

  console.log("[clipboard-check] Building Tauri desktop binary (release)...");
  const buildDesktopCode = await cargoBuildDesktopBinary();
  if (buildDesktopCode !== 0) process.exit(buildDesktopCode);

  const binary = desktopBinaryPath();
  if (!fs.existsSync(binary)) {
    console.error(`[clipboard-check] ERROR: expected desktop binary was not found at: ${binary}`);
    console.error(
      "[clipboard-check] Cargo build succeeded but the output binary is missing. Check the package name and --bin value.",
    );
    process.exit(1);
  }
  console.log(`[clipboard-check] Running packaged app check: ${binary}`);

  const args = ["--clipboard-smoke-check"];

  let runCmd = binary;
  let runArgs = args;
  if (process.platform === "linux") {
    // Ensure we have a virtual display in CI/headless environments.
    // xvfb-run-safe.sh is also safe to use on developer machines: if a working
    // DISPLAY is already available it will simply `exec` the command directly.
    runCmd = path.join(repoRoot, "scripts", "xvfb-run-safe.sh");
    runArgs = [binary, ...args];
  }

  const runCode = await run(runCmd, runArgs, {
    // Running the produced `.exe` directly is more reliable than going through a shell
    // (avoids quoting issues when the repo path contains spaces).
    shell: process.platform === "win32" ? false : undefined,
  });

  if (runCode !== 0) {
    if (runCode === 1) {
      console.error(
        `[clipboard-check] FAILED: clipboard smoke check reported a functional failure (exit code ${runCode}).`,
      );
    } else if (runCode === 2) {
      console.error(
        `[clipboard-check] ERROR: clipboard smoke check encountered an internal error/timeout (exit code ${runCode}).`,
      );
    } else {
      console.error(`[clipboard-check] FAILED: clipboard smoke check exited with code ${runCode}.`);
    }
  } else {
    console.log(
      "[clipboard-check] OK: packaged Tauri build can round-trip clipboard formats (text/plain + rich formats).",
    );
  }

  process.exit(runCode);
}

main().catch((err) => {
  console.error("[clipboard-check] ERROR:", err);
  process.exit(1);
});

