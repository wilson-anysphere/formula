import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const desktopDir = path.resolve(__dirname, "..");
const repoRoot = path.resolve(desktopDir, "../..");

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
  const exe = process.platform === "win32" ? "formula-desktop.exe" : "formula-desktop";
  return path.join(repoRoot, "target", "release", exe);
}

function desktopCargoPackageName() {
  // The desktop Tauri crate historically used the package name `desktop`. It was renamed to
  // `formula-desktop-tauri` to avoid an overly-generic workspace package name.
  //
  // We read the package name from Cargo.toml so this script works across the rename and in forks.
  const cargoTomlPath = path.join(desktopDir, "src-tauri", "Cargo.toml");
  let cargoToml;
  try {
    cargoToml = fs.readFileSync(cargoTomlPath, "utf8");
  } catch {
    return "formula-desktop-tauri";
  }

  const lines = cargoToml.split(/\r?\n/);
  let inPackage = false;
  for (const line of lines) {
    if (/^\s*\[package\]\s*$/.test(line)) {
      inPackage = true;
      continue;
    }
    if (inPackage && /^\s*\[/.test(line)) {
      inPackage = false;
    }
    if (!inPackage) continue;

    const m = line.match(/^\s*name\s*=\s*"([^"]+)"\s*$/);
    if (m) return m[1];
  }

  return "formula-desktop-tauri";
}

async function cargoBuildDesktopBinary() {
  const pkg = desktopCargoPackageName();
  const cargoArgs = [
    "build",
    "-p",
    pkg,
    "--features",
    "desktop",
    "--bin",
    "formula-desktop",
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
      console.warn("[coi-check] `bash` not found; falling back to `cargo`.");
      return await run("cargo", cargoArgs);
    }
    throw err;
  }
}

async function main() {
  console.log("[coi-check] Building desktop frontend (Vite)...");
  const buildFrontendCode = await run("pnpm", ["build"], { cwd: desktopDir });
  if (buildFrontendCode !== 0) process.exit(buildFrontendCode);

  console.log("[coi-check] Building Tauri desktop binary (release)...");
  const buildDesktopCode = await cargoBuildDesktopBinary();
  if (buildDesktopCode !== 0) process.exit(buildDesktopCode);

  const binary = desktopBinaryPath();
  if (!fs.existsSync(binary)) {
    console.error(`[coi-check] ERROR: expected desktop binary was not found at: ${binary}`);
    console.error(
      "[coi-check] Cargo build succeeded but the output binary is missing. Check the package name and --bin value.",
    );
    process.exit(1);
  }
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
