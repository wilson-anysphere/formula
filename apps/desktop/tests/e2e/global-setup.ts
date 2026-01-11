import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

function run(command: string, args: string[], cwd: string) {
  return new Promise<void>((resolve, reject) => {
    const child = spawn(command, args, { cwd, stdio: "inherit" });
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(new Error(`${command} ${args.join(" ")} exited with code ${code}`));
    });
  });
}

export default async function globalSetup() {
  const here = path.dirname(fileURLToPath(import.meta.url));
  const desktopRoot = path.resolve(here, "../..");
  const pnpmBin = process.platform === "win32" ? "pnpm.cmd" : "pnpm";

  // Playwright e2e uses `@formula/engine`, which depends on generated wasm-pack
  // output under `packages/engine/pkg`. Ensure the artifacts exist before the
  // dev server starts serving modules.
  try {
    await run(pnpmBin, ["-w", "smoke:wasm"], desktopRoot);
  } catch {
    await run(pnpmBin, ["-w", "build:wasm"], desktopRoot);
  }

  await run(process.execPath, ["scripts/ensure-pyodide-assets.mjs"], desktopRoot);
}
