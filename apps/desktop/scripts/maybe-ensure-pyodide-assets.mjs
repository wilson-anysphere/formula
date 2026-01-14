import { readdir, rm } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

function envFlagEnabled(name) {
  const raw = process.env[name];
  if (typeof raw !== "string") return false;
  switch (raw.trim().toLowerCase()) {
    case "1":
    case "true":
    case "yes":
    case "on":
      return true;
    default:
      return false;
  }
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const pyodidePublicDir = path.resolve(__dirname, "../public/pyodide");

async function cleanBundledPyodideAssets(dir) {
  let entries;
  try {
    entries = await readdir(dir);
  } catch {
    return;
  }

  const toRemove = entries.filter((name) => name !== ".gitignore");
  if (toRemove.length === 0) return;

  await Promise.all(
    toRemove.map(async (name) => {
      await rm(path.join(dir, name), { recursive: true, force: true });
    }),
  );

  console.log(
    `Removed bundled Pyodide assets from ${path.relative(process.cwd(), dir)} (set FORMULA_BUNDLE_PYODIDE_ASSETS=1 to bundle them)`,
  );
}

async function main() {
  if (envFlagEnabled("FORMULA_BUNDLE_PYODIDE_ASSETS")) {
    await import("./ensure-pyodide-assets.mjs");
    return;
  }

  // Defense-in-depth: avoid accidentally bundling Pyodide into `dist/` if a
  // developer/CI cache previously populated `public/pyodide/...`.
  //
  // Vite copies everything under `public/` into `dist/`, so remove the entire Pyodide subtree
  // unless bundling is explicitly enabled.
  await rm(pyodidePublicDir, { recursive: true, force: true });
  // If removal failed (e.g. permission issue) and files still exist, fall back to a per-directory
  // cleanup to avoid leaving large artifacts around.
  await cleanBundledPyodideAssets(pyodidePublicDir);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
