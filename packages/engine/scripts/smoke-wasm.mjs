import { accessSync, constants } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// `packages/engine/scripts/*` → repo root
const repoRoot = path.resolve(__dirname, "..", "..", "..");

const outDir = path.join(repoRoot, "packages", "engine", "pkg");
const runtimeIgnore = new Set([".gitignore", "package.json"]);

const publicTargets = [
  path.join(repoRoot, "apps", "web", "public", "engine"),
  path.join(repoRoot, "apps", "desktop", "public", "engine")
];

function assertReadable(filePath) {
  try {
    accessSync(filePath, constants.R_OK);
  } catch {
    console.error(`[formula] Missing WASM artifact: ${path.relative(repoRoot, filePath)}`);
    console.error("Run `pnpm build:wasm` first.");
    process.exit(1);
  }
}

async function collectRuntimeFiles(dir, relativeDir = "") {
  const { readdir } = await import("node:fs/promises");
  const entries = await readdir(dir, { withFileTypes: true });

  /** @type {string[]} */
  const out = [];
  for (const entry of entries) {
    if (runtimeIgnore.has(entry.name)) continue;
    if (entry.name.endsWith(".d.ts")) continue;

    const rel = relativeDir ? path.join(relativeDir, entry.name) : entry.name;
    const abs = path.join(dir, entry.name);

    if (entry.isDirectory()) {
      out.push(...(await collectRuntimeFiles(abs, rel)));
      continue;
    }

    if (!entry.isFile()) continue;
    out.push(rel);
  }
  return out;
}

// Fail fast with a helpful message when the output directory doesn't exist yet.
// This is the common case when contributors haven't run `pnpm build:wasm`.
assertReadable(outDir);

const runtimeFiles = await collectRuntimeFiles(outDir);
if (runtimeFiles.length === 0) {
  console.error("[formula] No WASM runtime files found. Run `pnpm build:wasm` first.");
  process.exit(1);
}

for (const relPath of runtimeFiles) {
  assertReadable(path.join(outDir, relPath));
}

for (const dir of publicTargets) {
  for (const relPath of runtimeFiles) {
    assertReadable(path.join(dir, relPath));
  }
}

console.log("[formula] WASM artifacts present.");
