import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

function repoRoot() {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // packages/python-runtime/scripts -> repo root
  return path.resolve(here, "../../..");
}

async function collectPythonFiles(dir, baseDir, out) {
  const entries = await fs.readdir(dir, { withFileTypes: true });
  entries.sort((a, b) => a.name.localeCompare(b.name));
  for (const entry of entries) {
    if (entry.name === "__pycache__") continue;
    const abs = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectPythonFiles(abs, baseDir, out);
      continue;
    }
    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".py")) continue;
    const rel = path.relative(baseDir, abs).split(path.sep).join("/");
    out[rel] = await fs.readFile(abs, "utf8");
  }
}

async function main() {
  const root = repoRoot();
  const sourceRoot = path.join(root, "python", "formula_api");
  const target = path.join(root, "packages", "python-runtime", "src", "formula-files.generated.js");

  /** @type {Record<string, string>} */
  const files = {};
  await collectPythonFiles(sourceRoot, sourceRoot, files);

  const header = `// THIS FILE IS AUTO-GENERATED\n// Run: node packages/python-runtime/scripts/generate-formula-files.js\n//\n// It bundles the in-repo python/formula_api package into a JS object so the\n// Pyodide worker can install it into its virtual filesystem.\n\n`;

  const output = `${header}export const formulaFiles = ${JSON.stringify(files, null, 2)};\n`;
  await fs.writeFile(target, output, "utf8");
}

await main();
