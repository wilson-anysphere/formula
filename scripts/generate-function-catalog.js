import { spawn } from "node:child_process";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Generates `shared/functionCatalog.json` by enumerating the Rust formula engine's
// inventory-backed registry of built-in functions.
//
// This is intentionally opt-in. CI/tests consume the committed JSON artifact so
// JavaScript/TypeScript workflows do not require compiling Rust.
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputPath = path.join(repoRoot, "shared", "functionCatalog.json");

/**
 * @param {string} command
 * @param {string[]} args
 * @returns {Promise<string>}
 */
function run(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: repoRoot,
      stdio: ["ignore", "pipe", "inherit"],
    });

    /** @type {Buffer[]} */
    const chunks = [];
    child.stdout.on("data", (chunk) => chunks.push(chunk));
    child.on("error", reject);
    child.on("close", (code) => {
      if (code !== 0) {
        reject(new Error(`${command} ${args.join(" ")} exited with code ${code}`));
        return;
      }
      resolve(Buffer.concat(chunks).toString("utf8"));
    });
  });
}

function isCatalogShape(value) {
  return value && typeof value === "object" && Array.isArray(value.functions);
}

const raw = await run("cargo", [
  "run",
  "--quiet",
  "-p",
  "formula-engine",
  "--bin",
  "function_catalog",
]);

/** @type {any} */
let parsed;
try {
  parsed = JSON.parse(raw);
} catch (err) {
  throw new Error(`Rust function_catalog output was not valid JSON: ${err}`);
}

if (!isCatalogShape(parsed)) {
  throw new Error("Rust function_catalog output did not match expected shape: { functions: [...] }");
}

for (const entry of parsed.functions) {
  if (!entry || typeof entry.name !== "string") {
    throw new Error("Rust function_catalog output contained invalid function entry");
  }
}

await mkdir(path.dirname(outputPath), { recursive: true });
await writeFile(outputPath, JSON.stringify(parsed, null, 2) + "\n", "utf8");

console.log(`Wrote ${path.relative(repoRoot, outputPath)} (${parsed.functions.length} functions)`);
