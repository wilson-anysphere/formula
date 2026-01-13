import { existsSync, readFileSync, writeFileSync } from "node:fs";

function usage() {
  console.error("Usage: node scripts/merge-benchmark-results.mjs <base.json> <extra.json>");
  process.exit(1);
}

const [basePath, extraPath] = process.argv.slice(2);
if (!basePath || !extraPath) usage();

/** @returns {any[]} */
function readArray(path, { allowMissing }) {
  if (!existsSync(path)) {
    if (allowMissing) return [];
    throw new Error(`File not found: ${path}`);
  }
  const raw = readFileSync(path, "utf8");
  const parsed = JSON.parse(raw);
  if (!Array.isArray(parsed)) {
    throw new Error(`Expected JSON array in ${path}`);
  }
  return parsed;
}

const base = readArray(basePath, { allowMissing: true });
const extra = readArray(extraPath, { allowMissing: false });

/** @type {Map<string, any>} */
const byName = new Map();

for (const item of base) {
  const name = item?.name;
  if (typeof name !== "string" || name.length === 0) continue;
  byName.set(name, item);
}

for (const item of extra) {
  const name = item?.name;
  if (typeof name !== "string" || name.length === 0) continue;
  byName.set(name, item);
}

const merged = [...byName.values()].sort((a, b) => String(a.name).localeCompare(String(b.name)));
writeFileSync(basePath, JSON.stringify(merged, null, 2));

