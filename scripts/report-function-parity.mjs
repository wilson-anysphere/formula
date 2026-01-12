import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const catalogPath = path.join(repoRoot, "shared", "functionCatalog.json");
const ftabPath = path.join(repoRoot, "crates", "formula-biff", "src", "ftab.rs");

const TOP_N = 50;

function normalizeName(name) {
  return name.trim().toUpperCase();
}

/**
 * @param {string} raw
 * @returns {string[]}
 */
function extractFtabNames(raw) {
  const match = raw.match(/pub const FTAB:\s*\[&str;\s*\d+\]\s*=\s*\[([\s\S]*?)\n\];/m);
  if (!match) {
    throw new Error(
      `failed to locate \`pub const FTAB: [&str; N] = [ ... ];\` in ${path.relative(repoRoot, ftabPath)}`
    );
  }

  const block = match[1];
  const names = [];
  const stringLiteral = /"([^"]*)"/g;
  for (;;) {
    const m = stringLiteral.exec(block);
    if (!m) break;
    names.push(m[1]);
  }

  return names;
}

/**
 * @param {string[]} names
 * @param {number} limit
 */
function printNameList(names, limit) {
  const top = names.slice(0, limit);
  for (const name of top) {
    console.log(`  - ${name}`);
  }
  if (names.length > limit) {
    console.log(`  … (${names.length - limit} more)`);
  }
}

/** @type {any} */
let catalogParsed;
try {
  catalogParsed = JSON.parse(await readFile(catalogPath, "utf8"));
} catch (err) {
  throw new Error(`failed to read/parse ${path.relative(repoRoot, catalogPath)}: ${err}`);
}

if (!catalogParsed || typeof catalogParsed !== "object" || !Array.isArray(catalogParsed.functions)) {
  throw new Error(
    `${path.relative(repoRoot, catalogPath)} did not match expected shape: { functions: [...] }`
  );
}

const catalogNamesRaw = catalogParsed.functions
  .map((entry) => (entry && typeof entry.name === "string" ? normalizeName(entry.name) : ""))
  .filter((name) => name.length > 0);

const catalogNameSet = new Set(catalogNamesRaw);

const ftabRaw = await readFile(ftabPath, "utf8");
const ftabNamesRaw = extractFtabNames(ftabRaw)
  .map((name) => normalizeName(name))
  .filter((name) => name.length > 0);

const ftabNameSet = new Set(ftabNamesRaw);

const catalogTotal = catalogNamesRaw.length;
const ftabNonEmptyTotal = ftabNamesRaw.length;
const matchingCatalogCount = catalogNamesRaw.filter((name) => ftabNameSet.has(name)).length;

const ftabMissingFromCatalog = Array.from(ftabNameSet)
  .filter((name) => !catalogNameSet.has(name))
  .sort();

const catalogNotInFtab = Array.from(catalogNameSet)
  .filter((name) => !ftabNameSet.has(name))
  .sort();

console.log("Function parity report (catalog ↔ BIFF FTAB)");
console.log("");
console.log(`Catalog functions (${path.relative(repoRoot, catalogPath)}): ${catalogTotal}`);
console.log(`FTAB functions (${path.relative(repoRoot, ftabPath)}): ${ftabNonEmptyTotal}`);
console.log(`Catalog ∩ FTAB (case-insensitive name match): ${matchingCatalogCount}`);
console.log(`FTAB \\ Catalog (missing from catalog): ${ftabMissingFromCatalog.length}`);
console.log(`Catalog \\ FTAB (not present in FTAB): ${catalogNotInFtab.length}`);
console.log("");

console.log(`FTAB \\ Catalog (missing from catalog): ${ftabMissingFromCatalog.length}`);
printNameList(ftabMissingFromCatalog, TOP_N);
console.log("");

console.log(`Catalog \\ FTAB (not present in FTAB): ${catalogNotInFtab.length}`);
printNameList(catalogNotInFtab, TOP_N);
