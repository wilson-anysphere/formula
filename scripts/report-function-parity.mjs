import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const catalogPath = path.join(repoRoot, "shared", "functionCatalog.json");
const ftabPath = path.join(repoRoot, "crates", "formula-biff", "src", "ftab.rs");
const docsPath = path.join(repoRoot, "docs", "15-excel-feature-parity.md");

const TOP_N = 50;

function normalizeName(name) {
  return name.trim().toUpperCase();
}

function formatRelPath(absPath) {
  // `path.relative` uses platform separators, which makes output non-deterministic across OSes.
  // Normalize to `/` so reports are stable (and match docs snapshots) on all platforms.
  return path.relative(repoRoot, absPath).split(path.sep).join("/");
}

function buildSummaryLines({
  catalogTotal,
  ftabNonEmptyTotal,
  matchingCatalogCount,
  missingFromCatalogCount,
  catalogNotInFtabCount,
}) {
  const relCatalog = formatRelPath(catalogPath);
  const relFtab = formatRelPath(ftabPath);
  return [
    "Function parity report (catalog ↔ BIFF FTAB)",
    "",
    `Catalog functions (${relCatalog}): ${catalogTotal}`,
    `FTAB functions (${relFtab}): ${ftabNonEmptyTotal}`,
    `Catalog ∩ FTAB (case-insensitive name match): ${matchingCatalogCount}`,
    `FTAB \\ Catalog (missing from catalog): ${missingFromCatalogCount}`,
    `Catalog \\ FTAB (not present in FTAB): ${catalogNotInFtabCount}`,
  ];
}

/**
 * @param {string} raw
 * @returns {string[]}
 */
function extractFtabNames(raw) {
  const match = raw.match(/pub const FTAB:\s*\[&str;\s*\d+\]\s*=\s*\[([\s\S]*?)\n\];/m);
  if (!match) {
    throw new Error(
      `failed to locate \`pub const FTAB: [&str; N] = [ ... ];\` in ${formatRelPath(ftabPath)}`
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

const summaryLines = buildSummaryLines({
  catalogTotal,
  ftabNonEmptyTotal,
  matchingCatalogCount,
  missingFromCatalogCount: ftabMissingFromCatalog.length,
  catalogNotInFtabCount: catalogNotInFtab.length,
});

for (const line of summaryLines) {
  console.log(line);
}
console.log("");

const args = new Set(process.argv.slice(2));
if (args.has("--update-doc")) {
  const beginMarker = "<!-- BEGIN GENERATED: report-function-parity -->";
  const endMarker = "<!-- END GENERATED: report-function-parity -->";
  const rawDoc = await readFile(docsPath, "utf8");
  const begin = rawDoc.indexOf(beginMarker);
  const end = rawDoc.indexOf(endMarker);
  if (begin === -1 || end === -1 || begin > end) {
    throw new Error(
      `failed to update ${formatRelPath(docsPath)}: could not find expected markers:\n${beginMarker}\n${endMarker}`
    );
  }

  const replacementBody = `\n\`\`\`text\n${summaryLines.join("\n")}\n\`\`\`\n`;
  const updated =
    rawDoc.slice(0, begin + beginMarker.length) + replacementBody + rawDoc.slice(end);
  if (updated !== rawDoc) {
    await writeFile(docsPath, updated, "utf8");
  }
}

console.log(`FTAB \\ Catalog (missing from catalog): ${ftabMissingFromCatalog.length}`);
printNameList(ftabMissingFromCatalog, TOP_N);
console.log("");

console.log(`Catalog \\ FTAB (not present in FTAB): ${catalogNotInFtab.length}`);
printNameList(catalogNotInFtab, TOP_N);
