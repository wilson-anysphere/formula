#!/usr/bin/env node

import { readFile, writeFile, readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

/**
 * Normalizes locale function translation source JSONs under:
 *   crates/formula-engine/src/locale/data/sources/*.json
 *
 * The TSV generator (`scripts/generate-locale-function-tsv.js`) treats missing
 * translations (or translations that case-fold to the canonical name) as identity
 * mappings. Some locales were initially extracted with all identity mappings
 * included, which is noisy and confusing.
 *
 * Note: this normalizer does **not** validate that a locale source is complete.
 * If a source JSON was generated from a partial table (or a misconfigured Excel
 * install), missing entries will still silently fall back to canonical (English)
 * spellings in the generated TSVs. For `es-ES` in particular, sources should be
 * extracted from a real Excel install via:
 *   tools/excel-oracle/extract-function-translations.ps1
 * and should cover the full function catalog.
 *
 * This script rewrites sources into a minimal deterministic form:
 * - Keep only non-identity mappings: casefold(localized) !== canonical
 * - Trim whitespace around localized values
 * - Case-fold localized values using the same logic as the TSV generator
 * - Sort keys deterministically
 *
 * Usage:
 *   node scripts/normalize-locale-function-sources.js
 *   node scripts/normalize-locale-function-sources.js --check
 */

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/**
 * Normalize newlines so `--check` is stable across platforms / git autocrlf settings.
 * @param {string} value
 */
function normalizeNewlines(value) {
  return value.replace(/\r\n/g, "\n");
}

/**
 * Locale translation uses case-insensitive identifier matching that behaves like Excel.
 * Mirror the TSV generator's normalization by Unicode-aware uppercasing.
 *
 * Note: JS `toUpperCase()` is locale-insensitive; this is intentional to keep output stable.
 * @param {string} ident
 */
function casefoldIdent(ident) {
  return ident.toUpperCase();
}

/**
 * @param {any} value
 * @returns {value is Record<string, unknown>}
 */
function isPlainObject(value) {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function isCatalogShape(value) {
  return value && typeof value === "object" && Array.isArray(value.functions);
}

/**
 * @param {string[]} argv
 * @returns {{ check: boolean }}
 */
function parseArgs(argv) {
  const args = [...argv];
  let check = false;

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === "--check") {
      check = true;
      continue;
    }
    if (arg === "--help" || arg === "-h") {
      console.log(`Usage:
  node scripts/normalize-locale-function-sources.js [--check]

Options:
  --check    Fail if any locale source JSON differs from the normalized form.

Notes:
  - This script only normalizes existing locale sources; it does not validate completeness.
  - Missing translations still silently fall back to canonical (English) spellings in generated TSVs.
  - For \`es-ES\`, sources should be extracted from a real Excel install via:
      tools/excel-oracle/extract-function-translations.ps1
`);
      process.exit(0);
    }
    throw new Error(`Unknown argument: ${arg}`);
  }

  return { check };
}

/**
 * @param {string} filePath
 * @returns {Promise<any>}
 */
async function readJson(filePath) {
  const raw = await readFile(filePath, "utf8");
  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error(`Failed to parse JSON ${path.relative(repoRoot, filePath)}: ${err}`);
  }
}

/**
 * @param {string[]} values
 * @returns {string[]}
 */
function sortedUnique(values) {
  /** @type {string[]} */
  const out = [];
  const seen = new Set();
  for (const v of values) {
    if (seen.has(v)) continue;
    seen.add(v);
    out.push(v);
  }
  out.sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
  return out;
}

/**
 * @param {string} filePath
 * @param {Set<string>} canonicalSet
 * @returns {{ normalized: string; removedIdentityCount: number; keptCount: number; }}
 */
async function normalizeLocaleSource(filePath, canonicalSet) {
  const parsed = await readJson(filePath);
  if (!isPlainObject(parsed)) {
    throw new Error(`Locale source ${path.relative(repoRoot, filePath)} must be a JSON object`);
  }

  const keys = Object.keys(parsed);
  if (!(keys.length === 2 && keys.includes("source") && keys.includes("translations"))) {
    throw new Error(
      `Locale source ${path.relative(repoRoot, filePath)} must have shape { source: string, translations: object }`
    );
  }

  const sourceLabel = parsed.source;
  if (typeof sourceLabel !== "string" || sourceLabel.trim() === "") {
    throw new Error(
      `Locale source ${path.relative(repoRoot, filePath)} must contain a non-empty "source" string`
    );
  }

  const translationsValue = parsed.translations;
  if (!isPlainObject(translationsValue)) {
    throw new Error(
      `Locale source ${path.relative(repoRoot, filePath)} must contain a "translations" object`
    );
  }

  /** @type {Array<[string, string]>} */
  const normalizedEntries = [];
  let removedIdentityCount = 0;

  for (const [canonical, localizedRaw] of Object.entries(translationsValue)) {
    if (!canonicalSet.has(canonical)) {
      throw new Error(
        `Locale source ${path.relative(repoRoot, filePath)} contains unknown canonical function name: ${canonical}`
      );
    }
    if (typeof localizedRaw !== "string") {
      throw new Error(
        `Locale source ${path.relative(repoRoot, filePath)}: translation for ${canonical} must be a string`
      );
    }
    const trimmed = localizedRaw.trim();
    if (trimmed === "") {
      throw new Error(
        `Locale source ${path.relative(repoRoot, filePath)}: translation for ${canonical} must not be empty`
      );
    }

    const casefolded = casefoldIdent(trimmed);
    if (casefolded === canonical) {
      removedIdentityCount++;
      continue;
    }

    normalizedEntries.push([canonical, casefolded]);
  }

  normalizedEntries.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));

  /** @type {Record<string, string>} */
  const normalizedTranslations = {};
  for (const [canonical, localized] of normalizedEntries) {
    normalizedTranslations[canonical] = localized;
  }

  const normalizedObj = {
    source: sourceLabel,
    translations: normalizedTranslations,
  };

  return {
    normalized: JSON.stringify(normalizedObj, null, 2) + "\n",
    removedIdentityCount,
    keptCount: normalizedEntries.length,
  };
}

const { check } = parseArgs(process.argv.slice(2));

const catalogPath = path.join(repoRoot, "shared", "functionCatalog.json");
const catalog = await readJson(catalogPath);
if (!isCatalogShape(catalog)) {
  throw new Error(
    `Function catalog ${path.relative(repoRoot, catalogPath)} did not match expected shape: { functions: [...] }`
  );
}

const canonicalFunctionsSorted = sortedUnique(
  catalog.functions.map((entry) => {
    if (!entry || typeof entry.name !== "string" || entry.name.length === 0) {
      throw new Error(`Function catalog contained invalid function entry: ${JSON.stringify(entry)}`);
    }
    return entry.name;
  })
);
const canonicalSet = new Set(canonicalFunctionsSorted);

const localeSourceDir = path.join(repoRoot, "crates", "formula-engine", "src", "locale", "data", "sources");
const entries = await readdir(localeSourceDir, { withFileTypes: true });

const localeIds = entries
  .filter((entry) => entry.isFile() && entry.name.endsWith(".json"))
  .map((entry) => entry.name.replace(/\.json$/u, ""))
  .sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));

if (localeIds.length === 0) {
  throw new Error(
    `No locale source JSON files found under ${path.relative(repoRoot, localeSourceDir)} (expected e.g. de-DE.json)`
  );
}

let hadMismatch = false;
for (const localeId of localeIds) {
  const filePath = path.join(localeSourceDir, `${localeId}.json`);
  const { normalized, removedIdentityCount, keptCount } = await normalizeLocaleSource(
    filePath,
    canonicalSet
  );
  const existing = await readFile(filePath, "utf8");

  if (check) {
    if (normalizeNewlines(existing) !== normalizeNewlines(normalized)) {
      hadMismatch = true;
      console.error(
        `Locale source JSON mismatch: ${path.relative(repoRoot, filePath)} (run node scripts/normalize-locale-function-sources.js to update)`
      );
    }
  } else {
    if (normalizeNewlines(existing) !== normalizeNewlines(normalized)) {
      await writeFile(filePath, normalized, "utf8");
      console.log(
        `Wrote ${path.relative(repoRoot, filePath)} (${keptCount} entries; removed ${removedIdentityCount} identity mappings)`
      );
    } else {
      console.log(
        `Up-to-date ${path.relative(repoRoot, filePath)} (${keptCount} entries; removed ${removedIdentityCount} identity mappings)`
      );
    }
  }
}

if (check) {
  if (hadMismatch) {
    process.exitCode = 1;
  } else {
    console.log("Locale source JSONs are normalized.");
  }
}
