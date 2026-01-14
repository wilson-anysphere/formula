#!/usr/bin/env node

import { readFile, writeFile, mkdir, readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

/**
 * Generates `crates/formula-engine/src/locale/data/{de-DE,fr-FR,es-ES}.tsv`
 * from:
 * - The engine's canonical function catalog (`shared/functionCatalog.json`)
 * - Locale-specific translation sources (`crates/formula-engine/src/locale/data/sources/*.json`)
 *
 * The translation sources should come from a real Microsoft Excel install via:
 *   tools/excel-oracle/extract-function-translations.ps1
 * whenever possible (especially for `es-ES`). Missing translations silently fall back to the
 * canonical (English) name in the generated TSVs.
 *
 * The output TSVs intentionally contain exactly one entry per canonical function name.
 * Any missing translation (or translation that equals canonical) is emitted as an identity mapping.
 *
 * Usage:
 *   node scripts/generate-locale-function-tsv.js
 *   node scripts/generate-locale-function-tsv.js --check
 *
 * For reproducibility, the generator date in the TSV header is derived from:
 * - `--date YYYY-MM-DD` if provided
 * - else `SOURCE_DATE_EPOCH` (seconds) if set
 * - else `0`
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
 * @param {string[]} argv
 * @returns {{ check: boolean; dateOverride: string | null }}
 */
function parseArgs(argv) {
  const args = [...argv];
  let check = false;
  /** @type {string | null} */
  let dateOverride = null;

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === "--check") {
      check = true;
      continue;
    }
    if (arg === "--date") {
      const value = args[i + 1];
      if (!value || value.startsWith("-")) {
        throw new Error("--date requires a value in YYYY-MM-DD format");
      }
      dateOverride = value;
      i++;
      continue;
    }
    if (arg === "--help" || arg === "-h") {
      console.log(`Usage:
  node scripts/generate-locale-function-tsv.js [--check] [--date YYYY-MM-DD]

Options:
  --check    Fail if any generated TSV differs from what is committed.
  --date     Override the generation date written into TSV header comments.
             If not provided, SOURCE_DATE_EPOCH is used (seconds since Unix epoch),
             falling back to 0 for reproducible output.

Notes:
  - Locale sources live under crates/formula-engine/src/locale/data/sources/*.json.
  - Missing translations silently fall back to canonical (English) spellings in the generated TSVs.
  - After re-extracting locale sources from Excel, run:
      node scripts/normalize-locale-function-sources.js
`);
      process.exit(0);
    }

    throw new Error(`Unknown argument: ${arg}`);
  }

  return { check, dateOverride };
}

/**
 * @param {string | null} dateOverride
 * @returns {string} YYYY-MM-DD (UTC)
 */
function generationDate(dateOverride) {
  if (dateOverride != null) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateOverride)) {
      throw new Error(`--date must be in YYYY-MM-DD format; got ${JSON.stringify(dateOverride)}`);
    }
    return dateOverride;
  }

  const env = process.env.SOURCE_DATE_EPOCH;
  const epochSeconds = env != null && env !== "" ? Number.parseInt(env, 10) : 0;
  if (!Number.isFinite(epochSeconds) || Number.isNaN(epochSeconds) || epochSeconds < 0) {
    throw new Error(
      `SOURCE_DATE_EPOCH must be a non-negative integer (seconds since Unix epoch); got ${JSON.stringify(
        env
      )}`
    );
  }

  return new Date(epochSeconds * 1000).toISOString().slice(0, 10);
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

function isCatalogShape(value) {
  return value && typeof value === "object" && Array.isArray(value.functions);
}

/**
 * Locale translation uses case-insensitive identifier matching that behaves like Excel.
 * Mirror the engine's normalization by Unicode-aware uppercasing.
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

/**
 * @param {string} canonical
 * @param {string} localized
 */
function validateTsvEntry(canonical, localized) {
  if (!canonical || typeof canonical !== "string") {
    throw new Error(`Invalid canonical function name: ${JSON.stringify(canonical)}`);
  }
  if (!localized || typeof localized !== "string") {
    throw new Error(`Invalid localized function name for ${canonical}: ${JSON.stringify(localized)}`);
  }
  if (canonical.includes("\t") || canonical.includes("\n") || canonical.includes("\r")) {
    throw new Error(`Canonical function name contains invalid characters: ${JSON.stringify(canonical)}`);
  }
  if (localized.includes("\t") || localized.includes("\n") || localized.includes("\r")) {
    throw new Error(
      `Localized function name for ${canonical} contains invalid characters: ${JSON.stringify(localized)}`
    );
  }
}

/**
 * @param {{ localeId: string; sourcePath: string; outputPath: string; }} locale
 * @param {string[]} canonicalFunctionsSorted
 * @param {string} date
 * @returns {Promise<string>}
 */
async function generateTsvForLocale(locale, canonicalFunctionsSorted, date) {
  const source = await readJson(locale.sourcePath);
  if (!isPlainObject(source)) {
    throw new Error(
      `Locale source ${path.relative(repoRoot, locale.sourcePath)} must be a JSON object`
    );
  }

  const sourceLabel = source.source;
  if (typeof sourceLabel !== "string" || sourceLabel.trim() === "") {
    throw new Error(
      `Locale source ${path.relative(repoRoot, locale.sourcePath)} must contain a non-empty "source" string`
    );
  }

  const translationsValue = source.translations;
  if (!isPlainObject(translationsValue)) {
    throw new Error(
      `Locale source ${path.relative(repoRoot, locale.sourcePath)} must contain a "translations" object`
    );
  }

  /** @type {Map<string, string>} */
  const translations = new Map();
  for (const [canonical, localized] of Object.entries(translationsValue)) {
    if (typeof localized !== "string") {
      throw new Error(
        `Locale source ${path.relative(repoRoot, locale.sourcePath)}: translation for ${canonical} must be a string`
      );
    }
    const trimmed = localized.trim();
    if (trimmed === "") {
      throw new Error(
        `Locale source ${path.relative(repoRoot, locale.sourcePath)}: translation for ${canonical} must not be empty`
      );
    }
    translations.set(canonical, trimmed);
  }

  const canonicalSet = new Set(canonicalFunctionsSorted);
  for (const canonical of translations.keys()) {
    if (!canonicalSet.has(canonical)) {
      throw new Error(
        `Locale source ${path.relative(repoRoot, locale.sourcePath)} contains unknown canonical function name: ${canonical}`
      );
    }
  }

  const header = [
    "# Canonical\tLocalized",
    `# Source: ${sourceLabel.trim()}`,
    `# Generated by scripts/generate-locale-function-tsv.js on ${date} (UTC). Do not edit by hand.`,
    "# See `src/locale/data/README.md` for format + generators.",
    "",
  ];

  /** @type {string[]} */
  const lines = [...header];

  /** @type {Map<string, string>} */
  const localizedToCanonical = new Map();
  for (const canonical of canonicalFunctionsSorted) {
    const translated = translations.get(canonical);
    const localized = translated == null ? canonical : casefoldIdent(translated);
    validateTsvEntry(canonical, localized);
    const existing = localizedToCanonical.get(localized);
    if (existing != null && existing !== canonical) {
      throw new Error(
        `Locale source ${path.relative(repoRoot, locale.sourcePath)} maps multiple functions to the same localized name: ${existing} and ${canonical} -> ${localized}`
      );
    }
    localizedToCanonical.set(localized, canonical);
    lines.push(`${canonical}\t${localized}`);
  }

  return lines.join("\n") + "\n";
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

const { check, dateOverride } = parseArgs(process.argv.slice(2));
const date = generationDate(dateOverride);

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

const localeSourceDir = path.join(repoRoot, "crates", "formula-engine", "src", "locale", "data", "sources");
const localeOutputDir = path.join(repoRoot, "crates", "formula-engine", "src", "locale", "data");

const locales = await (async () => {
  /** @type {Array<{ localeId: string; sourcePath: string; outputPath: string }>} */
  const out = [];
  let entries;
  try {
    entries = await readdir(localeSourceDir, { withFileTypes: true });
  } catch (err) {
    throw new Error(
      `Failed to read locale sources directory ${path.relative(repoRoot, localeSourceDir)}: ${err}`,
    );
  }

  const localeIds = entries
    .filter((entry) => entry.isFile() && entry.name.endsWith(".json"))
    .map((entry) => entry.name.replace(/\.json$/u, ""))
    .sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));

  if (localeIds.length === 0) {
    throw new Error(
      `No locale source JSON files found under ${path.relative(repoRoot, localeSourceDir)} (expected e.g. de-DE.json)`,
    );
  }

  for (const localeId of localeIds) {
    out.push({
      localeId,
      sourcePath: path.join(localeSourceDir, `${localeId}.json`),
      outputPath: path.join(localeOutputDir, `${localeId}.tsv`),
    });
  }

  return out;
})();

let hadMismatch = false;
for (const locale of locales) {
  const generated = await generateTsvForLocale(locale, canonicalFunctionsSorted, date);
  await mkdir(path.dirname(locale.outputPath), { recursive: true });

  if (check) {
    let existing = null;
    try {
      existing = await readFile(locale.outputPath, "utf8");
    } catch (err) {
      // Treat missing files as mismatches in `--check` mode, with a clear error message.
      if (err && typeof err === "object" && "code" in err && err.code === "ENOENT") {
        existing = null;
      } else {
        throw err;
      }
    }
    const existingNormalized = existing == null ? null : normalizeNewlines(existing);
    const generatedNormalized = normalizeNewlines(generated);
    if (existingNormalized !== generatedNormalized) {
      hadMismatch = true;
      console.error(
        `Locale TSV mismatch: ${path.relative(repoRoot, locale.outputPath)} (run node scripts/generate-locale-function-tsv.js to update)`
      );
    }
  } else {
    await writeFile(locale.outputPath, generated, "utf8");
    console.log(`Wrote ${path.relative(repoRoot, locale.outputPath)}`);
  }
}

if (check) {
  if (hadMismatch) {
    process.exitCode = 1;
  } else {
    console.log("Locale TSVs are up-to-date.");
  }
}
