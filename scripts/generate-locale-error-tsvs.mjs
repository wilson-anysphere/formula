#!/usr/bin/env node
/**
 * Generate locale error-literal translation TSVs from committed upstream sources.
 *
 * This script is intentionally dependency-free (Node built-ins only) so it can be
 * run anywhere we build/test the repo.
 *
 * Usage:
 *   node scripts/generate-locale-error-tsvs.mjs        # (re)generate TSVs in-place
 *   node scripts/generate-locale-error-tsvs.mjs --check # verify committed TSVs are up-to-date
 *
 * See `crates/formula-engine/src/locale/data/README.md` for details.
 */
import { readFile, writeFile, mkdir, readdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { inspect } from "node:util";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const rustErrorKindPath = path.join(
  repoRoot,
  "crates",
  "formula-engine",
  "src",
  "value",
  "mod.rs",
);

const upstreamDir = path.join(
  repoRoot,
  "crates",
  "formula-engine",
  "src",
  "locale",
  "data",
  "upstream",
  "errors",
);

const outDir = path.join(repoRoot, "crates", "formula-engine", "src", "locale", "data");

function parseArgs(argv) {
  /** @type {{ check: boolean }} */
  const out = { check: false };
  for (const arg of argv) {
    if (arg === "--check") {
      out.check = true;
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }
  return out;
}

/**
 * @param {string} rustSource
 * @returns {string[]} canonical error literals (e.g. "#VALUE!")
 */
function extractCanonicalErrorLiterals(rustSource) {
  // We intentionally scrape the source instead of duplicating the list to ensure
  // this generator stays in sync with `ErrorKind::as_code`.
  //
  // Match lines like: `ErrorKind::Value => "#VALUE!",`
  const re = /ErrorKind::[A-Za-z0-9_]+\s*=>\s*"([^"]+)"/g;
  /** @type {string[]} */
  const codes = [];
  for (const match of rustSource.matchAll(re)) {
    codes.push(match[1]);
  }
  const unique = Array.from(new Set(codes));
  if (unique.length === 0) {
    throw new Error(
      `Failed to extract any error literals from ${path.relative(repoRoot, rustErrorKindPath)}; regex=${re}`,
    );
  }
  if (unique.length !== codes.length) {
    throw new Error(
      `Duplicate error literals found in ${path.relative(repoRoot, rustErrorKindPath)}: ${inspect(codes)}`,
    );
  }
  // Preserve source order (which matches the engine's canonical ErrorKind order).
  return codes;
}

/**
 * Error TSV convention:
 * - `# ` (hash + space) starts a comment line (error literals themselves start with `#`)
 * - empty lines ignored
 * - each entry: `Canonical<TAB>Localized`
 *
 * @param {string} contents
 * @param {string} label
 * @returns {Map<string, string>}
 */
function parseErrorTsv(contents, label) {
  /** @type {Map<string, string>} */
  const map = new Map();
  const lines = contents.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const raw = lines[i];
    const trimmed = raw.trim();
    // Error TSVs allow data lines that start with `#` (e.g. `#VALUE!`), so we treat
    // comments as `#` followed by whitespace (or `#` alone).
    const isComment =
      trimmed === "#" ||
      (trimmed.startsWith("#") && (trimmed[1] === " " || trimmed[1] === "\t"));
    if (trimmed.length === 0 || isComment) {
      continue;
    }
    const parts = raw.split("\t");
    if (parts.length < 2) {
      throw new Error(`Invalid TSV line in ${label}:${i + 1} (expected 2 columns): ${inspect(raw)}`);
    }
    const canonical = parts[0].trim();
    const localized = parts[1].trim();
    if (!canonical || !localized) {
      throw new Error(`Invalid TSV line in ${label}:${i + 1} (empty field): ${inspect(raw)}`);
    }
    if (map.has(canonical)) {
      throw new Error(
        `Duplicate canonical entry in ${label}:${i + 1}: ${inspect(canonical)} (previously defined)`,
      );
    }
    map.set(canonical, localized);
  }
  return map;
}

/**
 * @param {object} params
 * @param {string} params.locale
 * @param {string[]} params.canonicalLiteralsSorted
 * @param {Map<string, string>} params.upstreamMap
 * @param {string} params.upstreamRelPath
 */
function renderOutputTsv({ locale, canonicalLiteralsSorted, upstreamMap, upstreamRelPath }) {
  const lines = [];
  lines.push("# Canonical\tLocalized");
  lines.push("# See `src/locale/data/README.md` for format + generators.");
  lines.push("");

  for (const canonical of canonicalLiteralsSorted) {
    const localized = upstreamMap.get(canonical);
    if (localized == null) {
      throw new Error(
        `Upstream mapping ${upstreamRelPath} is missing an entry for ${inspect(canonical)} (locale ${locale})`,
      );
    }
    lines.push(`${canonical}\t${localized}`);
  }

  return lines.join("\n") + "\n";
}

/**
 * @param {string} filePath
 * @returns {Promise<string|null>}
 */
async function readUtf8IfExists(filePath) {
  if (!existsSync(filePath)) {
    return null;
  }
  return await readFile(filePath, "utf8");
}

async function main() {
  const { check } = parseArgs(process.argv.slice(2));

  const rustSource = await readFile(rustErrorKindPath, "utf8");
  const canonicalLiteralsSorted = extractCanonicalErrorLiterals(rustSource);

  if (!existsSync(upstreamDir)) {
    throw new Error(
      `Upstream directory not found: ${path.relative(repoRoot, upstreamDir)} (expected committed mapping sources)`,
    );
  }

  const entries = (await readdir(upstreamDir, { withFileTypes: true }))
    .filter((e) => e.isFile() && e.name.endsWith(".tsv"))
    .map((e) => e.name)
    .sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));

  if (entries.length === 0) {
    throw new Error(
      `No upstream TSV sources found under ${path.relative(repoRoot, upstreamDir)} (expected *.tsv files named like de-DE.tsv)`,
    );
  }

  await mkdir(outDir, { recursive: true });

  /** @type {string[]} */
  const mismatches = [];
  /** @type {string[]} */
  const updated = [];

  for (const fileName of entries) {
    const locale = fileName.replace(/\.tsv$/u, "");
    const upstreamPath = path.join(upstreamDir, fileName);
    const upstreamRelPath = path.relative(repoRoot, upstreamPath).replaceAll(path.sep, "/");
    const upstreamContents = await readFile(upstreamPath, "utf8");
    const upstreamMap = parseErrorTsv(upstreamContents, upstreamRelPath);

    // Validate upstream entries are in our canonical set.
    for (const key of upstreamMap.keys()) {
      if (!canonicalLiteralsSorted.includes(key)) {
        throw new Error(
          `Upstream mapping ${upstreamRelPath} contains unknown canonical error literal ${inspect(
            key,
          )}. Canonical set is derived from ErrorKind::as_code.`,
        );
      }
    }

    const output = renderOutputTsv({
      locale,
      canonicalLiteralsSorted,
      upstreamMap,
      upstreamRelPath,
    });

    const outPath = path.join(outDir, `${locale}.errors.tsv`);
    const existing = await readUtf8IfExists(outPath);

    if (check) {
      if (existing == null) {
        mismatches.push(path.relative(repoRoot, outPath));
      } else if (existing !== output) {
        mismatches.push(path.relative(repoRoot, outPath));
      }
    } else {
      if (existing !== output) {
        await writeFile(outPath, output, "utf8");
        updated.push(path.relative(repoRoot, outPath));
      }
    }
  }

  if (check) {
    if (mismatches.length > 0) {
      console.error("Locale error TSVs are out of date:");
      for (const file of mismatches) {
        console.error(`  - ${file}`);
      }
      console.error(
        "\nRegenerate them with:\n  node scripts/generate-locale-error-tsvs.mjs\n",
      );
      process.exit(1);
    }
    console.log("Locale error TSVs are up to date.");
  } else {
    if (updated.length === 0) {
      console.log("Locale error TSVs already up to date.");
    } else {
      console.log("Updated locale error TSVs:");
      for (const file of updated) {
        console.log(`  - ${file}`);
      }
    }
  }
}

await main();
