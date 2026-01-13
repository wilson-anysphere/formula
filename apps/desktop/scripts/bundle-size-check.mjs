import { appendFile, readFile, readdir, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { gzipSync } from "node:zlib";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.resolve(__dirname, "..");
const defaultDistDir = path.join(desktopRoot, "dist");

const KB = 1024;
const numberFmt = new Intl.NumberFormat("en-US");

function fmtInt(n) {
  return numberFmt.format(n);
}

function isTruthyEnv(val) {
  if (val == null) return false;
  return ["1", "true", "yes", "y", "on"].includes(String(val).trim().toLowerCase());
}

function fmtKiB(bytes, decimals = 1) {
  return (bytes / KB).toFixed(decimals);
}

function toPosixPath(p) {
  return p.split(path.sep).join("/");
}

function usage() {
  console.log(
    [
      "Check Vite desktop bundle sizes (apps/desktop/dist).",
      "",
      "Usage:",
      "  node apps/desktop/scripts/bundle-size-check.mjs [--dist <path>]",
      "",
      "Environment (budgets are interpreted as KiB = 1024 bytes):",
      "  FORMULA_DESKTOP_JS_TOTAL_BUDGET_KB   Total Vite JS (dist/assets/**/*.js) budget.",
      "  FORMULA_DESKTOP_JS_ENTRY_BUDGET_KB   Entry JS budget (scripts referenced by dist/index.html).",
      "  FORMULA_DESKTOP_JS_DIST_TOTAL_BUDGET_KB   Optional: total JS across dist/**/*.js (includes copied public assets).",
      "",
      "Optional:",
      "  FORMULA_DESKTOP_BUNDLE_SIZE_WARN_ONLY=1   Print budget errors but exit 0.",
      "  FORMULA_DESKTOP_BUNDLE_SIZE_SKIP_GZIP=1   Skip gzip computation (faster).",
    ].join("\n"),
  );
}

/**
 * @param {string[]} args
 */
function parseArgs(args) {
  /** @type {{ distDir?: string }} */
  const out = {};
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (!arg) continue;
    if (arg === "--help" || arg === "-h") {
      out.help = true;
      continue;
    }
    if (arg === "--dist") {
      out.distDir = args[i + 1];
      i++;
      continue;
    }
    if (arg.startsWith("--dist=")) {
      out.distDir = arg.slice("--dist=".length);
      continue;
    }
    throw new Error(`Unknown argument: ${arg}`);
  }
  return out;
}

async function pathExists(p) {
  try {
    await stat(p);
    return true;
  } catch {
    return false;
  }
}

/**
 * @param {string} root
 * @param {(absPath: string) => boolean} predicate
 * @returns {Promise<string[]>}
 */
async function collectFiles(root, predicate) {
  /** @type {string[]} */
  const out = [];
  const entries = await readdir(root, { withFileTypes: true });
  for (const entry of entries) {
    const absPath = path.join(root, entry.name);
    if (entry.isDirectory()) {
      out.push(...(await collectFiles(absPath, predicate)));
      continue;
    }
    if (entry.isFile() && predicate(absPath)) {
      out.push(absPath);
    }
  }
  return out;
}

function normalizeHtmlAssetPath(rawSrc) {
  const src = rawSrc.trim();
  if (!src) return null;
  if (src.startsWith("data:")) return null;
  if (src.startsWith("http://") || src.startsWith("https://") || src.startsWith("//")) return null;

  // Strip search/hash.
  const cleaned = src.split("?", 1)[0]?.split("#", 1)[0] ?? "";
  if (!cleaned) return null;

  // Vite typically emits `/assets/...` but support relative forms too.
  let rel = cleaned;
  if (rel.startsWith("/")) rel = rel.slice(1);
  if (rel.startsWith("./")) rel = rel.slice(2);

  return rel || null;
}

function extractHtmlAttr(tagHtml, attrName) {
  // Basic attribute parsing; good enough for Vite-generated HTML.
  const re = new RegExp(`\\b${attrName}\\s*=\\s*(?:\"([^\"]+)\"|'([^']+)'|([^\\s>]+))`, "i");
  const match = tagHtml.match(re);
  return match?.[1] ?? match?.[2] ?? match?.[3] ?? null;
}

function extractScriptSrcs(html) {
  /** @type {string[]} */
  const out = [];
  const scriptRe = /<script\b[^>]*>/gi;
  for (const match of html.matchAll(scriptRe)) {
    const tag = match[0];
    const src = extractHtmlAttr(tag, "src");
    if (!src) continue;
    const normalized = normalizeHtmlAssetPath(src);
    if (!normalized) continue;
    if (!normalized.endsWith(".js")) continue;
    out.push(normalized);
  }
  return out;
}

function extractModulePreloadHrefs(html) {
  /** @type {string[]} */
  const out = [];
  const linkRe = /<link\b[^>]*>/gi;
  for (const match of html.matchAll(linkRe)) {
    const tag = match[0];
    const rel = extractHtmlAttr(tag, "rel");
    if (!rel) continue;
    if (rel.toLowerCase() !== "modulepreload") continue;
    const href = extractHtmlAttr(tag, "href");
    if (!href) continue;
    const normalized = normalizeHtmlAssetPath(href);
    if (!normalized) continue;
    if (!normalized.endsWith(".js")) continue;
    out.push(normalized);
  }
  return out;
}

function parseBudgetKiB(envVarName) {
  const raw = process.env[envVarName];
  if (!raw) return null;
  const n = Number(raw);
  if (!Number.isFinite(n) || n < 0) {
    throw new Error(`Invalid ${envVarName}="${raw}" (expected a non-negative number of KiB).`);
  }
  return Math.round(n);
}

function budgetStatus(actualBytes, budgetKiB) {
  if (budgetKiB == null) return { ok: true, text: "—" };
  const budgetBytes = budgetKiB * KB;
  if (actualBytes <= budgetBytes) return { ok: true, text: "OK" };
  const deltaKiB = (actualBytes - budgetBytes) / KB;
  return { ok: false, text: `OVER (+${deltaKiB.toFixed(1)} KiB)` };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    usage();
    process.exit(0);
  }

  const distDir = args.distDir ? path.resolve(args.distDir) : defaultDistDir;
  const indexHtmlPath = path.join(distDir, "index.html");
  const assetsDir = path.join(distDir, "assets");

  if (!(await pathExists(distDir))) {
    throw new Error(`Missing dist directory: ${distDir}\nRun: pnpm -C apps/desktop build`);
  }
  if (!(await pathExists(indexHtmlPath))) {
    throw new Error(`Missing index.html: ${indexHtmlPath}\nRun: pnpm -C apps/desktop build`);
  }

  const skipGzip = isTruthyEnv(process.env.FORMULA_DESKTOP_BUNDLE_SIZE_SKIP_GZIP);

  const jsAbsPaths = await collectFiles(distDir, (p) => p.endsWith(".js") && !p.endsWith(".js.map"));

  /** @type {{ absPath: string; relPath: string; bytes: number; gzipBytes: number | null }[]} */
  const jsFiles = [];
  for (const absPath of jsAbsPaths) {
    const info = await stat(absPath);
    const bytes = info.size;
    let gzipBytes = null;
    if (!skipGzip) {
      const buf = await readFile(absPath);
      gzipBytes = gzipSync(buf, { level: 9 }).length;
    }
    jsFiles.push({
      absPath,
      relPath: toPosixPath(path.relative(distDir, absPath)),
      bytes,
      gzipBytes,
    });
  }

  const totalBytes = jsFiles.reduce((sum, f) => sum + f.bytes, 0);
  const totalGzipBytes = skipGzip ? null : jsFiles.reduce((sum, f) => sum + (f.gzipBytes ?? 0), 0);
  const assetFiles = jsFiles.filter((f) => f.relPath.startsWith("assets/"));
  const assetsBytes = assetFiles.reduce((sum, f) => sum + f.bytes, 0);
  const assetsGzipBytes = skipGzip ? null : assetFiles.reduce((sum, f) => sum + (f.gzipBytes ?? 0), 0);

  const indexHtml = await readFile(indexHtmlPath, "utf8");
  const entryRelPaths = extractScriptSrcs(indexHtml);
  const preloadRelPaths = extractModulePreloadHrefs(indexHtml);

  /** @type {Map<string, { relPath: string; bytes: number; gzipBytes: number | null }>} */
  const byRelPath = new Map();
  for (const f of jsFiles) byRelPath.set(f.relPath, f);

  const missingEntry = [];
  /** @type {{ relPath: string; bytes: number; gzipBytes: number | null }[]} */
  const entryFiles = [];
  for (const rel of entryRelPaths) {
    const normalized = toPosixPath(rel);
    const file = byRelPath.get(normalized);
    if (!file) {
      missingEntry.push(rel);
      continue;
    }
    entryFiles.push(file);
  }

  const missingPreloads = [];
  /** @type {{ relPath: string; bytes: number; gzipBytes: number | null }[]} */
  const preloadFiles = [];
  for (const rel of preloadRelPaths) {
    const normalized = toPosixPath(rel);
    const file = byRelPath.get(normalized);
    if (!file) {
      missingPreloads.push(rel);
      continue;
    }
    preloadFiles.push(file);
  }

  if (missingEntry.length > 0) {
    throw new Error(
      [
        `index.html references JS scripts that were not found in dist:`,
        ...missingEntry.map((p) => `- ${p}`),
      ].join("\n"),
    );
  }

  const entryBytes = entryFiles.reduce((sum, f) => sum + f.bytes, 0);
  const entryGzipBytes = skipGzip ? null : entryFiles.reduce((sum, f) => sum + (f.gzipBytes ?? 0), 0);

  const initialFiles = new Map();
  for (const f of entryFiles) initialFiles.set(f.relPath, f);
  for (const f of preloadFiles) initialFiles.set(f.relPath, f);
  const initialBytes = Array.from(initialFiles.values()).reduce((sum, f) => sum + f.bytes, 0);
  const initialGzipBytes = skipGzip
    ? null
    : Array.from(initialFiles.values()).reduce((sum, f) => sum + (f.gzipBytes ?? 0), 0);

  const viteTotalBudgetKiB = parseBudgetKiB("FORMULA_DESKTOP_JS_TOTAL_BUDGET_KB");
  const entryBudgetKiB = parseBudgetKiB("FORMULA_DESKTOP_JS_ENTRY_BUDGET_KB");
  const distTotalBudgetKiB = parseBudgetKiB("FORMULA_DESKTOP_JS_DIST_TOTAL_BUDGET_KB");
  const warnOnly = isTruthyEnv(process.env.FORMULA_DESKTOP_BUNDLE_SIZE_WARN_ONLY);

  const viteTotalStatus = budgetStatus(assetsBytes, viteTotalBudgetKiB);
  const entryStatus = budgetStatus(entryBytes, entryBudgetKiB);
  const distTotalStatus = budgetStatus(totalBytes, distTotalBudgetKiB);

  const lines = [];
  lines.push("### Desktop bundle size (Vite)");
  lines.push("");
  lines.push(`Measured JS bundles in \`${toPosixPath(path.relative(process.cwd(), distDir))}\`.`);
  lines.push("");
  lines.push("| Metric | Files | Bytes | KiB | Gzip KiB | Budget (KiB) | Status |");
  lines.push("| --- | ---: | ---: | ---: | ---: | ---: | --- |");
  lines.push(
    `| Total JS (dist/assets/**/*.js) | ${assetFiles.length} | ${fmtInt(assetsBytes)} | ${fmtKiB(assetsBytes)} | ${
      assetsGzipBytes == null ? "—" : fmtKiB(assetsGzipBytes)
    } | ${viteTotalBudgetKiB == null ? "—" : viteTotalBudgetKiB} | ${viteTotalStatus.text} |`,
  );
  lines.push(
    `| Total JS (dist/**/*.js) | ${jsFiles.length} | ${fmtInt(totalBytes)} | ${fmtKiB(totalBytes)} | ${
      totalGzipBytes == null ? "—" : fmtKiB(totalGzipBytes)
    } | ${distTotalBudgetKiB == null ? "—" : distTotalBudgetKiB} | ${distTotalStatus.text} |`,
  );
  lines.push(
    `| Entry JS (script tags) | ${entryFiles.length} | ${fmtInt(entryBytes)} | ${fmtKiB(entryBytes)} | ${
      entryGzipBytes == null ? "—" : fmtKiB(entryGzipBytes)
    } | ${entryBudgetKiB == null ? "—" : entryBudgetKiB} | ${entryStatus.text} |`,
  );
  lines.push(
    `| Initial JS (scripts + modulepreload) | ${initialFiles.size} | ${fmtInt(initialBytes)} | ${fmtKiB(
      initialBytes,
    )} | ${initialGzipBytes == null ? "—" : fmtKiB(initialGzipBytes)} | — | — |`,
  );

  // Always print entry and top bundles so CI logs are actionable even without budgets.
  if (entryFiles.length > 0) {
    lines.push("");
    lines.push("#### Entry scripts (`dist/index.html`)");
    lines.push("");
    lines.push("| File | Bytes | KiB | Gzip KiB |");
    lines.push("| --- | ---: | ---: | ---: |");
    for (const f of entryFiles) {
      lines.push(
        `| \`${f.relPath}\` | ${fmtInt(f.bytes)} | ${fmtKiB(f.bytes)} | ${f.gzipBytes == null ? "—" : fmtKiB(f.gzipBytes)} |`,
      );
    }
  }

  if (preloadFiles.length > 0) {
    lines.push("");
    lines.push("#### Modulepreload JS (`dist/index.html`)");
    lines.push("");
    lines.push("| File | Bytes | KiB | Gzip KiB |");
    lines.push("| --- | ---: | ---: | ---: |");
    for (const f of preloadFiles) {
      lines.push(
        `| \`${f.relPath}\` | ${fmtInt(f.bytes)} | ${fmtKiB(f.bytes)} | ${f.gzipBytes == null ? "—" : fmtKiB(f.gzipBytes)} |`,
      );
    }
  } else if (missingPreloads.length > 0) {
    // Not fatal; modulepreloads are optional and can vary.
    lines.push("");
    lines.push("#### Modulepreload JS (`dist/index.html`)");
    lines.push("");
    lines.push("Found modulepreload tags that did not resolve to `dist/assets` JS files:");
    lines.push("");
    for (const p of missingPreloads) lines.push(`- \`${p}\``);
  }

  const largestAssets = [...assetFiles].sort((a, b) => b.bytes - a.bytes).slice(0, 10);
  const largest = [...jsFiles].sort((a, b) => b.bytes - a.bytes).slice(0, 10);

  lines.push("");
  lines.push("#### Largest Vite JS bundles (`dist/assets`)");
  lines.push("");
  lines.push("| File | Bytes | KiB | Gzip KiB |");
  lines.push("| --- | ---: | ---: | ---: |");
  for (const f of largestAssets) {
    const isEntry = entryFiles.some((e) => e.relPath === f.relPath);
    lines.push(
      `| \`${f.relPath}\`${isEntry ? " (entry)" : ""} | ${fmtInt(f.bytes)} | ${fmtKiB(f.bytes)} | ${
        f.gzipBytes == null ? "—" : fmtKiB(f.gzipBytes)
      } |`,
    );
  }

  lines.push("");
  lines.push("#### Largest JS files (`dist`)"); // Includes Vite bundles + public assets copied into dist.
  lines.push("");
  lines.push("| File | Bytes | KiB | Gzip KiB |");
  lines.push("| --- | ---: | ---: | ---: |");
  for (const f of largest) {
    const isEntry = entryFiles.some((e) => e.relPath === f.relPath);
    lines.push(
      `| \`${f.relPath}\`${isEntry ? " (entry)" : ""} | ${fmtInt(f.bytes)} | ${fmtKiB(f.bytes)} | ${
        f.gzipBytes == null ? "—" : fmtKiB(f.gzipBytes)
      } |`,
    );
  }

  const markdown = `${lines.join("\n")}\n`;
  console.log(markdown);

  const summaryPath = process.env.GITHUB_STEP_SUMMARY;
  if (summaryPath) {
    await appendFile(summaryPath, markdown);
  }

  const violations = [];
  if (viteTotalBudgetKiB != null && !viteTotalStatus.ok) {
    violations.push(
      `Total Vite JS: ${fmtKiB(assetsBytes)} KiB > budget ${viteTotalBudgetKiB} KiB (FORMULA_DESKTOP_JS_TOTAL_BUDGET_KB)`,
    );
  }
  if (entryBudgetKiB != null && !entryStatus.ok) {
    violations.push(
      `Entry JS: ${fmtKiB(entryBytes)} KiB > budget ${entryBudgetKiB} KiB (FORMULA_DESKTOP_JS_ENTRY_BUDGET_KB)`,
    );
  }
  if (distTotalBudgetKiB != null && !distTotalStatus.ok) {
    violations.push(
      `Total dist JS: ${fmtKiB(totalBytes)} KiB > budget ${distTotalBudgetKiB} KiB (FORMULA_DESKTOP_JS_DIST_TOTAL_BUDGET_KB)`,
    );
  }

  if (violations.length > 0) {
    const errorLines = [];
    errorLines.push("Bundle size budgets exceeded:");
    for (const v of violations) errorLines.push(`- ${v}`);
    errorLines.push("");
    errorLines.push("Offending files (entry scripts):");
    for (const f of entryFiles) {
      errorLines.push(`- ${f.relPath} (${fmtKiB(f.bytes)} KiB)`);
    }
    errorLines.push("");
    if (!viteTotalStatus.ok) {
      errorLines.push("Largest Vite JS bundles (dist/assets):");
      for (const f of largestAssets.slice(0, 5)) {
        errorLines.push(`- ${f.relPath} (${fmtKiB(f.bytes)} KiB)`);
      }
    } else {
      errorLines.push("Largest JS files (dist):");
      for (const f of largest.slice(0, 5)) {
        errorLines.push(`- ${f.relPath} (${fmtKiB(f.bytes)} KiB)`);
      }
    }

    if (warnOnly) {
      console.warn(errorLines.join("\n"));
      return;
    }

    console.error(errorLines.join("\n"));
    process.exitCode = 1;
  }
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
