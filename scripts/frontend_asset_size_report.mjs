#!/usr/bin/env node
/**
 * Frontend asset download size report (+ optional budget gate).
 *
 * Measures Vite-emitted assets under `<dist>/assets/` and approximates the network
 * download size by compressing each asset individually (Brotli or gzip).
 *
 * This is intentionally distinct from:
 * - Desktop *installer artifact* sizes (DMG/MSI/AppImage/etc) → `scripts/desktop_bundle_size_report.py`
 * - Desktop binary/dist on-disk sizes → `scripts/desktop_size_report.py`
 */

import { appendFileSync } from "node:fs";
import { mkdir, readdir, readFile, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { brotliCompressSync, constants as zlibConstants, gzipSync } from "node:zlib";

const DEFAULT_LIMIT_MB = 10;
const DEFAULT_COMPRESSION = "brotli";
const DEFAULT_DIST_DIR = path.join("apps", "web", "dist");
const MB_BYTES = 1_000_000;
const DEFAULT_JSON_OUT_ENV = "FORMULA_FRONTEND_ASSET_SIZE_JSON_PATH";

function isTruthyEnv(val) {
  if (val == null) return false;
  return ["1", "true", "yes", "y", "on"].includes(String(val).trim().toLowerCase());
}

function parseLimitMb(raw) {
  if (raw == null || String(raw).trim() === "") return DEFAULT_LIMIT_MB;
  const n = Number(raw);
  if (!Number.isFinite(n) || n <= 0) {
    throw new Error(`Invalid FORMULA_FRONTEND_ASSET_SIZE_LIMIT_MB=${JSON.stringify(raw)} (expected a number > 0)`);
  }
  return n;
}

function humanBytes(sizeBytes) {
  let size = Number(sizeBytes);
  const units = ["B", "KB", "MB", "GB", "TB"];
  for (const unit of units) {
    if (size < 1000 || unit === units[units.length - 1]) {
      if (unit === "B") return `${Math.trunc(size)} ${unit}`;
      return `${size.toFixed(1)} ${unit}`;
    }
    size /= 1000;
  }
  return `${sizeBytes} B`;
}

function usage() {
  return `Usage:
  node scripts/frontend_asset_size_report.mjs [--dist <path>] [--limit-mb <n>] [--compression brotli|gzip] [--enforce] [--json-out <path>]

Defaults:
  --dist ${DEFAULT_DIST_DIR}
  --limit-mb ${DEFAULT_LIMIT_MB}
  --compression ${DEFAULT_COMPRESSION}

Env overrides:
  FORMULA_FRONTEND_ASSET_SIZE_LIMIT_MB=10
  FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION=brotli|gzip
  FORMULA_ENFORCE_FRONTEND_ASSET_SIZE=1
  ${DEFAULT_JSON_OUT_ENV}=path/to/report.json
`;
}

function parseArgs(argv) {
  let args = argv.slice();
  // pnpm forwards a literal `--` delimiter into scripts. Strip the first occurrence so
  // `pnpm report:... -- --dist apps/desktop/dist` behaves the same as passing args directly.
  const delimiterIdx = args.indexOf("--");
  if (delimiterIdx >= 0) {
    args = [...args.slice(0, delimiterIdx), ...args.slice(delimiterIdx + 1)];
  }

  /** @type {{ dist?: string, limitMb?: string, compression?: string, enforce?: boolean, jsonOut?: string, help?: boolean }} */
  const out = {};
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === "--help" || arg === "-h") {
      out.help = true;
      continue;
    }
    if (arg === "--dist") {
      out.dist = argv[++i];
      continue;
    }
    if (arg.startsWith("--dist=")) {
      out.dist = arg.slice("--dist=".length);
      continue;
    }
    if (arg === "--limit-mb") {
      out.limitMb = argv[++i];
      continue;
    }
    if (arg.startsWith("--limit-mb=")) {
      out.limitMb = arg.slice("--limit-mb=".length);
      continue;
    }
    if (arg === "--compression") {
      out.compression = argv[++i];
      continue;
    }
    if (arg.startsWith("--compression=")) {
      out.compression = arg.slice("--compression=".length);
      continue;
    }
    if (arg === "--enforce") {
      out.enforce = true;
      continue;
    }
    if (arg === "--json-out") {
      out.jsonOut = args[++i];
      continue;
    }
    if (arg.startsWith("--json-out=")) {
      out.jsonOut = arg.slice("--json-out=".length);
      continue;
    }
    throw new Error(`Unknown argument: ${arg}\n\n${usage()}`);
  }
  return out;
}

async function walkFiles(rootDir) {
  /** @type {string[]} */
  const out = [];
  /** @type {string[]} */
  const stack = [rootDir];
  while (stack.length) {
    const dir = stack.pop();
    const entries = await readdir(dir, { withFileTypes: true });
    for (const ent of entries) {
      const full = path.join(dir, ent.name);
      if (ent.isDirectory()) {
        stack.push(full);
        continue;
      }
      if (ent.isFile()) out.push(full);
    }
  }
  return out;
}

function brotliParamsForExt(ext) {
  // Use text mode for JS/CSS and generic for WASM.
  const mode = ext === ".wasm" ? zlibConstants.BROTLI_MODE_GENERIC : zlibConstants.BROTLI_MODE_TEXT;
  return {
    params: {
      [zlibConstants.BROTLI_PARAM_QUALITY]: 11,
      [zlibConstants.BROTLI_PARAM_MODE]: mode,
    },
  };
}

function appendStepSummary(markdown) {
  const summaryPath = process.env.GITHUB_STEP_SUMMARY;
  if (!summaryPath) return;
  // eslint-disable-next-line no-sync
  appendFileSync(summaryPath, `${markdown}\n`, { encoding: "utf8" });
}

function toPosixPath(p) {
  return p.split(path.sep).join("/");
}

function reportPath(p, repoRoot) {
  const abs = path.resolve(p);
  const rel = path.relative(repoRoot, abs);
  if (rel === "") return ".";
  if (!rel.startsWith("..") && !path.isAbsolute(rel)) return toPosixPath(rel);
  return toPosixPath(abs);
}

async function writeJsonReport(jsonOutPath, report) {
  if (!jsonOutPath) return true;
  try {
    await mkdir(path.dirname(jsonOutPath), { recursive: true });
    await writeFile(jsonOutPath, `${JSON.stringify(report, null, 2)}\n`, "utf8");
    return true;
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error(`frontend-asset-size: ERROR failed to write JSON report to ${jsonOutPath}: ${msg}`);
    return false;
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    console.log(usage());
    return 0;
  }

  const repoRoot = process.cwd();
  const distDir = path.resolve(repoRoot, args.dist ?? DEFAULT_DIST_DIR);
  const assetsDir = path.join(distDir, "assets");

  const rawJsonOut = args.jsonOut ?? process.env[DEFAULT_JSON_OUT_ENV];
  const jsonOutPath = rawJsonOut == null ? "" : String(rawJsonOut).trim();
  const jsonOut = jsonOutPath !== "" ? path.resolve(repoRoot, jsonOutPath) : null;

  const enforce = Boolean(args.enforce) || isTruthyEnv(process.env.FORMULA_ENFORCE_FRONTEND_ASSET_SIZE);
  const limitMb = args.limitMb != null ? Number(args.limitMb) : parseLimitMb(process.env.FORMULA_FRONTEND_ASSET_SIZE_LIMIT_MB);
  if (!Number.isFinite(limitMb) || limitMb <= 0) {
    console.error(`frontend-asset-size: ERROR invalid --limit-mb=${JSON.stringify(args.limitMb)} (expected a number > 0)`);
    return 2;
  }
  const limitBytes = limitMb * MB_BYTES;

  const rawCompression =
    args.compression ?? process.env.FORMULA_FRONTEND_ASSET_SIZE_COMPRESSION ?? DEFAULT_COMPRESSION;
  const compression = (rawCompression && String(rawCompression).trim() ? rawCompression : DEFAULT_COMPRESSION)
    .trim()
    .toLowerCase();
  if (compression !== "brotli" && compression !== "gzip") {
    console.error(`frontend-asset-size: ERROR invalid compression mode: ${compression} (expected brotli|gzip)`);
    return 2;
  }

  const runnerOs = process.env.RUNNER_OS?.trim();

  const assetsStat = await stat(assetsDir).catch(() => null);
  if (!assetsStat?.isDirectory()) {
    const relAssetsDir = path.relative(repoRoot, assetsDir).split(path.sep).join("/");
    const msg = [
      `## Frontend asset download size (${relAssetsDir})`,
      "",
      `- Budget: **${limitMb} MB** total (${compression})`,
      `- Enforcement: **${enforce ? "enabled" : "disabled"}**`,
      "",
      "_Vite `dist/assets/` directory not found._",
      "",
      "Hint: build first (e.g. `pnpm build:web` or `pnpm build:desktop`).",
      "",
    ].join("\n");
    console.error(`frontend-asset-size: ERROR missing Vite assets directory: ${relAssetsDir}`);
    console.log(msg);
    try {
      appendStepSummary(msg);
    } catch {
      // Ignore summary write failures.
    }

    const report = {
      dist_dir: reportPath(distDir, repoRoot),
      assets_dir: reportPath(assetsDir, repoRoot),
      compression,
      limit_mb: limitMb,
      limit_bytes: limitBytes,
      enforce,
      file_count: 0,
      totals: {
        raw_bytes: 0,
        brotli_bytes: 0,
        gzip_bytes: 0,
        compressed_bytes: 0,
      },
      over_limit: false,
      missing_assets_dir: true,
      assets: [],
    };
    if (runnerOs) report.runner_os = runnerOs;
    const wroteJson = await writeJsonReport(jsonOut, report);
    if (!wroteJson) return 2;

    return 1;
  }

  const allFiles = await walkFiles(assetsDir);
  const assetFiles = allFiles.filter((f) => {
    if (f.endsWith(".map")) return false;
    const ext = path.extname(f);
    return ext === ".js" || ext === ".css" || ext === ".wasm";
  });

  /** @type {Array<{ rel: string, ext: string, rawBytes: number, gzipBytes: number, brotliBytes: number }>} */
  const rows = [];
  for (const absPath of assetFiles) {
    const info = await stat(absPath);
    const rawBytes = info.size;
    const buf = await readFile(absPath);
    const ext = path.extname(absPath);
    const gzipBytes = gzipSync(buf, { level: 9 }).length;
    const brotliBytes = brotliCompressSync(buf, brotliParamsForExt(ext)).length;
    rows.push({
      rel: path.relative(repoRoot, absPath).split(path.sep).join("/"),
      ext,
      rawBytes,
      gzipBytes,
      brotliBytes,
    });
  }

  const totalRaw = rows.reduce((sum, r) => sum + r.rawBytes, 0);
  const totalGzip = rows.reduce((sum, r) => sum + r.gzipBytes, 0);
  const totalBrotli = rows.reduce((sum, r) => sum + r.brotliBytes, 0);
  const totalCompressed = compression === "brotli" ? totalBrotli : totalGzip;
  const overLimit = totalCompressed > limitBytes;

  const reportAssets = [...rows]
    .sort((a, b) => a.rel.localeCompare(b.rel))
    .map((r) => ({
      path: r.rel,
      ext: r.ext,
      raw_bytes: r.rawBytes,
      brotli_bytes: r.brotliBytes,
      gzip_bytes: r.gzipBytes,
    }));

  rows.sort((a, b) => {
    const ak = compression === "brotli" ? a.brotliBytes : a.gzipBytes;
    const bk = compression === "brotli" ? b.brotliBytes : b.gzipBytes;
    return bk - ak;
  });

  const relAssetsDir = path.relative(repoRoot, assetsDir).split(path.sep).join("/");
  const lines = [];
  lines.push(`## Frontend asset download size (${relAssetsDir})`);
  lines.push("");
  lines.push(`- Budget: **${limitMb} MB** total (${compression})`);
  lines.push(
    `- Enforcement: **${enforce ? "enabled" : "disabled"}**${
      enforce ? "" : " (set `FORMULA_ENFORCE_FRONTEND_ASSET_SIZE=1` to fail on oversize)"
    }`,
  );
  lines.push("");
  lines.push(
    `Totals: raw **${humanBytes(totalRaw)}**, brotli **${humanBytes(totalBrotli)}**, gzip **${humanBytes(totalGzip)}**`,
  );
  lines.push("");

  if (rows.length === 0) {
    lines.push("_No matching assets found (expected .js/.css/.wasm under dist/assets)._");
    lines.push("");
  } else {
    lines.push("| Asset | Raw | Brotli | Gzip |");
    lines.push("| --- | ---: | ---: | ---: |");
    for (const r of rows.slice(0, 50)) {
      lines.push(
        `| \`${r.rel}\` | ${humanBytes(r.rawBytes)} | ${humanBytes(r.brotliBytes)} | ${humanBytes(r.gzipBytes)} |`,
      );
    }
    if (rows.length > 50) {
      lines.push("");
      lines.push(`_(${rows.length - 50} more assets omitted)_`);
      lines.push("");
    } else {
      lines.push("");
    }
  }

  if (overLimit) {
    lines.push(`Total ${compression} size is over the **${limitMb} MB** budget: **${humanBytes(totalCompressed)}**`);
    lines.push("");
  }

  const markdown = lines.join("\n");
  console.log(markdown);
  process.stdout.write("\n");

  try {
    appendStepSummary(markdown);
  } catch {
    // Ignore summary write failures.
  }

  const jsonReport = {
    dist_dir: reportPath(distDir, repoRoot),
    assets_dir: reportPath(assetsDir, repoRoot),
    compression,
    limit_mb: limitMb,
    limit_bytes: limitBytes,
    enforce,
    file_count: rows.length,
    totals: {
      raw_bytes: totalRaw,
      brotli_bytes: totalBrotli,
      gzip_bytes: totalGzip,
      compressed_bytes: totalCompressed,
    },
    over_limit: overLimit,
    missing_assets_dir: false,
    assets: reportAssets,
  };
  if (runnerOs) jsonReport.runner_os = runnerOs;
  const wroteJson = await writeJsonReport(jsonOut, jsonReport);
  if (!wroteJson) return 2;

  if (!enforce) return 0;
  if (overLimit) {
    console.error(
      `frontend-asset-size: ERROR total ${compression} size ${humanBytes(totalCompressed)} exceeds ${limitMb} MB ` +
        `(set FORMULA_FRONTEND_ASSET_SIZE_LIMIT_MB to adjust)`,
    );
    return 1;
  }
  return 0;
}

let exitCode = 0;
try {
  exitCode = await main();
} catch (err) {
  console.error("frontend-asset-size: ERROR failed to generate report.");
  console.error(err);
  exitCode = 2;
}
process.exitCode = exitCode;
