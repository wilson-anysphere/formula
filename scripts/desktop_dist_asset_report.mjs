import { appendFile, readdir, stat } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

const BYTES_PER_MB = 1_000_000;
const DEFAULT_TOP_N = 25;
const DEFAULT_GROUP_DEPTH = 1;

/**
 * @param {string} p
 */
function toPosixPath(p) {
  return p.split(path.sep).join("/");
}

/**
 * @param {number} n
 */
function formatInt(n) {
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(n);
}

/**
 * @param {number} bytes
 */
function formatMb(bytes) {
  return `${(bytes / BYTES_PER_MB).toFixed(2)} MB`;
}

/**
 * @param {number} bytes
 */
function formatBytesAndMb(bytes) {
  return `${formatInt(bytes)} (${formatMb(bytes)})`;
}

/**
 * @param {string} raw
 * @param {string} envName
 * @returns {number}
 */
function parseBudgetMb(raw, envName) {
  const cleaned = raw.trim();
  if (!cleaned) throw new Error(`Invalid ${envName}: empty string`);
  const value = Number.parseFloat(cleaned);
  const quoted = JSON.stringify(raw);
  if (!Number.isFinite(value)) throw new Error(`Invalid ${envName}=${quoted} (expected a number)`);
  if (value <= 0) throw new Error(`Invalid ${envName}=${quoted} (must be > 0)`);
  return value;
}

function usage() {
  return [
    "Desktop dist asset size report (top offenders + optional budgets).",
    "",
    "Usage:",
    "  node scripts/desktop_dist_asset_report.mjs [--dist-dir <path>] [--top N] [--group-depth N] [--no-groups]",
    "",
    "Options:",
    "  --dist-dir <path>   Directory to scan (default: apps/desktop/dist).",
    "  --top N             Number of largest files to show (default: 25).",
    "  --group-depth N     Group totals by the first N path segments (default: 1).",
    "  --no-groups         Disable grouped totals output.",
    "",
    "Budgets (env vars):",
    "  FORMULA_DESKTOP_DIST_TOTAL_BUDGET_MB        Fail if total dist size exceeds this value.",
    "  FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB  Fail if any single file exceeds this value.",
    "",
  ].join("\n");
}

/**
 * @param {string[]} argv
 */
function parseArgs(argv) {
  /** @type {{ distDir: string, topN: number, groupDepth: number, groups: boolean }} */
  const out = {
    distDir: path.join(repoRoot, "apps", "desktop", "dist"),
    topN: DEFAULT_TOP_N,
    groupDepth: DEFAULT_GROUP_DEPTH,
    groups: true,
  };

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];

    if (arg === "--help" || arg === "-h") {
      console.log(usage());
      process.exit(0);
    }

    if (arg === "--no-groups" || arg === "--no-group") {
      out.groups = false;
      continue;
    }

    const distDirMatch = arg.match(/^--dist-dir=(.*)$/);
    if (distDirMatch) {
      out.distDir = distDirMatch[1];
      continue;
    }
    if (arg === "--dist-dir") {
      const next = argv[i + 1];
      if (!next) throw new Error("Missing value for --dist-dir");
      out.distDir = next;
      i++;
      continue;
    }

    const topMatch = arg.match(/^--top=(.*)$/) || arg.match(/^--limit=(.*)$/);
    if (topMatch) {
      const value = Number.parseInt(topMatch[1], 10);
      if (!Number.isFinite(value) || value <= 0) throw new Error(`Invalid --top value: ${topMatch[1]}`);
      out.topN = value;
      continue;
    }
    if (arg === "--top" || arg === "--limit") {
      const next = argv[i + 1];
      if (!next) throw new Error(`Missing value for ${arg}`);
      const value = Number.parseInt(next, 10);
      if (!Number.isFinite(value) || value <= 0) throw new Error(`Invalid ${arg} value: ${next}`);
      out.topN = value;
      i++;
      continue;
    }

    const groupDepthMatch = arg.match(/^--group-depth=(.*)$/);
    if (groupDepthMatch) {
      const value = Number.parseInt(groupDepthMatch[1], 10);
      if (!Number.isFinite(value) || value <= 0) throw new Error(`Invalid --group-depth value: ${groupDepthMatch[1]}`);
      out.groupDepth = value;
      continue;
    }
    if (arg === "--group-depth") {
      const next = argv[i + 1];
      if (!next) throw new Error("Missing value for --group-depth");
      const value = Number.parseInt(next, 10);
      if (!Number.isFinite(value) || value <= 0) throw new Error(`Invalid --group-depth value: ${next}`);
      out.groupDepth = value;
      i++;
      continue;
    }

    throw new Error(`Unknown argument: ${arg}\n\n${usage()}`);
  }

  if (!path.isAbsolute(out.distDir)) {
    out.distDir = path.resolve(repoRoot, out.distDir);
  }

  return out;
}

/**
 * @param {string} dir
 * @returns {Promise<{ files: Array<{ absPath: string, relPath: string, sizeBytes: number, ext: string }>, totalBytes: number }>}
 */
async function scanDistDir(dir) {
  /** @type {Array<{ absPath: string, relPath: string, sizeBytes: number, ext: string }>} */
  const files = [];
  let totalBytes = 0;

  /** @type {string[]} */
  const stack = [dir];
  while (stack.length > 0) {
    const current = /** @type {string} */ (stack.pop());
    let entries;
    try {
      entries = await readdir(current, { withFileTypes: true });
    } catch (err) {
      // The dist directory shouldn't change during a report, but some environments
      // may clean build output concurrently. Skip vanished directories so the report
      // remains best-effort instead of failing with ENOENT.
      if (/** @type {any} */ (err)?.code === "ENOENT") continue;
      throw err;
    }
    for (const entry of entries) {
      const absPath = path.join(current, entry.name);
      if (entry.isDirectory()) {
        stack.push(absPath);
        continue;
      }
      if (!entry.isFile()) continue;

      let stats;
      try {
        stats = await stat(absPath);
      } catch (err) {
        if (/** @type {any} */ (err)?.code === "ENOENT") continue;
        throw err;
      }
      const sizeBytes = stats.size;
      totalBytes += sizeBytes;

      const relPath = toPosixPath(path.relative(dir, absPath));
      const ext = path.extname(entry.name).toLowerCase() || "(none)";
      files.push({ absPath, relPath, sizeBytes, ext });
    }
  }

  files.sort((a, b) => b.sizeBytes - a.sizeBytes);
  return { files, totalBytes };
}

/**
 * @param {string} distDir
 * @param {number} totalBytes
 * @param {number} totalFiles
 * @param {{
 *   totalBudgetMb: number | null,
 *   singleBudgetMb: number | null,
 *   singleFileOffenders: Array<{ relPath: string, sizeBytes: number }>,
 *   totalBudgetExceeded: boolean,
 *   totalBudgetOverByBytes: number,
 * }} budgets
 * @returns {string[]}
 */
function renderHeaderLines(distDir, totalBytes, totalFiles, budgets) {
  const { totalBudgetMb, singleBudgetMb, singleFileOffenders, totalBudgetExceeded, totalBudgetOverByBytes } =
    budgets;
  /** @type {string[]} */
  const lines = [];
  lines.push("## Desktop dist asset report");
  lines.push("");

  let displayDist = distDir;
  if (distDir.startsWith(repoRoot + path.sep)) {
    displayDist = toPosixPath(path.relative(repoRoot, distDir));
  }
  lines.push(`Dist dir: \`${displayDist}\``);
  lines.push(`Total files: **${formatInt(totalFiles)}**`);
  lines.push(`Total size: **${formatMb(totalBytes)}** (${formatInt(totalBytes)} bytes)`);
  lines.push("");

  if (totalBudgetMb === null && singleBudgetMb === null) {
    lines.push(
      "Budgets: _not configured_ (set `FORMULA_DESKTOP_DIST_TOTAL_BUDGET_MB` / `FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB` to enforce).",
    );
    lines.push("");
    return lines;
  }

  lines.push("Budgets:");
  if (totalBudgetMb !== null) {
    const status = totalBudgetExceeded
      ? `**FAIL** (over by ${formatMb(totalBudgetOverByBytes)})`
      : "PASS";
    lines.push(
      `- Total: **${totalBudgetMb} MB** (\`FORMULA_DESKTOP_DIST_TOTAL_BUDGET_MB\`) — ${status}`,
    );
  } else {
    lines.push("- Total: _(unset)_ (`FORMULA_DESKTOP_DIST_TOTAL_BUDGET_MB`)");
  }
  if (singleBudgetMb !== null) {
    const status =
      singleFileOffenders.length > 0 ? `**FAIL** (${singleFileOffenders.length} file(s))` : "PASS";
    lines.push(
      `- Single file: **${singleBudgetMb} MB** (\`FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB\`) — ${status}`,
    );
  } else {
    lines.push("- Single file: _(unset)_ (`FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB`)");
  }
  lines.push("");

  if (singleBudgetMb !== null && singleFileOffenders.length > 0) {
    lines.push(`Files over single-file budget (${singleBudgetMb} MB):`);
    const maxList = 25;
    for (const f of singleFileOffenders.slice(0, maxList)) {
      lines.push(`- \`${f.relPath}\` — ${formatBytesAndMb(f.sizeBytes)}`);
    }
    if (singleFileOffenders.length > maxList) {
      lines.push(`- … (${singleFileOffenders.length - maxList} more)`);
    }
    lines.push("");
  }

  return lines;
}

/**
 * @param {Array<{ relPath: string, sizeBytes: number, ext: string }>} files
 * @param {number} totalBytes
 * @param {number} topN
 * @param {number | null} singleBudgetMb
 * @returns {string[]}
 */
function renderTopFilesTable(files, totalBytes, topN, singleBudgetMb) {
  const topFiles = files.slice(0, topN);
  const hasSingleBudget = singleBudgetMb !== null;
  const singleBudgetBytes = hasSingleBudget ? singleBudgetMb * BYTES_PER_MB : 0;

  /** @type {string[]} */
  const lines = [];
  lines.push(`### Top ${topFiles.length} largest files`);
  lines.push("");

  if (topFiles.length === 0) {
    lines.push("_No files found._");
    lines.push("");
    return lines;
  }

  if (hasSingleBudget) {
    lines.push("| File | Type | Size | Share | Over budget |");
    lines.push("| --- | :---: | ---: | ---: | :---: |");
  } else {
    lines.push("| File | Type | Size | Share |");
    lines.push("| --- | :---: | ---: | ---: |");
  }

  for (const f of topFiles) {
    const share = totalBytes > 0 ? `${((f.sizeBytes / totalBytes) * 100).toFixed(1)}%` : "0.0%";
    if (hasSingleBudget) {
      const over = f.sizeBytes > singleBudgetBytes ? "YES" : "";
      lines.push(`| \`${f.relPath}\` | \`${f.ext}\` | ${formatBytesAndMb(f.sizeBytes)} | ${share} | ${over} |`);
    } else {
      lines.push(`| \`${f.relPath}\` | \`${f.ext}\` | ${formatBytesAndMb(f.sizeBytes)} | ${share} |`);
    }
  }

  lines.push("");
  return lines;
}

/**
 * @param {Array<{ relPath: string, sizeBytes: number }>} files
 * @param {number} totalBytes
 * @param {number} groupDepth
 * @returns {string[]}
 */
function renderGroupedTotals(files, totalBytes, groupDepth) {
  /** @type {Map<string, { bytes: number, files: number }>} */
  const groups = new Map();

  for (const f of files) {
    const parts = f.relPath.split("/");
    const dirParts = parts.slice(0, -1);
    const key = dirParts.length === 0 ? "(root)" : `${dirParts.slice(0, groupDepth).join("/")}/`;
    const prev = groups.get(key);
    if (prev) {
      prev.bytes += f.sizeBytes;
      prev.files += 1;
    } else {
      groups.set(key, { bytes: f.sizeBytes, files: 1 });
    }
  }

  const sorted = Array.from(groups.entries()).sort((a, b) => b[1].bytes - a[1].bytes);

  /** @type {string[]} */
  const lines = [];
  lines.push("### Grouped totals");
  lines.push("");
  lines.push("| Group | Files | Size | Share |");
  lines.push("| --- | ---: | ---: | ---: |");

  for (const [key, value] of sorted) {
    const share = totalBytes > 0 ? `${((value.bytes / totalBytes) * 100).toFixed(1)}%` : "0.0%";
    lines.push(
      `| \`${key}\` | ${formatInt(value.files)} | ${formatBytesAndMb(value.bytes)} | ${share} |`,
    );
  }

  lines.push("");
  return lines;
}

/**
 * @param {string} markdown
 */
async function appendStepSummary(markdown) {
  const summaryPath = process.env.GITHUB_STEP_SUMMARY;
  if (!summaryPath) return;
  try {
    await appendFile(summaryPath, `${markdown}\n`, "utf8");
  } catch (err) {
    // Non-fatal; the main report is still emitted to stdout/stderr.
    console.error(`desktop-dist: WARNING failed to append to GITHUB_STEP_SUMMARY (${summaryPath})`);
    console.error(err);
  }
}

/**
 * @param {Array<{ relPath: string, sizeBytes: number }>} files
 * @param {number} totalBytes
 * @param {number | null} totalBudgetMb
 * @param {number | null} singleBudgetMb
 */
function printBudgetFailures(files, totalBytes, totalBudgetMb, singleBudgetMb) {
  const totalBudgetBytes = totalBudgetMb !== null ? totalBudgetMb * BYTES_PER_MB : null;
  const singleBudgetBytes = singleBudgetMb !== null ? singleBudgetMb * BYTES_PER_MB : null;

  if (totalBudgetBytes !== null && totalBytes > totalBudgetBytes) {
    const overBy = totalBytes - totalBudgetBytes;
    console.error(
      `desktop-dist: ERROR total dist size ${formatMb(totalBytes)} exceeds budget ${totalBudgetMb} MB by ${formatMb(overBy)} (FORMULA_DESKTOP_DIST_TOTAL_BUDGET_MB).`,
    );
    console.error("desktop-dist: Largest files:");
    for (const f of files.slice(0, 10)) {
      console.error(`desktop-dist: - ${f.relPath} ${formatBytesAndMb(f.sizeBytes)}`);
    }
  }

  if (singleBudgetBytes !== null) {
    const offenders = files.filter((f) => f.sizeBytes > singleBudgetBytes);
    if (offenders.length > 0) {
      offenders.sort((a, b) => b.sizeBytes - a.sizeBytes);
      console.error(
        `desktop-dist: ERROR ${offenders.length} file(s) exceed single-file budget ${singleBudgetMb} MB (FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB).`,
      );
      for (const f of offenders) {
        console.error(`desktop-dist: - ${f.relPath} ${formatBytesAndMb(f.sizeBytes)}`);
      }
    }
  }
}

let exitCode = 0;
try {
  const args = parseArgs(process.argv.slice(2));

  const totalBudgetRaw = process.env.FORMULA_DESKTOP_DIST_TOTAL_BUDGET_MB;
  const singleBudgetRaw = process.env.FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB;
  const totalBudgetMb = totalBudgetRaw ? parseBudgetMb(totalBudgetRaw, "FORMULA_DESKTOP_DIST_TOTAL_BUDGET_MB") : null;
  const singleBudgetMb = singleBudgetRaw ? parseBudgetMb(singleBudgetRaw, "FORMULA_DESKTOP_DIST_SINGLE_FILE_BUDGET_MB") : null;

  let distStats;
  try {
    distStats = await stat(args.distDir);
  } catch {
    distStats = null;
  }

  if (!distStats || !distStats.isDirectory()) {
    const displayDist = args.distDir.startsWith(repoRoot + path.sep)
      ? toPosixPath(path.relative(repoRoot, args.distDir))
      : args.distDir;
    const msg = [
      "## Desktop dist asset report",
      "",
      `Dist dir: \`${displayDist}\``,
      "",
      "_dist directory not found._",
      "",
      "Hint: run `pnpm build:desktop` first (or pass `--dist-dir`).",
      "",
    ].join("\n");
    console.error(`desktop-dist: ERROR dist directory not found: ${args.distDir}`);
    console.log(msg);
    await appendStepSummary(msg);
    exitCode = 1;
  } else {
    const { files, totalBytes } = await scanDistDir(args.distDir);

    const totalBudgetBytes = totalBudgetMb !== null ? totalBudgetMb * BYTES_PER_MB : null;
    const singleBudgetBytes = singleBudgetMb !== null ? singleBudgetMb * BYTES_PER_MB : null;
    const singleFileOffenders =
      singleBudgetBytes !== null ? files.filter((f) => f.sizeBytes > singleBudgetBytes) : [];
    const totalBudgetExceeded = totalBudgetBytes !== null && totalBytes > totalBudgetBytes;
    const totalBudgetOverByBytes =
      totalBudgetExceeded && totalBudgetBytes !== null ? totalBytes - totalBudgetBytes : 0;

    const headerLines = renderHeaderLines(
      args.distDir,
      totalBytes,
      files.length,
      {
        totalBudgetMb,
        singleBudgetMb,
        singleFileOffenders,
        totalBudgetExceeded,
        totalBudgetOverByBytes,
      },
    );
    const topLines = renderTopFilesTable(files, totalBytes, args.topN, singleBudgetMb);
    const groupLines = args.groups ? renderGroupedTotals(files, totalBytes, args.groupDepth) : [];

    const markdown = [...headerLines, ...topLines, ...groupLines].join("\n");
    console.log(markdown);
    await appendStepSummary(markdown);

    if (totalBudgetExceeded || singleFileOffenders.length > 0) {
      printBudgetFailures(files, totalBytes, totalBudgetMb, singleBudgetMb);
      exitCode = 1;
    }
  }
} catch (err) {
  console.error("desktop-dist: ERROR failed to generate report.");
  console.error(err);
  exitCode = 2;
}

process.exitCode = exitCode;
