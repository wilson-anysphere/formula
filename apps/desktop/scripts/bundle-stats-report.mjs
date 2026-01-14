import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.resolve(__dirname, "..");
const defaultDistDir = path.join(desktopRoot, "dist");
const defaultStatsPath = path.join(desktopRoot, "dist", "bundle-stats.json");

const KB = 1024;
const numberFmt = new Intl.NumberFormat("en-US");

function fmtInt(n) {
  return numberFmt.format(n);
}

function fmtKiB(bytes, decimals = 1) {
  return (bytes / KB).toFixed(decimals);
}

function toPosixPath(p) {
  return p.split(path.sep).join("/");
}

function usage() {
  return [
    "Summarize rollup-plugin-visualizer output (apps/desktop/dist/bundle-stats.json).",
    "",
    "Usage:",
    "  pnpm -C apps/desktop build:analyze",
    "  pnpm -C apps/desktop report:bundle-stats",
    "",
    "Options:",
    "  --file <path>       Path to bundle-stats.json (default: dist/bundle-stats.json)",
    "  --dist <path>       Dist directory containing index.html (default: dist/)",
    "  --startup           Analyze only the initial JS loaded by dist/index.html (scripts + modulepreload).",
    "  --top <n>           Number of rows to print (default: 20)",
    "  --metric <m>        Sort metric: rendered | gzip | brotli (default: gzip)",
    "  --chunk <pattern>   Focus a specific chunk name (substring match). Default: largest chunk by metric.",
    "",
  ].join("\n");
}

/**
 * @param {string[]} argv
 */
function parseArgs(argv) {
  /** @type {{ file: string, distDir: string, startup: boolean, top: number, metric: "rendered"|"gzip"|"brotli", chunkPattern?: string }} */
  const out = { file: defaultStatsPath, distDir: defaultDistDir, startup: false, top: 20, metric: "gzip" };

  let args = argv.slice();
  // pnpm forwards a literal `--` delimiter into scripts; strip it so users can do:
  // `pnpm report:bundle-stats -- --top 50`.
  const delimiterIdx = args.indexOf("--");
  if (delimiterIdx >= 0) {
    args = [...args.slice(0, delimiterIdx), ...args.slice(delimiterIdx + 1)];
  }

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (!arg) continue;

    if (arg === "--help" || arg === "-h") {
      out.help = true;
      continue;
    }

    if (arg === "--startup") {
      out.startup = true;
      continue;
    }

    if (arg === "--file") {
      const next = args[i + 1];
      if (!next) throw new Error("Missing value for --file");
      out.file = next;
      i++;
      continue;
    }
    if (arg.startsWith("--file=")) {
      out.file = arg.slice("--file=".length);
      continue;
    }

    if (arg === "--dist") {
      const next = args[i + 1];
      if (!next) throw new Error("Missing value for --dist");
      out.distDir = next;
      i++;
      continue;
    }
    if (arg.startsWith("--dist=")) {
      out.distDir = arg.slice("--dist=".length);
      continue;
    }

    if (arg === "--top") {
      const next = args[i + 1];
      if (!next) throw new Error("Missing value for --top");
      const n = Number.parseInt(next, 10);
      if (!Number.isFinite(n) || n <= 0) throw new Error(`Invalid --top value: ${next}`);
      out.top = n;
      i++;
      continue;
    }
    if (arg.startsWith("--top=")) {
      const n = Number.parseInt(arg.slice("--top=".length), 10);
      if (!Number.isFinite(n) || n <= 0) throw new Error(`Invalid --top value: ${arg}`);
      out.top = n;
      continue;
    }

    if (arg === "--metric") {
      const next = args[i + 1];
      if (!next) throw new Error("Missing value for --metric");
      if (next !== "rendered" && next !== "gzip" && next !== "brotli") {
        throw new Error(`Invalid --metric: ${next}`);
      }
      out.metric = next;
      i++;
      continue;
    }
    if (arg.startsWith("--metric=")) {
      const m = arg.slice("--metric=".length);
      if (m !== "rendered" && m !== "gzip" && m !== "brotli") throw new Error(`Invalid --metric: ${m}`);
      out.metric = m;
      continue;
    }

    if (arg === "--chunk") {
      const next = args[i + 1];
      if (!next) throw new Error("Missing value for --chunk");
      out.chunkPattern = next;
      i++;
      continue;
    }
    if (arg.startsWith("--chunk=")) {
      out.chunkPattern = arg.slice("--chunk=".length);
      continue;
    }

    throw new Error(`Unknown argument: ${arg}\n\n${usage()}`);
  }

  // Resolve relative to the desktop root (so `--file dist/bundle-stats.json` works
  // from both repo root and apps/desktop).
  if (!path.isAbsolute(out.file)) {
    out.file = path.resolve(desktopRoot, out.file);
  }
  if (!path.isAbsolute(out.distDir)) {
    out.distDir = path.resolve(desktopRoot, out.distDir);
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

/**
 * @param {string} modulePath
 */
function inferGroup(modulePath) {
  const parts = modulePath.split("/");
  const nmIdx = parts.lastIndexOf("node_modules");
  if (nmIdx >= 0 && nmIdx + 1 < parts.length) {
    const first = parts[nmIdx + 1];
    if (first.startsWith("@") && nmIdx + 2 < parts.length) {
      return `${first}/${parts[nmIdx + 2]}`;
    }
    return first;
  }
  if (parts[0] === "packages" || parts[0] === "apps") {
    return parts.slice(0, 2).join("/");
  }
  return parts[0] || modulePath;
}

/**
 * @param {any} node
 * @param {Record<string, any>} nodeParts
 * @returns {{ rendered: number, gzip: number, brotli: number, leaves: number }}
 */
function sumNode(node, nodeParts) {
  /** @type {{ rendered: number, gzip: number, brotli: number, leaves: number }} */
  const out = { rendered: 0, gzip: 0, brotli: 0, leaves: 0 };

  if (node?.uid && nodeParts[node.uid]) {
    out.rendered += nodeParts[node.uid].renderedLength || 0;
    out.gzip += nodeParts[node.uid].gzipLength || 0;
    out.brotli += nodeParts[node.uid].brotliLength || 0;
    out.leaves += 1;
  }

  for (const child of node?.children || []) {
    const c = sumNode(child, nodeParts);
    out.rendered += c.rendered;
    out.gzip += c.gzip;
    out.brotli += c.brotli;
    out.leaves += c.leaves;
  }

  return out;
}

/**
 * @param {any} node
 * @param {Record<string, any>} nodeParts
 * @param {string[]} prefix
 * @param {Array<{ modulePath: string, rendered: number, gzip: number, brotli: number }>} out
 */
function collectLeafSizes(node, nodeParts, prefix, out) {
  const nextPrefix = [...prefix, node.name];
  if (node?.uid && nodeParts[node.uid]) {
    const modulePath = nextPrefix.slice(1).join("/");
    const part = nodeParts[node.uid];
    out.push({
      modulePath,
      rendered: part.renderedLength || 0,
      gzip: part.gzipLength || 0,
      brotli: part.brotliLength || 0,
    });
  }
  for (const child of node?.children || []) {
    collectLeafSizes(child, nodeParts, nextPrefix, out);
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    console.log(usage());
    process.exit(0);
  }

  const raw = await readFile(args.file, "utf8");
  /** @type {{ tree: any, nodeParts: Record<string, any> }} */
  const data = JSON.parse(raw);
  if (!data?.tree || !data?.nodeParts) {
    throw new Error(`Unexpected bundle-stats.json shape: ${args.file}`);
  }

  const chunks = (data.tree.children || []).map((chunk) => {
    const s = sumNode(chunk, data.nodeParts);
    return { name: chunk.name, ...s, node: chunk };
  });

  const metricKey = args.metric;
  chunks.sort((a, b) => (b[metricKey] || 0) - (a[metricKey] || 0));

  console.log(`[bundle-stats] ${toPosixPath(path.relative(desktopRoot, args.file))}`);
  console.log("");
  console.log(`Top chunks (sorted by ${args.metric}):`);

  for (const chunk of chunks.slice(0, args.top)) {
    console.log(
      [
        `${fmtKiB(chunk[args.metric]).padStart(8)} KiB`,
        `${chunk.name}`.padEnd(40),
        `(rendered ${fmtKiB(chunk.rendered)} KiB, gzip ${fmtKiB(chunk.gzip)} KiB, brotli ${fmtKiB(chunk.brotli)} KiB)`,
      ].join(" "),
    );
  }

  let startupEntryRelPaths = [];
  let startupPreloadRelPaths = [];
  /** @type {Set<string> | null} */
  let startupRelPaths = null;
  /** @type {Array<any> | null} */
  let startupChunks = null;

  if (args.startup) {
    const indexHtmlPath = path.join(args.distDir, "index.html");
    const indexHtml = await readFile(indexHtmlPath, "utf8");
    startupEntryRelPaths = extractScriptSrcs(indexHtml);
    startupPreloadRelPaths = extractModulePreloadHrefs(indexHtml);
    startupRelPaths = new Set([...startupEntryRelPaths, ...startupPreloadRelPaths]);
    startupChunks = chunks.filter((c) => startupRelPaths.has(String(c.name)));

    console.log("");
    console.log(`Startup chunks (from ${toPosixPath(path.relative(desktopRoot, indexHtmlPath))}):`);

    /** @type {string[]} */
    const missing = [];
    for (const rel of startupRelPaths) {
      if (!chunks.some((c) => String(c.name) === rel)) missing.push(rel);
    }
    if (missing.length > 0) {
      console.log("  (warning) index.html references JS files not present in bundle-stats:");
      for (const rel of missing) console.log(`  - ${rel}`);
      console.log("");
    }

    const row = (label, rel) => {
      const c = chunks.find((chunk) => String(chunk.name) === rel);
      if (!c) return;
      console.log(
        [
          `${fmtKiB(c[metricKey]).padStart(8)} KiB`,
          `${label}`.padEnd(16),
          `${c.name}`,
          `(rendered ${fmtKiB(c.rendered)} KiB, gzip ${fmtKiB(c.gzip)} KiB, brotli ${fmtKiB(c.brotli)} KiB)`,
        ].join(" "),
      );
    };

    for (const rel of startupEntryRelPaths) row("entry", rel);
    for (const rel of startupPreloadRelPaths) row("modulepreload", rel);

    if (startupChunks.length > 0) {
      const total = startupChunks.reduce(
        (acc, c) => {
          acc.rendered += c.rendered;
          acc.gzip += c.gzip;
          acc.brotli += c.brotli;
          return acc;
        },
        { rendered: 0, gzip: 0, brotli: 0 },
      );
      console.log(
        `\nStartup total: rendered ${fmtKiB(total.rendered)} KiB, gzip ${fmtKiB(total.gzip)} KiB, brotli ${fmtKiB(
          total.brotli,
        )} KiB`,
      );
    }
  }

  const focusChunks = (() => {
    if (args.startup) {
      const candidates = startupChunks ?? [];
      if (args.chunkPattern != null) {
        return candidates.filter((c) => String(c.name).includes(args.chunkPattern));
      }
      return candidates;
    }
    if (args.chunkPattern != null) {
      const hit = chunks.find((c) => String(c.name).includes(args.chunkPattern));
      return hit ? [hit] : [];
    }
    return chunks[0] ? [chunks[0]] : [];
  })();

  if (focusChunks.length === 0) {
    console.log("");
    console.log("[bundle-stats] No chunks found.");
    return;
  }

  console.log("");
  if (focusChunks.length === 1) {
    console.log(`Focus chunk: ${focusChunks[0].name}`);
  } else {
    console.log(`Focus chunks (${focusChunks.length}):`);
    for (const c of focusChunks) console.log(`- ${c.name}`);
  }

  /** @type {Array<{ modulePath: string, rendered: number, gzip: number, brotli: number }>} */
  const leaves = [];
  for (const c of focusChunks) {
    collectLeafSizes(c.node, data.nodeParts, [], leaves);
  }

  /** @type {Map<string, { rendered: number, gzip: number, brotli: number }>} */
  const byModule = new Map();
  for (const leaf of leaves) {
    const prev = byModule.get(leaf.modulePath) || { rendered: 0, gzip: 0, brotli: 0 };
    prev.rendered += leaf.rendered;
    prev.gzip += leaf.gzip;
    prev.brotli += leaf.brotli;
    byModule.set(leaf.modulePath, prev);
  }
  const mergedLeaves = Array.from(byModule.entries()).map(([modulePath, s]) => ({ modulePath, ...s }));

  /** @type {Map<string, { rendered: number, gzip: number, brotli: number }>} */
  const groupTotals = new Map();
  for (const leaf of mergedLeaves) {
    const group = inferGroup(leaf.modulePath);
    const prev = groupTotals.get(group) || { rendered: 0, gzip: 0, brotli: 0 };
    prev.rendered += leaf.rendered;
    prev.gzip += leaf.gzip;
    prev.brotli += leaf.brotli;
    groupTotals.set(group, prev);
  }

  const groups = Array.from(groupTotals.entries()).map(([name, s]) => ({ name, ...s }));
  groups.sort((a, b) => (b[metricKey] || 0) - (a[metricKey] || 0));

  console.log("");
  const focusLabel = focusChunks.length === 1 ? focusChunks[0].name : `${focusChunks.length} chunk(s)`;
  console.log(`Top dependency groups in ${focusLabel} (sorted by ${args.metric}):`);
  for (const g of groups.slice(0, args.top)) {
    console.log(
      [
        `${fmtKiB(g[metricKey]).padStart(8)} KiB`,
        g.name,
        `(rendered ${fmtKiB(g.rendered)} KiB, gzip ${fmtKiB(g.gzip)} KiB, brotli ${fmtKiB(g.brotli)} KiB)`,
      ].join(" "),
    );
  }

  mergedLeaves.sort((a, b) => (b[metricKey] || 0) - (a[metricKey] || 0));
  console.log("");
  console.log(`Top modules in ${focusLabel} (sorted by ${args.metric}):`);
  for (const leaf of mergedLeaves.slice(0, args.top)) {
    console.log(
      [
        `${fmtKiB(leaf[metricKey]).padStart(8)} KiB`,
        toPosixPath(leaf.modulePath),
        `(rendered ${fmtKiB(leaf.rendered)} KiB, gzip ${fmtKiB(leaf.gzip)} KiB, brotli ${fmtKiB(leaf.brotli)} KiB)`,
      ].join(" "),
    );
  }

  console.log("");
  console.log(
    `[bundle-stats] Done. Parsed ${fmtInt(mergedLeaves.length)} module(s) across ${fmtInt(chunks.length)} chunk(s).`,
  );
}

main().catch((err) => {
  console.error("[bundle-stats] ERROR:", err);
  process.exit(1);
});
