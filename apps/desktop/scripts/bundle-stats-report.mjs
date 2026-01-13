import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.resolve(__dirname, "..");
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
  /** @type {{ file: string, top: number, metric: "rendered"|"gzip"|"brotli", chunkPattern?: string }} */
  const out = { file: defaultStatsPath, top: 20, metric: "gzip" };

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

  const focus =
    args.chunkPattern != null
      ? chunks.find((c) => String(c.name).includes(args.chunkPattern))
      : chunks[0];
  if (!focus) {
    console.log("");
    console.log("[bundle-stats] No chunks found.");
    return;
  }

  console.log("");
  console.log(`Focus chunk: ${focus.name}`);

  /** @type {Array<{ modulePath: string, rendered: number, gzip: number, brotli: number }>} */
  const leaves = [];
  collectLeafSizes(focus.node, data.nodeParts, [], leaves);

  /** @type {Map<string, { rendered: number, gzip: number, brotli: number }>} */
  const groupTotals = new Map();
  for (const leaf of leaves) {
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
  console.log(`Top dependency groups in ${focus.name} (sorted by ${args.metric}):`);
  for (const g of groups.slice(0, args.top)) {
    console.log(
      [
        `${fmtKiB(g[metricKey]).padStart(8)} KiB`,
        g.name,
        `(rendered ${fmtKiB(g.rendered)} KiB, gzip ${fmtKiB(g.gzip)} KiB, brotli ${fmtKiB(g.brotli)} KiB)`,
      ].join(" "),
    );
  }

  leaves.sort((a, b) => (b[metricKey] || 0) - (a[metricKey] || 0));
  console.log("");
  console.log(`Top modules in ${focus.name} (sorted by ${args.metric}):`);
  for (const leaf of leaves.slice(0, args.top)) {
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
    `[bundle-stats] Done. Parsed ${fmtInt(leaves.length)} module(s) across ${fmtInt(chunks.length)} chunk(s).`,
  );
}

main().catch((err) => {
  console.error("[bundle-stats] ERROR:", err);
  process.exit(1);
});

