import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const repoRootReal = fs.realpathSync(repoRoot);
const defaultConfigPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");
const configPath = process.env.FORMULA_TAURI_CONF_PATH
  ? path.resolve(process.env.FORMULA_TAURI_CONF_PATH)
  : defaultConfigPath;

function safeRealpath(p) {
  try {
    return fs.realpathSync(p);
  } catch {
    return path.resolve(p);
  }
}

function isPathWithinDir(p, dir) {
  const rel = path.relative(dir, p);
  return rel === "" || (!rel.startsWith("..") && !path.isAbsolute(rel));
}

const configPathReal = safeRealpath(configPath);
const configIsInRepo = isPathWithinDir(configPathReal, repoRootReal);
const expectedRepoLicenseReal = configIsInRepo ? safeRealpath(path.join(repoRootReal, "LICENSE")) : "";
const expectedRepoNoticeReal = configIsInRepo ? safeRealpath(path.join(repoRootReal, "NOTICE")) : "";

function die(message) {
  console.error(`desktop-compliance: ERROR ${message}`);
  process.exit(1);
}

/**
 * @param {unknown} value
 * @returns {Record<string, unknown> | null}
 */
function asObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  // @ts-ignore - runtime type guard
  return value;
}

function normalizeExt(value) {
  return String(value).trim().toLowerCase().replace(/^\./, "");
}

/**
 * Parse shared-mime-info XML content into a mapping of:
 *   mimeType -> Set(glob patterns)
 *
 * This is intentionally lightweight (regex-based) to keep the guard dependency-free.
 *
 * @param {string} xml
 * @returns {Map<string, Set<string>>}
 */
function parseSharedMimeInfoXml(xml) {
  /** @type {Map<string, Set<string>>} */
  const out = new Map();
  if (typeof xml !== "string" || !xml.trim()) return out;

  const mimeTypeRe = /<mime-type\b[^>]*\btype\s*=\s*(['"])(.*?)\1[^>]*>([\s\S]*?)<\/mime-type>/gi;
  while (true) {
    const match = mimeTypeRe.exec(xml);
    if (!match) break;
    const type = String(match[2] ?? "").trim().toLowerCase();
    if (!type) continue;
    const inner = String(match[3] ?? "");
    const patterns = out.get(type) ?? new Set();

    const globRe = /<glob\b[^>]*\bpattern\s*=\s*(['"])(.*?)\1[^>]*\/?>/gi;
    while (true) {
      const g = globRe.exec(inner);
      if (!g) break;
      const pattern = String(g[2] ?? "").trim().toLowerCase();
      if (pattern) patterns.add(pattern);
    }

    out.set(type, patterns);
  }

  return out;
}

/**
 * @param {unknown} resources
 * @param {string} filename
 * @returns {boolean}
 */
function resourcesInclude(resources, filename) {
  if (!Array.isArray(resources)) return false;
  return resources.some((entry) => typeof entry === "string" && path.basename(entry) === filename);
}

/**
 * @param {string} src
 * @param {string} filename
 * @returns {string}
 */
function resolveAndAssertFileExists(src, filename) {
  const base = path.basename(src);
  if (base !== filename) {
    die(`expected source path basename to be ${filename} (got ${JSON.stringify(src)})`);
  }

  const resolved = path.isAbsolute(src) ? src : path.resolve(path.dirname(configPath), src);
  let stat;
  try {
    stat = fs.statSync(resolved);
  } catch (err) {
    die(`missing source file for ${filename}: ${resolved}`);
  }
  if (!stat.isFile()) {
    die(`expected source file for ${filename} to be a file: ${resolved}`);
  }
  return resolved;
}

function assertRepoRootComplianceFile(resolvedPath, filename, context) {
  if (!configIsInRepo) return;
  const resolvedReal = safeRealpath(resolvedPath);
  const expected =
    filename === "LICENSE"
      ? expectedRepoLicenseReal
      : filename === "NOTICE"
        ? expectedRepoNoticeReal
        : "";
  if (!expected) return;
  if (resolvedReal !== expected) {
    die(
      `${context} must reference the repo root ${filename} (${expected}). Found: ${resolvedReal}. ` +
        `This check ensures the distributed bundles ship the top-level LICENSE/NOTICE.`
    );
  }
}

/**
 * @param {unknown} files
 * @param {string} mainBinaryName
 * @param {string} kind
 */
function validateLinuxFiles(files, mainBinaryName, kind) {
  const obj = asObject(files);
  if (!obj) {
    die(`bundle.linux.${kind}.files must be an object mapping destination -> source`);
  }

  const requiredDest = {
    LICENSE: `usr/share/doc/${mainBinaryName}/LICENSE`,
    NOTICE: `usr/share/doc/${mainBinaryName}/NOTICE`,
  };

  for (const [filename, destPath] of Object.entries(requiredDest)) {
    if (!(destPath in obj)) {
      const keys = Object.keys(obj);
      die(
        `bundle.linux.${kind}.files is missing ${JSON.stringify(destPath)} (for ${filename}). Present keys: ${JSON.stringify(keys)}`,
      );
    }
    const src = obj[destPath];
    if (typeof src !== "string") {
      die(
        `bundle.linux.${kind}.files[${JSON.stringify(destPath)}] must be a string source path (got ${typeof src})`,
      );
    }
    const resolved = resolveAndAssertFileExists(src, filename);
    assertRepoRootComplianceFile(resolved, filename, `bundle.linux.${kind}.files[${JSON.stringify(destPath)}]`);
  }
}

/**
 * @param {unknown} fileAssociations
 * @returns {boolean}
 */
function hasParquetAssociation(fileAssociations) {
  if (!Array.isArray(fileAssociations)) return false;
  for (const assoc of fileAssociations) {
    if (!assoc || typeof assoc !== "object" || Array.isArray(assoc)) continue;
    // @ts-ignore - runtime inspection
    const mimeType = assoc.mimeType;
    if (
      typeof mimeType === "string" &&
      mimeType.trim().toLowerCase() === "application/vnd.apache.parquet"
    ) {
      return true;
    }
    // @ts-ignore - runtime inspection
    const ext = assoc.ext;
    if (Array.isArray(ext) && ext.some((v) => typeof v === "string" && v.trim().toLowerCase() === "parquet")) {
      return true;
    }
  }
  return false;
}

/**
 * @param {unknown} depends
 * @param {string} kind
 */
function validateLinuxDepends(depends, kind) {
  if (!Array.isArray(depends)) {
    die(`bundle.linux.${kind}.depends must be an array of package dependency strings`);
  }
  const deps = depends.map((v) => String(v));
  if (!deps.some((v) => v.toLowerCase().includes("shared-mime-info"))) {
    die(`bundle.linux.${kind}.depends must include shared-mime-info (required for MIME database integration)`);
  }
}

/**
 * @param {unknown} files
 * @param {unknown} fileAssociations
 * @param {string} kind
 */
function validateLinuxMimeFiles(files, fileAssociations, kind, identifier) {
  const obj = asObject(files);
  if (!obj) {
    die(`bundle.linux.${kind}.files must be an object mapping destination -> source`);
  }
  const normalizedIdentifier = String(identifier ?? "").trim();
  if (!normalizedIdentifier) {
    die("tauri.conf.json identifier must be a non-empty string when Parquet file association is configured");
  }
  // `identifier` becomes the shared-mime-info XML filename on Linux:
  //   /usr/share/mime/packages/<identifier>.xml
  // Guard against path separators so we don't accidentally allow surprising nested paths or
  // fail later with confusing basename mismatches.
  if (normalizedIdentifier.includes("/") || normalizedIdentifier.includes("\\")) {
    die(
      "tauri.conf.json identifier must be a valid filename when Parquet file association is configured " +
        `(no '/' or '\\\\' path separators). Found: ${JSON.stringify(normalizedIdentifier)}`,
    );
  }

  const mimeFilename = `${normalizedIdentifier}.xml`;
  const dest = `usr/share/mime/packages/${mimeFilename}`;
  if (!(dest in obj)) {
    const keys = Object.keys(obj);
    die(
      `bundle.linux.${kind}.files is missing ${JSON.stringify(dest)} (required for Parquet/shared-mime-info integration). Present keys: ${JSON.stringify(keys)}`,
    );
  }
  const src = obj[dest];
  if (typeof src !== "string") {
    die(`bundle.linux.${kind}.files[${JSON.stringify(dest)}] must be a string source path (got ${typeof src})`);
  }
  const resolved = resolveAndAssertFileExists(src, mimeFilename);
  let xml = "";
  try {
    xml = fs.readFileSync(resolved, "utf8");
  } catch (err) {
    die(`failed to read Parquet shared-mime-info definition file: ${resolved} (${err instanceof Error ? err.message : err})`);
  }

  if (!Array.isArray(fileAssociations) || fileAssociations.length === 0) {
    die("bundle.fileAssociations must be an array when Parquet file association is configured");
  }

  /** @type {Map<string, string>} */
  const mimeByExt = new Map();
  for (const assoc of fileAssociations) {
    if (!assoc || typeof assoc !== "object" || Array.isArray(assoc)) continue;
    // @ts-ignore
    const mimeTypeRaw = assoc.mimeType;
    const mimeType = typeof mimeTypeRaw === "string" ? mimeTypeRaw.trim().toLowerCase() : "";
    if (!mimeType) continue;
    // @ts-ignore
    const extRaw = assoc.ext;
    const exts = Array.isArray(extRaw) ? extRaw : typeof extRaw === "string" ? [extRaw] : [];
    for (const extValue of exts) {
      const ext = normalizeExt(extValue);
      if (!ext) continue;
      const existing = mimeByExt.get(ext);
      if (existing && existing !== mimeType) {
        die(
          `bundle.fileAssociations contains conflicting mimeType entries for .${ext}: ` +
            `${existing} vs ${mimeType}. Fix: ensure each extension maps to a single MIME type.`,
        );
      }
      mimeByExt.set(ext, mimeType);
    }
  }

  const parsed = parseSharedMimeInfoXml(xml);
  /** @type {Array<{ ext: string, mimeType: string, expectedGlob: string }>} */
  const missing = [];
  for (const [ext, mimeType] of mimeByExt.entries()) {
    const expectedGlob = `*.${ext}`;
    const patterns = parsed.get(mimeType) ?? new Set();
    if (!patterns.has(expectedGlob)) {
      missing.push({ ext, mimeType, expectedGlob });
    }
  }

  if (missing.length > 0) {
    const formatted = missing
      .sort((a, b) => a.ext.localeCompare(b.ext))
      .map(({ ext, mimeType, expectedGlob }) => `- .${ext} → ${mimeType} (expected glob ${expectedGlob})`)
      .join("\n");
    die(
      `shared-mime-info definition file is missing required glob mappings: ${resolved}\n` +
        `Missing:\n${formatted}\n` +
        "Fix: add <glob pattern=\"*.ext\" /> entries to mime/<identifier>.xml for all configured file associations so Linux file managers can resolve extensions to the advertised MIME types.",
    );
  }
}

let raw = "";
try {
  raw = fs.readFileSync(configPath, "utf8");
} catch (err) {
  die(`failed to read config: ${configPath}\n${err}`);
}

let config;
try {
  config = JSON.parse(raw);
} catch (err) {
  die(`failed to parse JSON: ${configPath}\n${err}`);
}

const mainBinaryName = String(config?.mainBinaryName ?? "").trim() || "formula-desktop";
const identifier = String(config?.identifier ?? "").trim();
const bundle = asObject(config?.bundle) ?? {};

const resources = bundle.resources;
if (!resourcesInclude(resources, "LICENSE")) {
  die(
    `bundle.resources must include LICENSE (expected an entry whose basename is "LICENSE"). Found: ${JSON.stringify(resources ?? null)}`,
  );
}
if (!resourcesInclude(resources, "NOTICE")) {
  die(
    `bundle.resources must include NOTICE (expected an entry whose basename is "NOTICE"). Found: ${JSON.stringify(resources ?? null)}`,
  );
}

const linux = asObject(bundle.linux) ?? {};
validateLinuxFiles(asObject(linux.deb)?.files, mainBinaryName, "deb");
validateLinuxFiles(asObject(linux.rpm)?.files, mainBinaryName, "rpm");

// Tauri's config key is `appimage` (lowercase).
validateLinuxFiles(asObject(linux.appimage)?.files, mainBinaryName, "appimage");

// When the app advertises Parquet support on Linux, we also ship a shared-mime-info XML definition
// so the MIME database can map `*.parquet` correctly even on distros whose shared-mime-info
// package does not include a Parquet glob by default.
if (hasParquetAssociation(bundle.fileAssociations)) {
  validateLinuxMimeFiles(asObject(linux.deb)?.files, bundle.fileAssociations, "deb", identifier);
  validateLinuxMimeFiles(asObject(linux.rpm)?.files, bundle.fileAssociations, "rpm", identifier);
  validateLinuxMimeFiles(asObject(linux.appimage)?.files, bundle.fileAssociations, "appimage", identifier);

  // Ensure the distro-native packages pull in `update-mime-database` on install.
  validateLinuxDepends(asObject(linux.deb)?.depends, "deb");
  validateLinuxDepends(asObject(linux.rpm)?.depends, "rpm");
}

// Ensure the declared resources actually exist on disk as files.
if (!Array.isArray(resources)) {
  die("bundle.resources must be an array");
}
const licenseEntry = resources.find((entry) => typeof entry === "string" && path.basename(entry) === "LICENSE");
const noticeEntry = resources.find((entry) => typeof entry === "string" && path.basename(entry) === "NOTICE");
if (typeof licenseEntry !== "string") {
  die("bundle.resources must contain an entry for LICENSE");
}
if (typeof noticeEntry !== "string") {
  die("bundle.resources must contain an entry for NOTICE");
}
const resolvedLicense = resolveAndAssertFileExists(licenseEntry, "LICENSE");
const resolvedNotice = resolveAndAssertFileExists(noticeEntry, "NOTICE");
assertRepoRootComplianceFile(resolvedLicense, "LICENSE", "bundle.resources entry for LICENSE");
assertRepoRootComplianceFile(resolvedNotice, "NOTICE", "bundle.resources entry for NOTICE");

console.log(`desktop-compliance: OK (LICENSE/NOTICE bundling configured for ${mainBinaryName})`);
