#!/usr/bin/env node
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));

/**
 * Resolve a path that may be absolute or repo-relative.
 * @param {string | undefined} value
 * @param {string} fallback
 */
function resolvePath(value, fallback) {
  if (value && String(value).trim()) {
    const p = String(value).trim();
    return path.isAbsolute(p) ? p : path.join(repoRoot, p);
  }
  return fallback;
}

const tauriConfigPath = resolvePath(
  process.env.FORMULA_TAURI_CONF_PATH,
  path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json"),
);
const infoPlistPath = resolvePath(
  process.env.FORMULA_INFO_PLIST_PATH,
  path.join(repoRoot, "apps", "desktop", "src-tauri", "Info.plist"),
);

const REQUIRED_SCHEME = "formula";
// Desktop builds are expected to open common spreadsheet/data file formats.
// Keep this list stable and explicit so CI fails if we accidentally drop
// associations from `bundle.fileAssociations` (packaging regressions are
// high-impact for end-user UX).
const REQUIRED_FILE_EXTS = [
  "xlsx",
  "xls",
  "xlt",
  "xla",
  "xlsm",
  "xltx",
  "xltm",
  "xlam",
  "xlsb",
  "csv",
  "parquet",
];

/**
 * @param {string} value
 */
function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * @param {string} message
 */
function err(message) {
  process.exitCode = 1;
  console.error(message);
}

/**
 * @param {string} heading
 * @param {string[]} details
 */
function errBlock(heading, details) {
  err(`\n${heading}\n${details.map((d) => `  - ${d}`).join("\n")}`);
}

/**
 * Best-effort extraction of the `<array>...</array>` block immediately following a
 * `<key>...</key>` in an XML plist.
 *
 * This avoids false positives where a `<string>xlsx</string>` appears elsewhere
 * (e.g. UT*TypeDeclarations) but not under CFBundleDocumentTypes.
 *
 * @param {string} plistXml
 * @param {string} keyName
 * @returns {string | null}
 */
function extractPlistArrayBlock(plistXml, keyName) {
  const keyRe = new RegExp(`<key>\\s*${escapeRegExp(keyName)}\\s*<\\/key>`, "i");
  const keyMatch = keyRe.exec(plistXml);
  if (!keyMatch || keyMatch.index == null) return null;

  // Find the `<array>` that follows this key.
  const afterKeyIdx = keyMatch.index + keyMatch[0].length;
  const arrayOpenRe = /<array\b[^>]*>/gi;
  arrayOpenRe.lastIndex = afterKeyIdx;
  const openMatch = arrayOpenRe.exec(plistXml);
  if (!openMatch || openMatch.index == null) return null;

  const startIdx = openMatch.index;

  // Scan forward tracking nested <array> depth until we close the initial array.
  const tagRe = /<\/?array\b[^>]*>/gi;
  tagRe.lastIndex = startIdx;
  let depth = 0;
  let endIdx = -1;
  while (true) {
    const m = tagRe.exec(plistXml);
    if (!m || m.index == null) break;

    const tag = m[0].toLowerCase();
    if (tag.startsWith("</array")) {
      depth -= 1;
      if (depth === 0) {
        endIdx = m.index + m[0].length;
        break;
      }
    } else {
      depth += 1;
    }
  }

  if (endIdx < 0) return null;
  return plistXml.slice(startIdx, endIdx);
}

/**
 * Best-effort extraction of the `<dict>...</dict>` block immediately following a
 * `<key>...</key>` in an XML plist.
 *
 * @param {string} plistXml
 * @param {string} keyName
 * @returns {string | null}
 */
function extractPlistDictBlock(plistXml, keyName) {
  const keyRe = new RegExp(`<key>\\s*${escapeRegExp(keyName)}\\s*<\\/key>`, "i");
  const keyMatch = keyRe.exec(plistXml);
  if (!keyMatch || keyMatch.index == null) return null;

  const afterKeyIdx = keyMatch.index + keyMatch[0].length;
  const dictOpenRe = /<dict\b[^>]*>/gi;
  dictOpenRe.lastIndex = afterKeyIdx;
  const openMatch = dictOpenRe.exec(plistXml);
  if (!openMatch || openMatch.index == null) return null;

  const startIdx = openMatch.index;

  const tagRe = /<\/?dict\b[^>]*>/gi;
  tagRe.lastIndex = startIdx;
  let depth = 0;
  let endIdx = -1;
  while (true) {
    const m = tagRe.exec(plistXml);
    if (!m || m.index == null) break;

    const tag = m[0].toLowerCase();
    if (tag.startsWith("</dict")) {
      depth -= 1;
      if (depth === 0) {
        endIdx = m.index + m[0].length;
        break;
      }
    } else {
      depth += 1;
    }
  }

  if (endIdx < 0) return null;
  return plistXml.slice(startIdx, endIdx);
}

/**
 * Extract all `<array>...</array>` blocks that immediately follow `<key>${keyName}</key>`
 * occurrences inside a larger plist XML snippet.
 *
 * @param {string} plistXml
 * @param {string} keyName
 * @returns {string[]}
 */
function extractAllPlistArrayBlocks(plistXml, keyName) {
  /** @type {string[]} */
  const out = [];
  const keyRe = new RegExp(`<key>\\s*${escapeRegExp(keyName)}\\s*<\\/key>`, "gi");
  while (true) {
    const keyMatch = keyRe.exec(plistXml);
    if (!keyMatch || keyMatch.index == null) break;

    const afterKeyIdx = keyMatch.index + keyMatch[0].length;
    const arrayOpenRe = /<array\b[^>]*>/gi;
    arrayOpenRe.lastIndex = afterKeyIdx;
    const openMatch = arrayOpenRe.exec(plistXml);
    if (!openMatch || openMatch.index == null) continue;

    const startIdx = openMatch.index;
    const tagRe = /<\/?array\b[^>]*>/gi;
    tagRe.lastIndex = startIdx;
    let depth = 0;
    let endIdx = -1;
    while (true) {
      const m = tagRe.exec(plistXml);
      if (!m || m.index == null) break;
      const tag = m[0].toLowerCase();
      if (tag.startsWith("</array")) {
        depth -= 1;
        if (depth === 0) {
          endIdx = m.index + m[0].length;
          break;
        }
      } else {
        depth += 1;
      }
    }
    if (endIdx < 0) continue;
    out.push(plistXml.slice(startIdx, endIdx));
    // Continue scanning after the array we just captured.
    keyRe.lastIndex = endIdx;
  }
  return out;
}

/**
 * Extract `<string>...</string>` values from a snippet of plist XML.
 * @param {string} xml
 * @returns {string[]}
 */
function extractPlistStrings(xml) {
  /** @type {string[]} */
  const out = [];
  const re = /<string\b[^>]*>\s*([^<]*?)\s*<\/string>/gi;
  while (true) {
    const m = re.exec(xml);
    if (!m) break;
    const val = String(m[1] ?? "").trim();
    if (val) out.push(val);
  }
  return out;
}

/**
 * Extract top-level `<dict>...</dict>` blocks inside an `<array>...</array>` snippet.
 *
 * @param {string} arrayBlock
 * @returns {string[]}
 */
function extractTopLevelDictBlocks(arrayBlock) {
  /** @type {string[]} */
  const out = [];
  const arrayOpenEnd = arrayBlock.indexOf(">");
  const startScan = arrayOpenEnd >= 0 ? arrayOpenEnd + 1 : 0;

  const tagRe = /<\/?dict\b[^>]*>/gi;
  tagRe.lastIndex = startScan;

  let depth = 0;
  let dictStart = -1;
  while (true) {
    const m = tagRe.exec(arrayBlock);
    if (!m || m.index == null) break;
    const tag = m[0].toLowerCase();
    if (tag.startsWith("</dict")) {
      if (depth === 1 && dictStart >= 0) {
        out.push(arrayBlock.slice(dictStart, m.index + m[0].length));
        dictStart = -1;
      }
      depth -= 1;
    } else {
      depth += 1;
      if (depth === 1) {
        dictStart = m.index;
      }
    }
  }
  return out;
}

/**
 * Extract a string value from a `<dict>` snippet.
 * @param {string} dictXml
 * @param {string} keyName
 * @returns {string}
 */
function extractPlistStringValue(dictXml, keyName) {
  const keyRe = new RegExp(`<key>\\s*${escapeRegExp(keyName)}\\s*<\\/key>`, "i");
  const keyMatch = keyRe.exec(dictXml);
  if (!keyMatch || keyMatch.index == null) return "";
  const afterKeyIdx = keyMatch.index + keyMatch[0].length;
  const stringRe = /<string\b[^>]*>\s*([^<]*?)\s*<\/string>/gi;
  stringRe.lastIndex = afterKeyIdx;
  const m = stringRe.exec(dictXml);
  if (!m) return "";
  return String(m[1] ?? "").trim();
}

/**
 * Extract either `<string>` or `<array><string>..</string></array>` values from a dict snippet.
 * @param {string} dictXml
 * @param {string} keyName
 * @returns {string[]}
 */
function extractPlistStringOrArrayValues(dictXml, keyName) {
  const array = extractPlistArrayBlock(dictXml, keyName);
  if (array) return extractPlistStrings(array);
  const single = extractPlistStringValue(dictXml, keyName);
  return single ? [single] : [];
}

/**
 * Best-effort: determine which file extensions are registered by Info.plist.
 *
 * We accept extensions registered:
 *   - directly via CFBundleDocumentTypes/CFBundleTypeExtensions
 *   - indirectly via CFBundleDocumentTypes/LSItemContentTypes referencing UTIs whose
 *     UT*TypeDeclarations define public.filename-extension.
 *
 * @param {string} plistXml
 */
function collectMacosRegisteredExtensions(plistXml) {
  const docTypesBlock = extractPlistArrayBlock(plistXml, "CFBundleDocumentTypes");
  /** @type {Set<string>} */
  const docExts = new Set();
  /** @type {Set<string>} */
  const docUtis = new Set();
  /** @type {Map<string, Set<string>>} */
  const utiToExts = new Map();

  if (docTypesBlock) {
    for (const block of extractAllPlistArrayBlocks(docTypesBlock, "CFBundleTypeExtensions")) {
      for (const raw of extractPlistStrings(block)) {
        const ext = raw.trim().toLowerCase().replace(/^\./, "");
        if (ext) docExts.add(ext);
      }
    }
    for (const block of extractAllPlistArrayBlocks(docTypesBlock, "LSItemContentTypes")) {
      for (const raw of extractPlistStrings(block)) {
        const uti = raw.trim().toLowerCase();
        if (uti) docUtis.add(uti);
      }
    }
  }

  for (const key of ["UTExportedTypeDeclarations", "UTImportedTypeDeclarations"]) {
    const declsBlock = extractPlistArrayBlock(plistXml, key);
    if (!declsBlock) continue;
    for (const declDict of extractTopLevelDictBlocks(declsBlock)) {
      const uti = extractPlistStringValue(declDict, "UTTypeIdentifier").trim().toLowerCase();
      if (!uti) continue;
      const tagSpec = extractPlistDictBlock(declDict, "UTTypeTagSpecification");
      if (!tagSpec) continue;
      const extValues = extractPlistStringOrArrayValues(tagSpec, "public.filename-extension");
      for (const raw of extValues) {
        const ext = raw.trim().toLowerCase().replace(/^\./, "");
        if (!ext) continue;
        const set = utiToExts.get(uti) ?? new Set();
        set.add(ext);
        utiToExts.set(uti, set);
      }
    }
  }

  /** @type {Set<string>} */
  const viaUtis = new Set();
  for (const uti of docUtis) {
    const exts = utiToExts.get(uti);
    if (!exts) continue;
    for (const ext of exts) viaUtis.add(ext);
  }

  /** @type {Set<string>} */
  const registered = new Set([...docExts, ...viaUtis]);
  return { docTypesBlock, registered, docExts, docUtis, viaUtis };
}

/**
 * @param {unknown} pluginConfig
 * @returns {string[]}
 */
function extractDeepLinkSchemes(pluginConfig) {
  if (!pluginConfig || typeof pluginConfig !== "object") return [];

  const desktop = /** @type {any} */ (pluginConfig).desktop;
  if (!desktop) return [];

  // plugins.deep-link.desktop is a `DeepLinkProtocol` or an array of `DeepLinkProtocol`.
  if (Array.isArray(desktop)) {
    return desktop
      .flatMap((p) => {
        const raw = p?.schemes;
        if (typeof raw === "string") return [raw];
        if (Array.isArray(raw)) return raw;
        return [];
      })
      .filter(Boolean);
  }
  if (typeof desktop === "object") {
    const raw = desktop.schemes;
    if (typeof raw === "string") return [raw].filter(Boolean);
    if (Array.isArray(raw)) return raw.filter(Boolean);
    return [];
  }
  return [];
}

/**
 * @param {any} config
 */
function isParquetAssociationConfigured(config) {
  const fileAssociations = Array.isArray(config?.bundle?.fileAssociations) ? config.bundle.fileAssociations : [];
  return fileAssociations.some((assoc) => {
    const rawExt = /** @type {any} */ (assoc)?.ext;
    const exts = Array.isArray(rawExt) ? rawExt : typeof rawExt === "string" ? [rawExt] : [];
    return exts.some((e) => String(e).trim().toLowerCase().replace(/^\./, "") === "parquet");
  });
}

/**
 * @param {any} config
 * @returns {string}
 */
function getTauriIdentifier(config) {
  const raw = config?.identifier;
  if (typeof raw === "string" && raw.trim()) return raw.trim();
  return "";
}

/**
 * Parquet is not consistently defined in distros' shared-mime-info DB.
 *
 * If we advertise Parquet (`application/vnd.apache.parquet`) in `.desktop` MimeType=
 * we should also ship a shared-mime-info definition file inside Linux bundles so
 * `*.parquet` resolves to that MIME type after install (via update-mime-database).
 *
 * @param {any} config
 */
function validateParquetMimeDefinition(config) {
  if (!isParquetAssociationConfigured(config)) return;

  const identifier = getTauriIdentifier(config);
  if (!identifier) {
    errBlock("Parquet file association configured, but identifier is missing (tauri.conf.json)", [
      "Expected a non-empty top-level `identifier` field in apps/desktop/src-tauri/tauri.conf.json.",
      "This is used as the shared-mime-info XML filename under:",
      "  - /usr/share/mime/packages/<identifier>.xml",
    ]);
    return;
  }

  const linux = config?.bundle?.linux;
  if (!linux || typeof linux !== "object") {
    errBlock("Parquet file association configured, but bundle.linux is missing (tauri.conf.json)", [
      "Expected bundle.linux.{deb,rpm,appimage} to be configured so we can ship a shared-mime-info definition for Parquet.",
    ]);
    return;
  }

  const expectedDest = `usr/share/mime/packages/${identifier}.xml`;
  const expectedSrc = `mime/${identifier}.xml`;

  for (const target of ["deb", "rpm", "appimage"]) {
    const files = linux?.[target]?.files;
    if (!files || typeof files !== "object") {
      errBlock("Parquet file association configured, but Linux bundle file mappings are missing", [
        `Expected bundle.linux.${target}.files to map:`,
        `  - ${expectedDest} -> ${expectedSrc}`,
      ]);
      continue;
    }
    if (files[expectedDest] !== expectedSrc) {
      errBlock("Parquet shared-mime-info mapping mismatch (tauri.conf.json)", [
        `Expected bundle.linux.${target}.files["${expectedDest}"] = "${expectedSrc}"`,
        `Found: ${JSON.stringify(files[expectedDest] ?? null)}`,
      ]);
    }
  }

  // Ensure update-mime-database triggers exist at install time.
  const debDepends = linux?.deb?.depends;
  if (!Array.isArray(debDepends) || !debDepends.includes("shared-mime-info")) {
    errBlock("Parquet file association configured, but shared-mime-info is not declared as a DEB dependency", [
      'Expected bundle.linux.deb.depends to include "shared-mime-info" so update-mime-database triggers run.',
    ]);
  }
  const rpmDepends = linux?.rpm?.depends;
  if (!Array.isArray(rpmDepends) || !rpmDepends.includes("shared-mime-info")) {
    errBlock("Parquet file association configured, but shared-mime-info is not declared as an RPM dependency", [
      'Expected bundle.linux.rpm.depends to include "shared-mime-info" so update-mime-database triggers run.',
    ]);
  }

  const parquetMimeDefinitionPath = path.join(path.dirname(tauriConfigPath), "mime", `${identifier}.xml`);
  try {
    const xml = readFileSync(parquetMimeDefinitionPath, "utf8");
    if (!xml.includes('mime-type type="application/vnd.apache.parquet"') || !xml.includes('glob pattern="*.parquet"')) {
      errBlock("Parquet shared-mime-info definition file is missing expected content", [
        `File: ${parquetMimeDefinitionPath}`,
        'Expected to find: mime-type type="application/vnd.apache.parquet"',
        'Expected to find: glob pattern="*.parquet"',
      ]);
    }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    errBlock("Parquet file association configured, but shared-mime-info definition file is missing", [
      `Expected file to exist: ${parquetMimeDefinitionPath}`,
      `Error: ${msg}`,
    ]);
  }
}

/**
 * @param {any} config
 * @returns {string[]}
 */
function collectFileAssociationExtensions(config) {
  const fileAssociations = Array.isArray(config?.bundle?.fileAssociations) ? config.bundle.fileAssociations : [];
  /** @type {Set<string>} */
  const out = new Set();
  for (const assoc of fileAssociations) {
    const rawExt = /** @type {any} */ (assoc)?.ext;
    const exts = Array.isArray(rawExt) ? rawExt : typeof rawExt === "string" ? [rawExt] : [];
    for (const e of exts) {
      const normalized = String(e).trim().toLowerCase().replace(/^\./, "");
      if (normalized) out.add(normalized);
    }
  }
  return Array.from(out).sort();
}

/**
 * @param {any[]} fileAssociations
 * @returns {Map<string, Set<string>>}
 */
function collectMimeTypesByExtension(fileAssociations) {
  /** @type {Map<string, Set<string>>} */
  const out = new Map();
  for (const assoc of fileAssociations) {
    const rawExt = /** @type {any} */ (assoc)?.ext;
    const exts = Array.isArray(rawExt) ? rawExt : typeof rawExt === "string" ? [rawExt] : [];
    const rawMime = /** @type {any} */ (assoc)?.mimeType;
    const mime = typeof rawMime === "string" ? rawMime.trim().toLowerCase() : "";
    if (!mime) continue;
    for (const e of exts) {
      const normalized = String(e).trim().toLowerCase().replace(/^\./, "");
      if (!normalized) continue;
      const existing = out.get(normalized) ?? new Set();
      existing.add(mime);
      out.set(normalized, existing);
    }
  }
  return out;
}

/**
 * Ensure `bundle.fileAssociations` is structurally sane: every entry has an `ext` list,
 * and extensions are unique across entries (so platform bundlers don't silently ignore
 * duplicates or behave inconsistently).
 *
 * @param {any[]} fileAssociations
 */
function validateFileAssociationsShape(fileAssociations) {
  /** @type {Set<string>} */
  const seen = new Set();
  /** @type {Set<string>} */
  const duplicates = new Set();
  /** @type {number[]} */
  const missingExtIndices = [];

  for (let i = 0; i < fileAssociations.length; i += 1) {
    const assoc = fileAssociations[i];
    const rawExt = /** @type {any} */ (assoc)?.ext;
    const exts = Array.isArray(rawExt) ? rawExt : typeof rawExt === "string" ? [rawExt] : [];
    if (exts.length === 0) {
      missingExtIndices.push(i);
      continue;
    }
    for (const e of exts) {
      const normalized = String(e).trim().toLowerCase().replace(/^\./, "");
      if (!normalized) continue;
      if (seen.has(normalized)) {
        duplicates.add(normalized);
      } else {
        seen.add(normalized);
      }
    }
  }

  if (missingExtIndices.length > 0) {
    errBlock("Invalid file association entry (tauri.conf.json)", [
      "Expected every bundle.fileAssociations[] entry to include a non-empty `ext` list.",
      `Missing ext in fileAssociations indices: ${missingExtIndices.join(", ")}`,
      "Fix: add ext: [\"...\"] for each file association entry.",
    ]);
  }
  if (duplicates.size > 0) {
    errBlock("Duplicate file association extensions (tauri.conf.json)", [
      `Duplicate extension(s) found: ${Array.from(duplicates).sort().join(", ")}`,
      "Fix: ensure each extension appears in exactly one bundle.fileAssociations entry.",
    ]);
  }
}

function main() {
  /** @type {any} */
  let config;
  try {
    config = JSON.parse(readFileSync(tauriConfigPath, "utf8"));
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    errBlock("Desktop URL scheme preflight failed", [
      "Failed to read/parse apps/desktop/src-tauri/tauri.conf.json.",
      `Error: ${msg}`,
    ]);
    return;
  }
 
  // ---- macOS: Info.plist contains CFBundleURLSchemes -> formula, and CFBundleDocumentTypes includes xlsx.
  try {
    const plist = readFileSync(infoPlistPath, "utf8");
    const schemeBlocks = extractAllPlistArrayBlocks(plist, "CFBundleURLSchemes");
    const deepLinkConfig = config?.plugins?.["deep-link"];
    const expectedSchemes = new Set(
      [REQUIRED_SCHEME, ...extractDeepLinkSchemes(deepLinkConfig)]
        .map((s) => String(s).trim().replace(/[:/]+$/, "").toLowerCase())
        .filter(Boolean),
    );
    const observedSchemes = new Set(
      schemeBlocks
        .flatMap((block) => extractPlistStrings(block))
        .map((s) => String(s).trim().replace(/[:/]+$/, "").toLowerCase())
        .filter(Boolean),
    );
    const missingSchemes = Array.from(expectedSchemes).filter((s) => !observedSchemes.has(s));
    if (schemeBlocks.length === 0 || missingSchemes.length > 0) {
      errBlock("Missing macOS URL scheme registration (Info.plist)", [
        "Expected apps/desktop/src-tauri/Info.plist to declare CFBundleURLSchemes including:",
        ...Array.from(expectedSchemes)
          .sort()
          .map((s) => `  - ${s}`),
        ...(missingSchemes.length > 0 ? [`Missing scheme(s): ${missingSchemes.sort().join(", ")}`] : []),
        `Found scheme(s): ${Array.from(observedSchemes).sort().join(", ") || "(none)"}`,
        "Fix: add/update CFBundleURLTypes/CFBundleURLSchemes in Info.plist.",
      ]);
    }

    const docTypesBlock = extractPlistArrayBlock(plist, "CFBundleDocumentTypes");
    const expectedExts = collectFileAssociationExtensions(config);
    const { registered, docExts, docUtis, viaUtis } = collectMacosRegisteredExtensions(plist);
    const missingExts = expectedExts.filter((ext) => !registered.has(ext));

    if (!docTypesBlock || missingExts.length > 0) {
      errBlock("Missing macOS file association registration (Info.plist)", [
        "Expected apps/desktop/src-tauri/Info.plist to register handlers for all extensions in bundle.fileAssociations (tauri.conf.json).",
        "Acceptable forms:",
        "  - CFBundleDocumentTypes / CFBundleTypeExtensions includes the extension",
        "  - CFBundleDocumentTypes / LSItemContentTypes references a UTI whose UT*TypeDeclarations defines public.filename-extension",
        ...(missingExts.length > 0
          ? [`Missing extension(s): ${missingExts.join(", ")}`, `Expected extensions: ${expectedExts.join(", ")}`]
          : []),
        `Observed CFBundleTypeExtensions: ${Array.from(docExts).join(", ") || "(none)"}`,
        `Observed LSItemContentTypes: ${Array.from(docUtis).join(", ") || "(none)"}`,
        `Observed extensions via LSItemContentTypes: ${Array.from(viaUtis).join(", ") || "(none)"}`,
        "Fix: add/update CFBundleDocumentTypes/CFBundleTypeExtensions in Info.plist.",
      ]);
    }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    errBlock("Desktop URL scheme preflight failed", [
      "Failed to read apps/desktop/src-tauri/Info.plist.",
      `Error: ${msg}`,
    ]);
  }

  // ---- All platforms: bundler must know about formula:// so installers register it.
  const deepLinkConfig = config?.plugins?.["deep-link"];
  const schemes = extractDeepLinkSchemes(deepLinkConfig);
  const normalizedSchemes = schemes
    .map((s) => String(s).trim().replace(/[:/]+$/, "").toLowerCase())
    .filter(Boolean);

  if (!normalizedSchemes.includes(REQUIRED_SCHEME)) {
    const found = normalizedSchemes.length > 0 ? normalizedSchemes.join(", ") : "(none)";
    errBlock("Missing desktop deep-link scheme configuration (tauri.conf.json)", [
      "Expected apps/desktop/src-tauri/tauri.conf.json to include:",
      `  plugins[\"deep-link\"].desktop.schemes = [\"${REQUIRED_SCHEME}\"]`,
      `Found schemes: ${found}`,
      "Fix: add the deep-link plugin config so installers register the URL scheme.",
    ]);
  }

  // ---- Linux: freedesktop .desktop generation only includes explicit file association mimeType values.
  const fileAssociations = config?.bundle?.fileAssociations;
  if (!Array.isArray(fileAssociations) || fileAssociations.length === 0) {
    errBlock("Missing bundle.fileAssociations (tauri.conf.json)", [
      "Expected apps/desktop/src-tauri/tauri.conf.json to include bundle.fileAssociations so Linux .desktop metadata can include file MIME types.",
    ]);
  } else {
    validateFileAssociationsShape(fileAssociations);
    const configuredExts = new Set(collectFileAssociationExtensions(config));
    const missingRequiredExts = REQUIRED_FILE_EXTS.filter((ext) => !configuredExts.has(ext));

    if (missingRequiredExts.length > 0) {
      errBlock("Missing required desktop file associations (tauri.conf.json)", [
        "Expected apps/desktop/src-tauri/tauri.conf.json bundle.fileAssociations to include entries for:",
        ...REQUIRED_FILE_EXTS.map((ext) => `  - .${ext}`),
        `Missing extension(s): ${missingRequiredExts.join(", ")}`,
        "Fix: add/update bundle.fileAssociations so these types are registered with the OS.",
      ]);
    }

    const missingMime = fileAssociations
      .map((assoc, i) => ({ assoc, i }))
      .filter(({ assoc }) => typeof assoc?.mimeType !== "string" || assoc.mimeType.trim().length === 0);

    if (missingMime.length > 0) {
      errBlock("Missing Linux mimeType fields in bundle.fileAssociations", [
        "Tauri's Linux .desktop file generation only includes file MIME types when bundle.fileAssociations[].mimeType is set.",
        ...missingMime.map(({ assoc, i }) => {
          const exts = Array.isArray(assoc?.ext) ? assoc.ext.join(", ") : "(missing ext)";
          return `fileAssociations[${i}] ext=[${exts}] is missing mimeType`;
        }),
        "Fix: add mimeType to each file association entry (Linux-only field).",
      ]);
    }

    // Ensure the most important extensions are wired to stable MIME types so OS integration works.
    // If these drift, the `.desktop` (Linux) and file manager resolution may no longer line up.
    const expectedMimeByExt = {
      xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      xls: "application/vnd.ms-excel",
      xlt: "application/vnd.ms-excel",
      xla: "application/vnd.ms-excel",
      xlsm: "application/vnd.ms-excel.sheet.macroenabled.12",
      xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
      xltm: "application/vnd.ms-excel.template.macroenabled.12",
      xlam: "application/vnd.ms-excel.addin.macroenabled.12",
      xlsb: "application/vnd.ms-excel.sheet.binary.macroenabled.12",
      csv: "text/csv",
      parquet: "application/vnd.apache.parquet",
    };
    const mimeByExt = collectMimeTypesByExtension(fileAssociations);
    for (const [ext, expectedMime] of Object.entries(expectedMimeByExt)) {
      const observed = mimeByExt.get(ext);
      if (!observed || observed.size === 0) continue;
      if (!observed.has(expectedMime)) {
        errBlock("File association mimeType mismatch (tauri.conf.json)", [
          `Extension .${ext} must use mimeType: ${expectedMime}`,
          `Found mimeType(s): ${Array.from(observed).join(", ")}`,
          "Fix: update bundle.fileAssociations so the configured MIME type matches OS conventions.",
        ]);
      } else if (observed.size > 1) {
        errBlock("File association mimeType is ambiguous (tauri.conf.json)", [
          `Extension .${ext} should map to a single mimeType (${expectedMime}).`,
          `Found multiple mimeType values: ${Array.from(observed).join(", ")}`,
          "Fix: remove duplicate/overlapping file association entries so each extension has one MIME type.",
        ]);
      }
    }
  }

  validateParquetMimeDefinition(config);

  if (process.exitCode) {
    err(
      "\nDesktop URL scheme + file association preflight failed. Fix the errors above before tagging a release.\n",
    );
    return;
  }

  console.log(
    `Desktop URL scheme + file association preflight passed: ${REQUIRED_SCHEME}:// is configured for installers (and Info.plist declares it), and bundle.fileAssociations includes required extensions (${REQUIRED_FILE_EXTS.join(", ")}).`,
  );
}

main();
