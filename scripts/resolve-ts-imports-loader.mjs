/**
 * Node ESM loader used by `scripts/run-node-tests.mjs`.
 *
 * Purpose
 * -------
 * Many TS sources in this repo follow the "TS source imports .js specifiers" pattern:
 *   import { foo } from "./foo.js";
 * while the source file on disk is `foo.ts`.
 *
 * Some sources also import `.jsx` specifiers (even when the file on disk is still `.ts`).
 *
 * Bundlers (and TypeScript's `moduleResolution: "Bundler"`) understand this, but
 * Node's default ESM resolver does not. When we run `.test.js` files directly via
 * `node --test` (executing TypeScript sources directly), those imports would fail
 * with `ERR_MODULE_NOT_FOUND`.
 *
 * This loader keeps runtime semantics identical when a real `.js`/`.jsx` file exists, but
 * falls back to the matching `.ts` source when it doesn't.
 *
 * Some TS sources also use `.jsx` specifiers (bundler-style) that should resolve to
 * `.ts` sources. We support that too (but do **not** resolve to `.tsx`, since Node's
 * built-in TypeScript execution cannot load `.tsx`).
 *
 * Additionally, some TS sources in this repo still use extensionless relative
 * imports (e.g. `import "./foo"`). Node ESM does not support extensionless path
 * resolution, so when the specifier is missing (and the file exists as `foo.ts`)
 * we also provide a `.ts` fallback. This includes directory imports (e.g.
 * `import "./foo"` where `foo/index.ts` exists).
 *
 * Notes
 * -----
 * - This loader is only used when `scripts/run-node-tests.mjs` is running via Node's
 *   built-in TypeScript "strip types" support. That mode only supports `.ts` files
 *   (Node does not execute `.tsx`/JSX without an additional transpile loader).
 * - The loader is intentionally dependency-free (no `typescript` package) so it can
 *   run in constrained environments.
 */

import { readFile, readdir, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/**
 * Split a module specifier into `[base, suffix]`, where suffix includes any query or
 * hash fragment (e.g. `?raw`, `#foo`).
 *
 * @param {string} specifier
 * @returns {{ base: string, suffix: string }}
 */
function splitSpecifier(specifier) {
  const queryIdx = specifier.indexOf("?");
  const hashIdx = specifier.indexOf("#");
  const idx =
    queryIdx === -1 ? hashIdx : hashIdx === -1 ? queryIdx : Math.min(queryIdx, hashIdx);
  if (idx === -1) return { base: specifier, suffix: "" };
  return { base: specifier.slice(0, idx), suffix: specifier.slice(idx) };
}

/**
 * @param {string} pathLike
 */
function hasExtension(pathLike) {
  const lastSlash = Math.max(pathLike.lastIndexOf("/"), pathLike.lastIndexOf("\\"));
  const lastDot = pathLike.lastIndexOf(".");
  return lastDot > lastSlash;
}

/**
 * @param {unknown} error
 */
function isResolutionMiss(error) {
  const code = /** @type {any} */ (error)?.code;
  return code === "ERR_MODULE_NOT_FOUND" || code === "ERR_UNSUPPORTED_DIR_IMPORT";
}

/**
 * Lazily-loaded map of workspace package name -> basic package.json info.
 *
 * This loader intentionally has no dependency on the `typescript` package (it is used when
 * relying on Node's built-in "strip types" support), so we implement a lightweight
 * workspace resolver for stale/partial `node_modules` trees that are missing some
 * workspace links.
 *
 * @type {Promise<Map<string, { rootDir: string, exports: any, main: string | null }>> | null}
 */
let workspacePackagesPromise = null;

async function loadWorkspacePackages() {
  if (workspacePackagesPromise) return workspacePackagesPromise;
  workspacePackagesPromise = (async () => {
    /** @type {string[]} */
    const packageJsonFiles = [];
    for (const dir of ["packages", "apps", "services", "shared"]) {
      const full = path.join(repoRoot, dir);
      try {
        const st = await stat(full);
        if (!st.isDirectory()) continue;
      } catch {
        continue;
      }
      await collectPackageJsonFiles(full, packageJsonFiles);
    }

    /** @type {Map<string, { rootDir: string, exports: any, main: string | null }>} */
    const out = new Map();
    for (const file of packageJsonFiles) {
      const raw = await readFile(file, "utf8").catch(() => null);
      if (!raw) continue;
      let parsed;
      try {
        parsed = JSON.parse(raw);
      } catch {
        continue;
      }
      if (!parsed || typeof parsed !== "object") continue;
      if (typeof parsed.name !== "string") continue;
      if (out.has(parsed.name)) continue;
      out.set(parsed.name, {
        rootDir: path.dirname(file),
        exports: parsed.exports ?? null,
        main: typeof parsed.main === "string" ? parsed.main : null,
      });
    }

    return out;
  })();

  return workspacePackagesPromise;
}

async function collectPackageJsonFiles(dir, out) {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    // Skip node_modules and other generated output.
    if (
      entry.name === ".git" ||
      entry.name === "node_modules" ||
      entry.name === "dist" ||
      entry.name === "coverage" ||
      entry.name === "target" ||
      entry.name === "build" ||
      entry.name === ".turbo" ||
      entry.name === ".pnpm-store" ||
      entry.name === ".cache" ||
      entry.name === ".vite" ||
      entry.name === "playwright-report" ||
      entry.name === "test-results" ||
      entry.name === "security-report" ||
      entry.name.startsWith(".tmp")
    ) {
      continue;
    }

    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectPackageJsonFiles(fullPath, out);
      continue;
    }

    if (!entry.isFile()) continue;
    if (entry.name !== "package.json") continue;
    out.push(fullPath);
  }
}

function parseWorkspaceSpecifier(specifier) {
  if (!specifier.startsWith("@formula/")) return null;
  const parts = specifier.split("/");
  if (parts.length < 2) return null;
  const packageName = `${parts[0]}/${parts[1]}`;
  const subpath = parts.slice(2).join("/");
  const exportKey = subpath ? `./${subpath}` : ".";
  return { packageName, exportKey };
}

function resolveWorkspaceExport(exportsMap, exportKey, main) {
  /**
   * @param {any} entry
   * @returns {string | null}
   */
  function pickExportPath(entry) {
    if (typeof entry === "string") return entry;
    if (Array.isArray(entry)) {
      for (const item of entry) {
        const picked = pickExportPath(item);
        if (picked) return picked;
      }
      return null;
    }
    if (entry && typeof entry === "object") {
      // Prefer Node's default ESM condition order.
      if (typeof entry.node === "string") return entry.node;
      if (typeof entry.import === "string") return entry.import;
      if (typeof entry.default === "string") return entry.default;
      if (typeof entry.require === "string") return entry.require;

      // Fall back to searching nested condition objects.
      for (const value of Object.values(entry)) {
        const picked = pickExportPath(value);
        if (picked) return picked;
      }
    }
    return null;
  }

  let target = null;
  if (exportsMap) {
    if (typeof exportsMap === "string") {
      if (exportKey === ".") target = exportsMap;
    } else if (exportsMap && typeof exportsMap === "object") {
      const keys = Object.keys(exportsMap);
      const looksLikeSubpathMap = keys.some((k) => k === "." || k.startsWith("./"));
      if (looksLikeSubpathMap) {
        target = pickExportPath(exportsMap?.[exportKey]);
      } else if (exportKey === ".") {
        target = pickExportPath(exportsMap);
      }
    }
  }

  // Packages without an `exports` map (like `@formula/marketplace-shared`) still allow deep
  // imports via `@scope/pkg/<path>`. When a workspace link is missing from `node_modules`,
  // fall back to resolving those subpaths directly from the package root directory.
  if (!exportsMap && exportKey !== ".") return exportKey;

  if (!target && exportKey === "." && typeof main === "string") target = main;
  return target;
}

/**
 * @param {string} specifier
 * @param {any} context
 * @param {(specifier: string, context: any, nextResolve: any) => Promise<any>} defaultResolve
 */
export async function resolve(specifier, context, defaultResolve) {
  if (typeof specifier !== "string") return defaultResolve(specifier, context, defaultResolve);

  try {
    return await defaultResolve(specifier, context, defaultResolve);
  } catch (err) {
    if (!isResolutionMiss(err)) throw err;

    const { base, suffix } = splitSpecifier(specifier);

    // Workspace package fallback for stale/minimal `node_modules` installs:
    // `@formula/foo` -> `<repoRoot>/packages/...` entrypoint.
    if (base.startsWith("@formula/")) {
      const parsed = parseWorkspaceSpecifier(base);
      if (parsed) {
        const pkgs = await loadWorkspacePackages();
        const info = pkgs.get(parsed.packageName);
        if (info) {
          const target = resolveWorkspaceExport(info.exports, parsed.exportKey, info.main);
          if (typeof target === "string" && target.trim() !== "") {
            const cleaned =
              target.startsWith("./") || target.startsWith("../") || target.startsWith("/")
                ? target
                : `./${target}`;
            const absPath = path.resolve(info.rootDir, cleaned);
            const url = pathToFileURL(absPath).href + suffix;
            return { url, shortCircuit: true };
          }
        }
      }
    }

    const isRelativeOrAbsolute =
      base.startsWith("./") || base.startsWith("../") || base.startsWith("/") || base.startsWith("file:");
    if (!isRelativeOrAbsolute) throw err;

    const isJs = base.endsWith(".js");
    const isJsx = base.endsWith(".jsx");
    const isExtensionless = !hasExtension(base);
    if (!isJs && !isJsx && !isExtensionless) throw err;

    if (isJs || isJsx) {
      const baseNoExt = isJs ? base.slice(0, -3) : base.slice(0, -4);
      try {
        return await defaultResolve(`${baseNoExt}.ts${suffix}`, context, defaultResolve);
      } catch (candidateErr) {
        if (!isResolutionMiss(candidateErr)) throw candidateErr;
      }
    }

    if (isExtensionless) {
      for (const ext of [".ts", ".js"]) {
        try {
          return await defaultResolve(`${base}${ext}${suffix}`, context, defaultResolve);
        } catch (candidateErr) {
          if (!isResolutionMiss(candidateErr)) throw candidateErr;
        }
      }

      // Directory imports: `./foo` -> `./foo/index.ts` (bundler-style resolution).
      for (const idx of ["index.ts", "index.js"]) {
        try {
          return await defaultResolve(`${base}/${idx}${suffix}`, context, defaultResolve);
        } catch (candidateErr) {
          if (!isResolutionMiss(candidateErr)) throw candidateErr;
        }
      }
    }

    throw err;
  }
}

export async function load(url, context, defaultLoad) {
  const urlObj = new URL(url);
  // Support Vite-style `?raw` imports when running node:test suites directly against
  // workspace TypeScript sources.
  if (urlObj.protocol === "file:" && urlObj.searchParams.has("raw")) {
    urlObj.search = "";
    urlObj.hash = "";
    const source = await readFile(urlObj, "utf8");
    return { format: "module", source: `export default ${JSON.stringify(source)};`, shortCircuit: true };
  }
  return defaultLoad(url, context, defaultLoad);
}
