import { readFile, readdir, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import ts from "typescript";

const shouldEmitSourceMaps =
  process.execArgv.includes("--enable-source-maps") || (process.env.NODE_OPTIONS?.includes("--enable-source-maps") ?? false);

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

/**
 * In-flight transpile de-dupe.
 *
 * Node's ESM loader caches modules after they've been loaded, so `load()` is usually
 * only called once per URL. However, Node may call loader hooks concurrently during
 * graph construction; this prevents redundant TypeScript transpiles without keeping
 * another long-lived copy of every module's JS in memory.
 *
 * @type {Map<string, Promise<{ format: "module", source: string }>>}
 */
const transpileInFlight = new Map();

/**
 * Minimal Node ESM loader that lets us run repo TypeScript sources directly under `node --test`.
 *
 * Motivation:
 * - Many workspace packages are authored as `.ts` but use `.js` specifiers (TypeScript's recommended
 *   ESM pattern when compiling TS -> JS).
 * - Node does not resolve `./foo.js` to `./foo.ts` by default.
 *
 * This loader:
 * - Resolves missing relative/absolute `.js` specifiers to `.ts` / `.tsx`
 * - Resolves missing relative/absolute `.jsx` specifiers to `.tsx` / `.ts`
 * - Resolves extensionless relative/absolute specifiers to `.ts` / `.tsx` / `.js`
 * - Resolves directory imports like `./foo` -> `./foo/index.ts` (bundler-style resolution)
 * - Transpiles `.ts` / `.tsx` sources on the fly (strip-only mode is insufficient for
 *   TypeScript runtime features like parameter properties and enums).
 */

function splitSpecifier(specifier) {
  const q = specifier.indexOf("?");
  const h = specifier.indexOf("#");
  const idx = q === -1 ? h : h === -1 ? q : Math.min(q, h);
  if (idx === -1) return { base: specifier, suffix: "" };
  return { base: specifier.slice(0, idx), suffix: specifier.slice(idx) };
}

function isPathLike(specifier) {
  return (
    specifier.startsWith("./") ||
    specifier.startsWith("../") ||
    specifier.startsWith("/") ||
    specifier.startsWith("file:")
  );
}

function hasExtension(pathLike) {
  const lastSlash = Math.max(pathLike.lastIndexOf("/"), pathLike.lastIndexOf("\\"));
  const lastDot = pathLike.lastIndexOf(".");
  return lastDot > lastSlash;
}

function isResolutionMiss(error) {
  const code = /** @type {any} */ (error)?.code;
  return code === "ERR_MODULE_NOT_FOUND" || code === "ERR_UNSUPPORTED_DIR_IMPORT";
}

/**
 * Lazily-loaded map of workspace package name -> basic package.json info.
 *
 * We use this as a fallback for monorepo environments where `node_modules/` exists but is
 * missing some workspace links (stale or partial installs). When a `@formula/*` package
 * isn't resolvable via Node's default algorithm, we point it directly at its on-disk
 * entrypoint (often a `.ts` source file).
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

  if (!target && exportKey === "." && typeof main === "string") target = main;
  return target;
}

export async function resolve(specifier, context, defaultResolve) {
  try {
    return await defaultResolve(specifier, context, defaultResolve);
  } catch (err) {
    const { base, suffix } = splitSpecifier(specifier);

    // Workspace package fallback for stale/minimal `node_modules` installs:
    // `@formula/foo` -> `<repoRoot>/packages/...` entrypoint.
    if (base.startsWith("@formula/") && isResolutionMiss(err)) {
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

    if (!isPathLike(specifier)) throw err;
    // Only fall back for missing modules; other resolution errors (like invalid exports)
    // should be surfaced to the caller.
    if (!isResolutionMiss(err)) throw err;

    // `./foo.js` -> `./foo.ts` fallback (TypeScript ESM convention).
    if (base.endsWith(".js")) {
      try {
        return await defaultResolve(base.slice(0, -3) + ".ts" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        if (!isResolutionMiss(candidateErr)) throw candidateErr;
      }
      try {
        return await defaultResolve(base.slice(0, -3) + ".tsx" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        if (!isResolutionMiss(candidateErr)) throw candidateErr;
      }
    }

    // `./foo.jsx` -> `./foo.tsx` fallback.
    if (base.endsWith(".jsx")) {
      try {
        return await defaultResolve(base.slice(0, -4) + ".tsx" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        if (!isResolutionMiss(candidateErr)) throw candidateErr;
      }
      try {
        return await defaultResolve(base.slice(0, -4) + ".ts" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        if (!isResolutionMiss(candidateErr)) throw candidateErr;
      }
    }

    // Extensionless `./foo` -> try TS/JS.
    if (!hasExtension(base)) {
      try {
        return await defaultResolve(base + ".ts" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        if (!isResolutionMiss(candidateErr)) throw candidateErr;
      }
      try {
        return await defaultResolve(base + ".tsx" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        if (!isResolutionMiss(candidateErr)) throw candidateErr;
      }
      try {
        return await defaultResolve(base + ".js" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        if (!isResolutionMiss(candidateErr)) throw candidateErr;
      }

      // Directory imports: `./foo` -> `./foo/index.ts` (bundler-style resolution).
      for (const idx of ["index.ts", "index.tsx", "index.js"]) {
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
  // `url` may include a query/hash; use the pathname to decide file type, and read
  // the on-disk file URL without the suffix.
  const urlObj = new URL(url);
  const pathname = urlObj.pathname;
  // Vite-style `?raw` imports are used in a few desktop/node contexts (e.g. locale TSVs,
  // bundled extension entrypoints). Node does not understand these by default, so treat
  // them as "read the file and export a string".
  if (urlObj.protocol === "file:" && urlObj.searchParams.has("raw")) {
    urlObj.search = "";
    urlObj.hash = "";
    const source = await readFile(urlObj, "utf8");
    return { format: "module", source: `export default ${JSON.stringify(source)};`, shortCircuit: true };
  }
  if (pathname.endsWith(".ts") || pathname.endsWith(".tsx")) {
    let promise = transpileInFlight.get(url);
    if (!promise) {
      promise = (async () => {
        urlObj.search = "";
        urlObj.hash = "";
        const source = await readFile(urlObj, "utf8");
        const isTsx = pathname.endsWith(".tsx");
        const result = ts.transpileModule(source, {
          fileName: pathname,
          reportDiagnostics: true,
          compilerOptions: {
            module: ts.ModuleKind.ESNext,
            target: ts.ScriptTarget.ES2022,
            // Only emit inline source maps when source map support is enabled; otherwise
            // they add overhead (larger module source strings) with no benefit.
            inlineSourceMap: shouldEmitSourceMaps,
            ...(isTsx ? { jsx: ts.JsxEmit.ReactJSX } : {}),
          },
        });

        const errorDiagnostics = result.diagnostics?.filter((d) => d.category === ts.DiagnosticCategory.Error) ?? [];
        if (errorDiagnostics.length > 0) {
          const details = errorDiagnostics
            .map((d) => {
              const msg = ts.flattenDiagnosticMessageText(d.messageText, "\n");
              const loc =
                d.file && typeof d.start === "number"
                  ? d.file.getLineAndCharacterOfPosition(d.start)
                  : null;
              const suffix = loc ? `:${loc.line + 1}:${loc.character + 1}` : "";
              return `${pathname}${suffix} - TS${d.code}: ${msg}`;
            })
            .join("\n");
          throw new SyntaxError(`Failed to transpile TypeScript module:\n${details}`);
        }

        /** @type {{ format: "module", source: string }} */
        return { format: "module", source: result.outputText };
      })();

      transpileInFlight.set(url, promise);
      // Avoid `.finally()` here: ignoring the returned promise can trigger an unhandled
      // rejection if the underlying transpile fails. Use `then(..., ...)` so the
      // cleanup chain always resolves.
      const cleanup = () => {
        transpileInFlight.delete(url);
      };
      promise.then(cleanup, cleanup);
    }

    const loaded = await promise;
    return { ...loaded, shortCircuit: true };
  }

  return defaultLoad(url, context, defaultLoad);
}
