import { readFile } from "node:fs/promises";

import ts from "typescript";

const shouldEmitSourceMaps =
  process.execArgv.includes("--enable-source-maps") || (process.env.NODE_OPTIONS?.includes("--enable-source-maps") ?? false);

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
 * - Resolves extensionless relative/absolute specifiers to `.ts` / `.tsx` / `.js`
 * - Resolves directory imports like `./foo` -> `./foo/index.ts` (bundler-style resolution)
 * - Transpiles `.ts` / `.tsx` sources on the fly (strip-only mode is insufficient for
 *   TypeScript features like parameter properties).
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

export async function resolve(specifier, context, defaultResolve) {
  try {
    return await defaultResolve(specifier, context, defaultResolve);
  } catch (err) {
    if (!isPathLike(specifier)) throw err;
    // Only fall back for missing modules; other resolution errors (like invalid exports)
    // should be surfaced to the caller.
    if (!isResolutionMiss(err)) throw err;

    const { base, suffix } = splitSpecifier(specifier);

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
      promise.finally(() => {
        transpileInFlight.delete(url);
      });
    }

    const loaded = await promise;
    return { ...loaded, shortCircuit: true };
  }

  return defaultLoad(url, context, defaultLoad);
}
