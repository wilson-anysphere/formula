import { readFile } from "node:fs/promises";

import ts from "typescript";

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

export async function resolve(specifier, context, defaultResolve) {
  try {
    return await defaultResolve(specifier, context, defaultResolve);
  } catch (err) {
    if (!isPathLike(specifier)) throw err;
    // Only fall back for missing modules; other resolution errors (like invalid exports)
    // should be surfaced to the caller.
    const code = /** @type {any} */ (err)?.code;
    if (code !== "ERR_MODULE_NOT_FOUND") throw err;

    const { base, suffix } = splitSpecifier(specifier);

    // `./foo.js` -> `./foo.ts` fallback (TypeScript ESM convention).
    if (base.endsWith(".js")) {
      try {
        return await defaultResolve(base.slice(0, -3) + ".ts" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        const candidateCode = /** @type {any} */ (candidateErr)?.code;
        if (candidateCode !== "ERR_MODULE_NOT_FOUND") throw candidateErr;
      }
      try {
        return await defaultResolve(base.slice(0, -3) + ".tsx" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        const candidateCode = /** @type {any} */ (candidateErr)?.code;
        if (candidateCode !== "ERR_MODULE_NOT_FOUND") throw candidateErr;
      }
    }

    // `./foo.jsx` -> `./foo.tsx` fallback.
    if (base.endsWith(".jsx")) {
      try {
        return await defaultResolve(base.slice(0, -4) + ".tsx" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        const candidateCode = /** @type {any} */ (candidateErr)?.code;
        if (candidateCode !== "ERR_MODULE_NOT_FOUND") throw candidateErr;
      }
    }

    // Extensionless `./foo` -> try TS/JS.
    if (!hasExtension(base)) {
      try {
        return await defaultResolve(base + ".ts" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        const candidateCode = /** @type {any} */ (candidateErr)?.code;
        if (candidateCode !== "ERR_MODULE_NOT_FOUND") throw candidateErr;
      }
      try {
        return await defaultResolve(base + ".tsx" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        const candidateCode = /** @type {any} */ (candidateErr)?.code;
        if (candidateCode !== "ERR_MODULE_NOT_FOUND") throw candidateErr;
      }
      try {
        return await defaultResolve(base + ".js" + suffix, context, defaultResolve);
      } catch (candidateErr) {
        const candidateCode = /** @type {any} */ (candidateErr)?.code;
        if (candidateCode !== "ERR_MODULE_NOT_FOUND") throw candidateErr;
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
    urlObj.search = "";
    urlObj.hash = "";
    const source = await readFile(urlObj, "utf8");
    const isTsx = pathname.endsWith(".tsx");
    const result = ts.transpileModule(source, {
      fileName: pathname,
      compilerOptions: {
        module: ts.ModuleKind.ESNext,
        target: ts.ScriptTarget.ES2022,
        ...(isTsx ? { jsx: ts.JsxEmit.ReactJSX } : {}),
      },
    });

    return {
      format: "module",
      source: result.outputText,
      shortCircuit: true,
    };
  }

  return defaultLoad(url, context, defaultLoad);
}
