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
 * @param {string} specifier
 * @param {any} context
 * @param {(specifier: string, context: any, nextResolve: any) => Promise<any>} defaultResolve
 */
export async function resolve(specifier, context, defaultResolve) {
  if (typeof specifier !== "string") return defaultResolve(specifier, context, defaultResolve);

  const { base, suffix } = splitSpecifier(specifier);
  const isRelativeOrAbsolute =
    base.startsWith("./") ||
    base.startsWith("../") ||
    base.startsWith("/") ||
    base.startsWith("file:");

  if (!isRelativeOrAbsolute) {
    return defaultResolve(specifier, context, defaultResolve);
  }

  const isJs = base.endsWith(".js");
  const isJsx = base.endsWith(".jsx");
  const isExtensionless = !hasExtension(base);
  if (!isJs && !isJsx && !isExtensionless) {
    return defaultResolve(specifier, context, defaultResolve);
  }

  try {
    return await defaultResolve(specifier, context, defaultResolve);
  } catch (err) {
    if (!isResolutionMiss(err)) throw err;

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
