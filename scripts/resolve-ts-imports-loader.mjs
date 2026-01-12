/**
 * Node ESM loader used by `scripts/run-node-tests.mjs`.
 *
 * Purpose
 * -------
 * Many TS sources in this repo follow the "TS source imports .js specifiers" pattern:
 *   import { foo } from "./foo.js";
 * while the source file on disk is `foo.ts` (or `foo.tsx`).
 *
 * Bundlers (and TypeScript's `moduleResolution: "Bundler"`) understand this, but
 * Node's default ESM resolver does not. When we run `.test.js` files directly via
 * `node --test` (with `--experimental-strip-types`), those imports would fail with
 * `ERR_MODULE_NOT_FOUND`.
 *
 * This loader keeps runtime semantics identical when a real `.js` file exists, but
 * falls back to the matching `.ts`/`.tsx` source when it doesn't.
 *
 * Notes
 * -----
 * - We rely on Node's `--experimental-strip-types` to execute the `.ts`/`.tsx` files.
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
  if (!isJs && !isJsx) {
    return defaultResolve(specifier, context, defaultResolve);
  }

  try {
    return await defaultResolve(specifier, context, defaultResolve);
  } catch (err) {
    const code = /** @type {any} */ (err)?.code;
    if (code !== "ERR_MODULE_NOT_FOUND") throw err;

    const baseNoExt = isJs ? base.slice(0, -3) : base.slice(0, -4);
    const candidates = isJs ? [".ts", ".tsx"] : [".tsx", ".ts"];
    for (const ext of candidates) {
      try {
        return await defaultResolve(`${baseNoExt}${ext}${suffix}`, context, defaultResolve);
      } catch (candidateErr) {
        const candidateCode = /** @type {any} */ (candidateErr)?.code;
        if (candidateCode !== "ERR_MODULE_NOT_FOUND") throw candidateErr;
      }
    }

    throw err;
  }
}

