const GLOBAL_KEY = "__formula_semver_range__";

// Load the implementation (CommonJS in Node, plain module in browsers).
import "./index.js";

const impl = globalThis?.[GLOBAL_KEY];
if (!impl) {
  throw new Error("shared/semver-range: failed to initialize implementation");
}

export const compareSemver = impl.compareSemver;
export const isValidSemver = impl.isValidSemver;
export const maxSemver = impl.maxSemver;
export const parseSemver = impl.parseSemver;
export const satisfies = impl.satisfies;

export default impl;
