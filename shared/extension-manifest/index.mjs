const GLOBAL_KEY = "__formula_extension_manifest__";

// Load the implementation (CommonJS in Node, plain module in browsers).
import "./index.js";

const impl = globalThis?.[GLOBAL_KEY];
if (!impl) {
  throw new Error("shared/extension-manifest: failed to initialize implementation");
}

export const ManifestError = impl.ManifestError;
export const VALID_PERMISSIONS = impl.VALID_PERMISSIONS;
export const validateExtensionManifest = impl.validateExtensionManifest;

export default impl;
