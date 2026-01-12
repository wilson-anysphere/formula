export type ValidateExtensionManifestOptions = {
  /**
   * The running Formula engine version (semver string).
   *
   * Required when `enforceEngine` is `true`.
   */
  engineVersion?: string;
  /**
   * When `true`, validate that `engines.formula` satisfies `engineVersion`.
   */
  enforceEngine?: boolean;
};

export declare class ManifestError extends Error {}

export declare const VALID_PERMISSIONS: ReadonlySet<string>;

export declare function validateExtensionManifest(
  manifest: Record<string, any>,
  options?: ValidateExtensionManifestOptions
): Record<string, any>;

declare const impl: {
  ManifestError: typeof ManifestError;
  VALID_PERMISSIONS: typeof VALID_PERMISSIONS;
  validateExtensionManifest: typeof validateExtensionManifest;
};

export default impl;
