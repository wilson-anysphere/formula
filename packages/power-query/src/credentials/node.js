import { OsKeychainProvider } from "../../../security/crypto/keychain/osKeychain.js";

import { KeychainCredentialStore } from "./stores/keychain.js";
import { EncryptedFileCredentialStore } from "./stores/encryptedFile.node.js";

/**
 * Create a Node-friendly credential store.
 *
 * - On platforms with a supported OS keychain, credentials are stored directly
 *   in the keychain.
 * - Otherwise, credentials are stored in an encrypted-on-disk file.
 *
 * @param {{
 *   filePath: string;
 *   keychainProvider?: any;
 *   service?: string;
 * }} opts
 */
export function createNodeCredentialStore(opts) {
  if (!opts?.filePath) throw new TypeError("filePath is required");
  const provider = opts.keychainProvider ?? new OsKeychainProvider();
  const service = opts.service ?? "formula.power-query";

  // OsKeychainProvider only supports non-interactive writes on macOS today.
  // Custom providers (e.g. tests) are assumed to support writes.
  const canWriteKeychain =
    provider &&
    typeof provider.setSecret === "function" &&
    typeof provider.deleteSecret === "function" &&
    (provider.platform == null || provider.platform === "darwin");
  if (canWriteKeychain) {
    return new KeychainCredentialStore({ keychainProvider: provider, service });
  }

  return new EncryptedFileCredentialStore({
    filePath: opts.filePath,
    // Best-effort: still use the provider for key material if it supports writes.
    keychainProvider: canWriteKeychain ? provider : null,
    keychainService: service,
  });
}
