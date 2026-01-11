/**
 * Credential scoping helpers.
 *
 * These scopes are intentionally JSON-serializable so they can be used as
 * stable keys in cache signatures and persisted stores.
 */

/**
 * @typedef {{
 *   type: "http";
 *   origin: string;
 *   realm?: string | null;
 * }} HttpCredentialScope
 *
 * @typedef {{
 *   type: "file";
 *   match: "exact" | "prefix";
 *   path: string;
 * }} FileCredentialScope
 *
 * @typedef {{
 *   type: "sql";
 *   server: string;
 *   database?: string | null;
 *   user?: string | null;
 * }} SqlCredentialScope
 *
 * @typedef {HttpCredentialScope | FileCredentialScope | SqlCredentialScope} CredentialScope
 */

/**
 * @param {{ url: string; realm?: string | null }} args
 * @returns {HttpCredentialScope}
 */
export function httpScope({ url, realm = null }) {
  const origin = new URL(url).origin;
  const out = { type: "http", origin };
  if (realm != null && realm !== "") out.realm = realm;
  return out;
}

/**
 * @param {{ path: string }} args
 * @returns {FileCredentialScope}
 */
export function fileScopeExact({ path }) {
  return { type: "file", match: "exact", path };
}

/**
 * @param {{ pathPrefix: string }} args
 * @returns {FileCredentialScope}
 */
export function fileScopePrefix({ pathPrefix }) {
  return { type: "file", match: "prefix", path: pathPrefix };
}

/**
 * @param {{ server: string; database?: string | null; user?: string | null }} args
 * @returns {SqlCredentialScope}
 */
export function sqlScope({ server, database = null, user = null }) {
  const out = { type: "sql", server };
  if (database != null && database !== "") out.database = database;
  if (user != null && user !== "") out.user = user;
  return out;
}

