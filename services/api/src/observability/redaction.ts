const SHARE_LINK_REDEEM_PATH = /^\/share-links\/[^/]+\/redeem$/;

/**
 * Redacts sensitive path segments + strips query params.
 *
 * - Query params can contain bearer/session/OIDC tokens.
 * - `/share-links/:token/redeem` contains a secret token in the path.
 */
export function sanitizeUrlPath(rawUrl: string): string {
  const path = rawUrl.split("?")[0] ?? rawUrl;
  if (SHARE_LINK_REDEEM_PATH.test(path)) return "/share-links/:token/redeem";
  return path;
}

