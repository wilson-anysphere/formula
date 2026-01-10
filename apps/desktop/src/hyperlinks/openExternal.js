const DEFAULT_ALLOWED_PROTOCOLS = new Set(["http", "https", "mailto"]);
const BLOCKED_PROTOCOLS = new Set(["javascript", "data"]);

/**
 * @typedef {{
 *   shellOpen: (uri: string) => Promise<void>,
 *   confirmUntrustedProtocol?: (message: string) => Promise<boolean>,
 *   permissions?: { request: (permission: string, context: any) => Promise<boolean> },
 *   allowedProtocols?: Set<string>,
 * }} OpenExternalHyperlinkDeps
 */

/**
 * Open an external hyperlink via the host OS (e.g. Tauri's `shell.open`).
 *
 * Security:
 * - http/https/mailto allowed by default
 * - javascript:/data: are always blocked
 * - any other protocol requires an explicit confirmation prompt
 *
 * Permission integration:
 * - if a `permissions` manager is provided, `external_navigation` is requested
 *   (and `external_navigation_untrusted_protocol` for non-allowlisted schemes).
 *
 * @param {string} uri
 * @param {OpenExternalHyperlinkDeps} deps
 * @returns {Promise<boolean>} Whether the link was opened
 */
export async function openExternalHyperlink(uri, deps) {
  if (!deps || typeof deps.shellOpen !== "function") {
    throw new Error("openExternalHyperlink requires deps.shellOpen");
  }

  let parsed;
  try {
    parsed = new URL(uri);
  } catch (err) {
    throw new Error(`Invalid URL: ${uri}`);
  }

  const protocol = parsed.protocol.replace(":", "").toLowerCase();
  if (BLOCKED_PROTOCOLS.has(protocol)) {
    return false;
  }

  const allowlist = deps.allowedProtocols ?? DEFAULT_ALLOWED_PROTOCOLS;
  const isTrusted = allowlist.has(protocol);
  if (!isTrusted) {
    const confirm = deps.confirmUntrustedProtocol ?? (async () => false);
    const ok = await confirm(
      `Open link with untrusted protocol "${protocol}"?\n\n${uri}\n\n` +
        `Only http/https/mailto are trusted by default.`,
    );
    if (!ok) return false;
  }

  if (deps.permissions && typeof deps.permissions.request === "function") {
    const ok = await deps.permissions.request("external_navigation", { uri, protocol });
    if (!ok) return false;
    if (!isTrusted) {
      const ok2 = await deps.permissions.request("external_navigation_untrusted_protocol", {
        uri,
        protocol,
      });
      if (!ok2) return false;
    }
  }

  await deps.shellOpen(uri);
  return true;
}

