const DEFAULT_ALLOWED_PROTOCOLS = new Set(["http", "https", "mailto"]);
// Schemes that should never be opened via external navigation (even with user confirmation).
const BLOCKED_PROTOCOLS = new Set(["javascript", "data", "file"]);

import { hasTauri } from "../tauri/api.js";

/**
 * @typedef {{
 *   shellOpen: (uri: string) => Promise<void>,
 *   confirmUntrustedProtocol?: (message: string) => Promise<boolean>,
 *   permissions?: { request: (permission: string, context: any) => Promise<boolean> },
 *   allowedProtocols?: Set<string>,
 * }} OpenExternalHyperlinkDeps
 */

/**
 * Open an external hyperlink via the host OS.
 *
 * In the desktop/Tauri shell, callers should pass the `shellOpen` helper from
 * `apps/desktop/src/tauri/shellOpen.ts`, which routes through the Rust `open_external_url`
 * command (strict scheme allowlist enforced in Rust).
 *
 * Security:
 * - http/https/mailto allowed by default
 * - javascript:/data:/file: are always blocked
 * - in Tauri builds, any other protocol is blocked (Rust allowlist boundary)
 * - in web builds, other protocols may be allowed with an explicit confirmation prompt
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
  if ((protocol === "http" || protocol === "https") && (parsed.username || parsed.password)) {
    // Userinfo can be used to construct misleading URLs (e.g. `https://trusted.com@evil.com/...`)
    // and is never required for typical external navigation.
    return false;
  }

  // In the desktop shell, link opening is routed through a Rust command that enforces a strict
  // scheme allowlist. Keep the JS allowlist in sync (do not allow overrides in Tauri builds).
  const isTauri = hasTauri();
  const allowlist = isTauri ? DEFAULT_ALLOWED_PROTOCOLS : deps.allowedProtocols ?? DEFAULT_ALLOWED_PROTOCOLS;
  const isTrusted = allowlist.has(protocol);
  if (!isTrusted) {
    // In the desktop shell, link opening is routed through a Rust command that enforces a strict
    // scheme allowlist. Avoid prompting for protocols that will be rejected anyway.
    if (isTauri) return false;

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
