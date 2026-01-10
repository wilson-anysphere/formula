import { isHyperlinkActivation } from "./activation.js";
import { navigateInternalHyperlink } from "./navigate.js";
import { openExternalHyperlink } from "./openExternal.js";

/**
 * @typedef {{
 *   range: { start: { row: number, col: number }, end: { row: number, col: number } },
 *   target:
 *     | { type: "external_url", uri: string }
 *     | { type: "email", uri: string }
 *     | { type: "internal", sheet: string, cell: { row: number, col: number } },
 *   display?: string,
 *   tooltip?: string,
 *   rel_id?: string,
 * }} Hyperlink
 */

/**
 * @typedef {{
 *   navigator?: import("./navigate.js").WorkbookNavigator,
 *   shellOpen?: (uri: string) => Promise<void>,
 *   confirmUntrustedProtocol?: (message: string) => Promise<boolean>,
 *   permissions?: { request: (permission: string, context: any) => Promise<boolean> },
 * }} HyperlinkActivationDeps
 */

/**
 * Handle a mouse click on a hyperlink. Only activates on Ctrl/Cmd+click.
 *
 * @param {Hyperlink} hyperlink
 * @param {{ button?: number, metaKey?: boolean, ctrlKey?: boolean } | null | undefined} event
 * @param {HyperlinkActivationDeps} deps
 * @returns {Promise<boolean>} Whether the hyperlink was activated
 */
export async function handleHyperlinkClick(hyperlink, event, deps) {
  if (!isHyperlinkActivation(event)) return false;
  return activateHyperlink(hyperlink, deps);
}

/**
 * Activate a hyperlink (internal navigation or external open).
 *
 * @param {Hyperlink} hyperlink
 * @param {HyperlinkActivationDeps} deps
 * @returns {Promise<boolean>}
 */
export async function activateHyperlink(hyperlink, deps) {
  if (!hyperlink || !hyperlink.target) return false;
  const target = hyperlink.target;

  if (target.type === "internal") {
    if (!deps?.navigator) throw new Error("Internal hyperlink activation requires deps.navigator");
    await navigateInternalHyperlink(target, deps.navigator);
    return true;
  }

  if (target.type === "external_url" || target.type === "email") {
    if (!deps?.shellOpen) throw new Error("External hyperlink activation requires deps.shellOpen");
    return openExternalHyperlink(target.uri, {
      shellOpen: deps.shellOpen,
      confirmUntrustedProtocol: deps.confirmUntrustedProtocol,
      permissions: deps.permissions,
    });
  }

  return false;
}

