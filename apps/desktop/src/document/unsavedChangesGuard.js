import { t } from "../i18n/index.js";

/**
 * @param {{ addEventListener: (type: string, listener: (event: any) => void) => void, removeEventListener: (type: string, listener: (event: any) => void) => void }} target
 * @param {{ isDirty: boolean }} controller
 * @param {{ message?: string }} [options]
 * @returns {() => void} cleanup
 */
export function installUnsavedChangesPrompt(target, controller, options = {}) {
  const message = options.message ?? t("prompt.unsavedChanges");

  /** @param {any} event */
  function onBeforeUnload(event) {
    if (!controller.isDirty) return;
    // The exact message is ignored by most browsers, but returnValue must be set.
    event.preventDefault?.();
    event.returnValue = message;
    return message;
  }

  target.addEventListener("beforeunload", onBeforeUnload);
  return () => target.removeEventListener("beforeunload", onBeforeUnload);
}
