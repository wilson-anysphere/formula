const TAURI_GLOBALS = ["__TAURI__", "__TAURI_IPC__", "__TAURI_INVOKE__", "__TAURI_INTERNALS__", "__TAURI_METADATA__"];

/**
 * Best-effort removal of Tauri-injected globals from an extension worker scope.
 *
 * We pass in the `lockDownGlobal` implementation so this helper stays pure and can be unit-tested
 * outside a real WorkerGlobalScope.
 *
 * @param {(prop: string, value: any) => void} lockDownGlobal
 */
function lockDownTauriGlobals(lockDownGlobal) {
  if (typeof lockDownGlobal !== "function") return;
  for (const prop of TAURI_GLOBALS) {
    try {
      lockDownGlobal(prop, undefined);
    } catch {
      // Ignore failures (eg: non-configurable properties) so non-Tauri environments don't break.
    }
  }
}

export { TAURI_GLOBALS, lockDownTauriGlobals };
