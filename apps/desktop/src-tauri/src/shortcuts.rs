use tauri::AppHandle;
use tauri_plugin_global_shortcut::GlobalShortcutExt;

pub fn register(app: &AppHandle) -> Result<(), tauri_plugin_global_shortcut::Error> {
    // Register global accelerators so they fire even when the window is not focused.
    //
    // In this repo, we handle shortcuts in Rust (see `main.rs`'s
    // `tauri_plugin_global_shortcut::Builder::with_handler(...)`) and forward them to the
    // frontend as explicit events (`shortcut-quick-open`, `shortcut-command-palette`).
    //
    // Note: if you add a new shortcut event, you must also update the Tauri v2 event allowlist in
    // `apps/desktop/src-tauri/capabilities/main.json` (and the `eventPermissions.vitest.ts` guardrail).
    app.global_shortcut().register("CmdOrCtrl+Shift+O")?;
    app.global_shortcut().register("CmdOrCtrl+Shift+P")?;

    Ok(())
}
