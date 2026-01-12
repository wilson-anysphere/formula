use tauri::AppHandle;
use tauri_plugin_global_shortcut::GlobalShortcutExt;

pub fn register(app: &AppHandle) -> Result<(), tauri_plugin_global_shortcut::Error> {
    // Register global accelerators so they fire even when the window is not focused.
    //
    // In this repo, we handle shortcuts in Rust (see `main.rs`'s
    // `tauri_plugin_global_shortcut::Builder::with_handler(...)`) and forward them to the
    // frontend as explicit `shortcut-*` events. We intentionally do not rely on the plugin's
    // frontend API surface for these built-in shortcuts.
    app.global_shortcut().register("CmdOrCtrl+Shift+O")?;
    app.global_shortcut().register("CmdOrCtrl+Shift+P")?;

    Ok(())
}
