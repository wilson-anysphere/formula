use tauri::AppHandle;
use tauri_plugin_global_shortcut::GlobalShortcutExt;

pub fn register(app: &AppHandle) -> Result<(), tauri_plugin_global_shortcut::Error> {
    // Note: `tauri-plugin-global-shortcut` dispatches shortcut events to the frontend plugin API.
    // We register the accelerators here so they are available globally (even when the window is
    // not focused). The UI layer can subscribe via `@tauri-apps/plugin-global-shortcut`.
    app.global_shortcut().register("CmdOrCtrl+Shift+O")?;
    app.global_shortcut().register("CmdOrCtrl+Shift+P")?;

    Ok(())
}
