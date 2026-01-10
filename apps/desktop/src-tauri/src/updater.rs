use tauri::{AppHandle, Emitter};
use tauri_plugin_updater::UpdaterExt;

pub fn spawn_update_check(app: &AppHandle) {
    let handle = app.clone();
    tauri::async_runtime::spawn(async move {
        if let Err(err) = check_for_updates(&handle).await {
            eprintln!("updater check failed: {err}");
        }
    });
}

async fn check_for_updates(app: &AppHandle) -> tauri_plugin_updater::Result<()> {
    let update = app.updater()?.check().await?;
    let Some(update) = update else {
        return Ok(());
    };

    let payload = serde_json::json!({
        "version": update.version,
        "body": update.body,
    });

    let _ = app.emit("update-available", payload);

    Ok(())
}
