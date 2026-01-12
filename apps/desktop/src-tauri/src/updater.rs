use serde::Serialize;
use tauri::{AppHandle, Emitter};
use tauri_plugin_updater::UpdaterExt;

#[derive(Clone, Copy, Debug, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum UpdateCheckSource {
    Startup,
    Manual,
}

pub fn spawn_update_check(app: &AppHandle, source: UpdateCheckSource) {
    let handle = app.clone();
    let _ = handle.emit(
        "update-check-started",
        serde_json::json!({ "source": source }),
    );
    tauri::async_runtime::spawn(async move {
        match check_for_updates(&handle, source).await {
            Ok(Some(payload)) => {
                let _ = handle.emit("update-available", payload);
            }
            Ok(None) => {
                let _ = handle.emit("update-not-available", serde_json::json!({ "source": source }));
            }
            Err(err) => {
                let msg = err.to_string();
                let _ = handle.emit(
                    "update-check-error",
                    serde_json::json!({ "source": source, "message": msg }),
                );
                eprintln!("updater check failed: {err}");
            }
        }
    });
}

async fn check_for_updates(
    app: &AppHandle,
    source: UpdateCheckSource,
) -> tauri_plugin_updater::Result<Option<serde_json::Value>> {
    let update = app.updater()?.check().await?;
    Ok(update.map(|update| {
        serde_json::json!({
            "source": source,
            "version": update.version,
            "body": update.body,
        })
    }))
}

