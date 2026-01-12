use serde::{Deserialize, Serialize};
use std::sync::atomic::{AtomicBool, Ordering};
use tauri::{AppHandle, Emitter};
use tauri_plugin_updater::UpdaterExt;

static UPDATE_CHECK_IN_FLIGHT: AtomicBool = AtomicBool::new(false);

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum UpdateCheckSource {
    Startup,
    Manual,
}

pub fn spawn_update_check(app: &AppHandle, source: UpdateCheckSource) {
    if UPDATE_CHECK_IN_FLIGHT.swap(true, Ordering::SeqCst) {
        // Startup checks should not interrupt the user; manual checks should provide feedback so
        // the frontend can display "Already checking..." without triggering a second network call.
        if matches!(source, UpdateCheckSource::Manual) {
            let _ = app.emit(
                "update-check-already-running",
                serde_json::json!({ "source": source }),
            );
        }
        return;
    }

    let handle = app.clone();
    let _ = handle.emit("update-check-started", serde_json::json!({ "source": source }));
    tauri::async_runtime::spawn(async move {
        struct UpdateCheckInFlightGuard;
        impl Drop for UpdateCheckInFlightGuard {
            fn drop(&mut self) {
                UPDATE_CHECK_IN_FLIGHT.store(false, Ordering::SeqCst);
            }
        }
        let _guard = UpdateCheckInFlightGuard;

        match check_for_updates(&handle, source).await {
            Ok(Some(payload)) => {
                let _ = handle.emit("update-available", payload);
            }
            Ok(None) => {
                let _ =
                    handle.emit("update-not-available", serde_json::json!({ "source": source }));
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
