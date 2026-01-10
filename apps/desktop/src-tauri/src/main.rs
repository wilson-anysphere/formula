#![cfg_attr(
    all(not(debug_assertions), target_os = "windows"),
    windows_subsystem = "windows"
)]

use formula_desktop_tauri::commands;
use formula_desktop_tauri::state::{AppState, SharedAppState};
use std::sync::{Arc, Mutex};
use tauri::Manager;

fn main() {
    let state: SharedAppState = Arc::new(Mutex::new(AppState::new()));

    tauri::Builder::default()
        .manage(state)
        .invoke_handler(tauri::generate_handler![
            commands::open_workbook,
            commands::save_workbook,
            commands::get_cell,
            commands::set_cell,
            commands::get_range,
            commands::set_range,
            commands::recalculate,
            commands::undo,
            commands::redo,
        ])
        .on_window_event(|event| match event.event() {
            tauri::WindowEvent::CloseRequested { api, .. } => {
                let state = event.window().state::<SharedAppState>();
                let state = state.lock().unwrap();
                if state.has_unsaved_changes() {
                    api.prevent_close();
                    let _ = event.window().emit("unsaved-changes", ());
                }
            }
            tauri::WindowEvent::FileDrop(file_drop) => {
                if let tauri::FileDropEvent::Dropped(paths) = file_drop {
                    let payload: Vec<String> = paths
                        .iter()
                        .map(|p| p.to_string_lossy().to_string())
                        .collect();
                    let _ = event.window().emit("file-dropped", payload);
                }
            }
            _ => {}
        })
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
