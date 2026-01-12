use tauri::menu::{Menu, MenuItem, PredefinedMenuItem};
use tauri::tray::{MouseButton, MouseButtonState, TrayIconBuilder, TrayIconEvent};
use tauri::{App, AppHandle, Emitter, Manager};

use desktop::tray_status::TrayStatusState;

// Keep these tray-menu item ids distinct from the menu bar `ITEM_*` constants in
// `menu.rs`. The desktop capability tests scan Rust sources and build a naive
// `const name -> value` map to verify emitted/listened events are allowlisted; reusing
// the same identifiers across modules can confuse those checks.
const TRAY_ITEM_NEW: &str = "new";
const TRAY_ITEM_OPEN: &str = "open";
const TRAY_ITEM_CHECK_UPDATES: &str = "check_updates";
const TRAY_ITEM_QUIT: &str = "quit";

pub fn init(app: &mut App) -> tauri::Result<()> {
    let handle = app.handle();

    let new = MenuItem::with_id(handle, TRAY_ITEM_NEW, "New Workbook", true, None::<&str>)?;
    let open = MenuItem::with_id(handle, TRAY_ITEM_OPEN, "Open…", true, None::<&str>)?;
    let check_updates = MenuItem::with_id(
        handle,
        TRAY_ITEM_CHECK_UPDATES,
        "Check for Updates",
        true,
        None::<&str>,
    )?;
    let quit = MenuItem::with_id(handle, TRAY_ITEM_QUIT, "Quit", true, None::<&str>)?;

    let menu = Menu::with_items(
        handle,
        &[
            &new,
            &open,
            &PredefinedMenuItem::separator(handle)?,
            &check_updates,
            &PredefinedMenuItem::separator(handle)?,
            &quit,
        ],
    )?;

    let tray = TrayIconBuilder::new()
        .icon(tauri::include_image!("icons/tray.png"))
        .tooltip("Formula")
        .menu(&menu)
        .on_menu_event(|app, event| match event.id().as_ref() {
            TRAY_ITEM_NEW => {
                let _ = app.emit("tray-new", ());
                show_main_window(app);
            }
            TRAY_ITEM_OPEN => {
                let _ = app.emit("tray-open", ());
                show_main_window(app);
            }
            TRAY_ITEM_CHECK_UPDATES => {
                desktop::updater::spawn_update_check(app, desktop::updater::UpdateCheckSource::Manual);
            }
            TRAY_ITEM_QUIT => {
                // Delegate quit-handling to the frontend so it can:
                // - fire `Workbook_BeforeClose` macros
                // - prompt for unsaved changes
                // - decide whether to exit or keep running
                show_main_window(app);
                let _ = app.emit("tray-quit", ());
            }
            _ => {}
        })
        .on_tray_icon_event(|tray, event| {
            if let TrayIconEvent::Click {
                button: MouseButton::Left,
                button_state: MouseButtonState::Up,
                ..
            } = event
            {
                show_main_window(tray.app_handle());
            }
        })
        .build(app)?;

    app.state::<TrayStatusState>().inner().set_tray(tray);

    Ok(())
}

fn show_main_window(app: &AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.show();
        let _ = window.set_focus();
    }
}
