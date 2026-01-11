use tauri::menu::{Menu, MenuItem, PredefinedMenuItem};
use tauri::tray::{MouseButton, MouseButtonState, TrayIconBuilder, TrayIconEvent};
use tauri::{App, AppHandle, Emitter, Manager};

const ITEM_NEW: &str = "new";
const ITEM_OPEN: &str = "open";
const ITEM_CHECK_UPDATES: &str = "check_updates";
const ITEM_QUIT: &str = "quit";

pub fn init(app: &mut App) -> tauri::Result<()> {
    let handle = app.handle();

    let new = MenuItem::with_id(handle, ITEM_NEW, "New Workbook", true, None::<&str>)?;
    let open = MenuItem::with_id(handle, ITEM_OPEN, "Open…", true, None::<&str>)?;
    let check_updates =
        MenuItem::with_id(handle, ITEM_CHECK_UPDATES, "Check for Updates", true, None::<&str>)?;
    let quit = MenuItem::with_id(handle, ITEM_QUIT, "Quit", true, None::<&str>)?;

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

    TrayIconBuilder::new()
        .icon(tauri::include_image!("icons/tray.png"))
        .menu(&menu)
        .on_menu_event(|app, event| match event.id().as_ref() {
            ITEM_NEW => {
                let _ = app.emit("tray-new", ());
                show_main_window(app);
            }
            ITEM_OPEN => {
                let _ = app.emit("tray-open", ());
                show_main_window(app);
            }
            ITEM_CHECK_UPDATES => crate::updater::spawn_update_check(app),
            ITEM_QUIT => {
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

    Ok(())
}

fn show_main_window(app: &AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.show();
        let _ = window.set_focus();
    }
}
