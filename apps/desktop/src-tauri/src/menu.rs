use tauri::menu::{Menu, MenuItem, PredefinedMenuItem, Submenu};
use tauri::{App, AppHandle, Emitter, Manager};

pub const ITEM_NEW: &str = "menu-new";
pub const ITEM_OPEN: &str = "menu-open";
pub const ITEM_SAVE: &str = "menu-save";
pub const ITEM_SAVE_AS: &str = "menu-save-as";
pub const ITEM_PRINT: &str = "menu-print";
pub const ITEM_PRINT_PREVIEW: &str = "menu-print-preview";
pub const ITEM_EXPORT_PDF: &str = "menu-export-pdf";
pub const ITEM_CLOSE_WINDOW: &str = "menu-close-window";
pub const ITEM_QUIT: &str = "menu-quit";
pub const ITEM_UNDO: &str = "menu-undo";
pub const ITEM_REDO: &str = "menu-redo";
pub const ITEM_CUT: &str = "menu-cut";
pub const ITEM_COPY: &str = "menu-copy";
pub const ITEM_PASTE: &str = "menu-paste";
pub const ITEM_PASTE_SPECIAL: &str = "menu-paste-special";
pub const ITEM_SELECT_ALL: &str = "menu-select-all";
pub const ITEM_ABOUT: &str = "menu-about";
pub const ITEM_CHECK_UPDATES: &str = "menu-check-updates";
pub const ITEM_OPEN_RELEASE_PAGE: &str = "menu-open-release-page";
pub const ITEM_ZOOM_IN: &str = "menu-zoom-in";
pub const ITEM_ZOOM_OUT: &str = "menu-zoom-out";
pub const ITEM_ZOOM_RESET: &str = "menu-zoom-reset";

#[cfg(debug_assertions)]
pub const ITEM_RELOAD: &str = "menu-reload";
#[cfg(debug_assertions)]
pub const ITEM_TOGGLE_DEVTOOLS: &str = "menu-toggle-devtools";

pub fn init(app: &mut App) -> tauri::Result<()> {
    let handle = app.handle();
    let menu = build_menu(&handle)?;

    app.set_menu(menu)?;
    Ok(())
}

pub fn on_menu_event(app: &AppHandle, event: tauri::menu::MenuEvent) {
    match event.id().as_ref() {
        ITEM_NEW => {
            show_main_window(app);
            let _ = app.emit(ITEM_NEW, ());
        }
        ITEM_OPEN => {
            show_main_window(app);
            let _ = app.emit(ITEM_OPEN, ());
        }
        ITEM_SAVE => {
            show_main_window(app);
            let _ = app.emit(ITEM_SAVE, ());
        }
        ITEM_SAVE_AS => {
            show_main_window(app);
            let _ = app.emit(ITEM_SAVE_AS, ());
        }
        ITEM_PRINT => {
            show_main_window(app);
            let _ = app.emit(ITEM_PRINT, ());
        }
        ITEM_PRINT_PREVIEW => {
            show_main_window(app);
            let _ = app.emit(ITEM_PRINT_PREVIEW, ());
        }
        ITEM_EXPORT_PDF => {
            show_main_window(app);
            let _ = app.emit(ITEM_EXPORT_PDF, ());
        }
        ITEM_CLOSE_WINDOW => {
            show_main_window(app);
            let _ = app.emit(ITEM_CLOSE_WINDOW, ());
        }
        ITEM_QUIT => {
            show_main_window(app);
            let _ = app.emit(ITEM_QUIT, ());
        }
        ITEM_UNDO => {
            show_main_window(app);
            let _ = app.emit(ITEM_UNDO, ());
        }
        ITEM_REDO => {
            show_main_window(app);
            let _ = app.emit(ITEM_REDO, ());
        }
        ITEM_CUT => {
            show_main_window(app);
            let _ = app.emit(ITEM_CUT, ());
        }
        ITEM_COPY => {
            show_main_window(app);
            let _ = app.emit(ITEM_COPY, ());
        }
        ITEM_PASTE => {
            show_main_window(app);
            let _ = app.emit(ITEM_PASTE, ());
        }
        ITEM_PASTE_SPECIAL => {
            show_main_window(app);
            let _ = app.emit(ITEM_PASTE_SPECIAL, ());
        }
        ITEM_SELECT_ALL => {
            show_main_window(app);
            let _ = app.emit(ITEM_SELECT_ALL, ());
        }
        ITEM_ZOOM_IN => {
            show_main_window(app);
            let _ = app.emit(ITEM_ZOOM_IN, ());
        }
        ITEM_ZOOM_OUT => {
            show_main_window(app);
            let _ = app.emit(ITEM_ZOOM_OUT, ());
        }
        ITEM_ZOOM_RESET => {
            show_main_window(app);
            let _ = app.emit(ITEM_ZOOM_RESET, ());
        }
        ITEM_ABOUT => {
            show_main_window(app);
            let _ = app.emit(ITEM_ABOUT, ());
        }
        ITEM_CHECK_UPDATES => {
            // Reuse the existing updater flow (also used by the tray menu).
            crate::updater::spawn_update_check(app, crate::updater::UpdateCheckSource::Manual);
            let _ = app.emit(ITEM_CHECK_UPDATES, ());
        }
        ITEM_OPEN_RELEASE_PAGE => {
            show_main_window(app);
            let _ = app.emit(ITEM_OPEN_RELEASE_PAGE, ());
        }
        #[cfg(debug_assertions)]
        ITEM_RELOAD => {
            if let Some(window) = app.get_webview_window("main") {
                // Reload the webview without requiring frontend changes.
                let _ = window.eval("window.location.reload()");
            }
        }
        #[cfg(debug_assertions)]
        ITEM_TOGGLE_DEVTOOLS => {
            if let Some(window) = app.get_webview_window("main") {
                if window.is_devtools_open() {
                    let _ = window.close_devtools();
                } else {
                    let _ = window.open_devtools();
                }
            }
        }
        _ => {}
    }
}

fn build_menu(handle: &AppHandle) -> tauri::Result<Menu<tauri::Wry>> {
    // Prefer the configured product name (tauri.conf.json `productName`) for menu labels.
    // `package_info().name` is the Cargo crate name (e.g. `desktop`), which isn't user-facing.
    let app_name = handle
        .config()
        .product_name
        .clone()
        .unwrap_or_else(|| handle.package_info().name.clone());

    let new = MenuItem::with_id(handle, ITEM_NEW, "New", true, Some("CmdOrCtrl+N"))?;
    let open = MenuItem::with_id(handle, ITEM_OPEN, "Open…", true, Some("CmdOrCtrl+O"))?;
    let save = MenuItem::with_id(handle, ITEM_SAVE, "Save", true, Some("CmdOrCtrl+S"))?;
    let save_as = MenuItem::with_id(
        handle,
        ITEM_SAVE_AS,
        "Save As…",
        true,
        Some("CmdOrCtrl+Shift+S"),
    )?;
    let print = MenuItem::with_id(handle, ITEM_PRINT, "Print…", true, Some("CmdOrCtrl+P"))?;
    let print_preview = MenuItem::with_id(handle, ITEM_PRINT_PREVIEW, "Print Preview", true, None::<&str>)?;
    let export_pdf = MenuItem::with_id(handle, ITEM_EXPORT_PDF, "Export PDF…", true, None::<&str>)?;
    let close_window = MenuItem::with_id(
        handle,
        ITEM_CLOSE_WINDOW,
        "Close Window",
        true,
        Some("CmdOrCtrl+W"),
    )?;

    let quit_label = format!("Quit {app_name}");
    let quit = MenuItem::with_id(handle, ITEM_QUIT, quit_label, true, Some("CmdOrCtrl+Q"))?;

    let undo = MenuItem::with_id(handle, ITEM_UNDO, "Undo", true, Some("CmdOrCtrl+Z"))?;
    let redo_accelerator = if cfg!(target_os = "macos") {
        "CmdOrCtrl+Shift+Z"
    } else {
        "CmdOrCtrl+Y"
    };
    let redo = MenuItem::with_id(handle, ITEM_REDO, "Redo", true, Some(redo_accelerator))?;
    let cut = MenuItem::with_id(handle, ITEM_CUT, "Cut", true, Some("CmdOrCtrl+X"))?;
    let copy = MenuItem::with_id(handle, ITEM_COPY, "Copy", true, Some("CmdOrCtrl+C"))?;
    let paste = MenuItem::with_id(handle, ITEM_PASTE, "Paste", true, Some("CmdOrCtrl+V"))?;
    // No accelerator: Cmd/Ctrl+Shift+V is handled by the WebView (Paste Special command).
    // Adding a menu accelerator risks triggering both the menu event and the keybinding handler.
    let paste_special = MenuItem::with_id(handle, ITEM_PASTE_SPECIAL, "Paste Special…", true, None::<&str>)?;
    let select_all = MenuItem::with_id(
        handle,
        ITEM_SELECT_ALL,
        "Select All",
        true,
        Some("CmdOrCtrl+A"),
    )?;

    let zoom_in = MenuItem::with_id(handle, ITEM_ZOOM_IN, "Zoom In", true, None::<&str>)?;
    let zoom_out = MenuItem::with_id(handle, ITEM_ZOOM_OUT, "Zoom Out", true, None::<&str>)?;
    let zoom_reset = MenuItem::with_id(handle, ITEM_ZOOM_RESET, "Actual Size", true, None::<&str>)?;

    let about_label = format!("About {app_name}");
    let about = MenuItem::with_id(handle, ITEM_ABOUT, about_label, true, None::<&str>)?;
    let check_updates = MenuItem::with_id(
        handle,
        ITEM_CHECK_UPDATES,
        "Check for Updates…",
        true,
        None::<&str>,
    )?;
    let open_release_page =
        MenuItem::with_id(handle, ITEM_OPEN_RELEASE_PAGE, "Open Release Page", true, None::<&str>)?;

    let sep_file_1 = PredefinedMenuItem::separator(handle)?;
    let sep_file_2 = PredefinedMenuItem::separator(handle)?;
    #[cfg(not(target_os = "macos"))]
    let sep_file_3 = PredefinedMenuItem::separator(handle)?;
    let sep_edit_1 = PredefinedMenuItem::separator(handle)?;
    let sep_edit_2 = PredefinedMenuItem::separator(handle)?;

    let file_menu = {
        #[cfg(target_os = "macos")]
        {
            Submenu::with_items(
                handle,
                "File",
                true,
                &[
                    &new,
                    &open,
                    &sep_file_1,
                    &save,
                    &save_as,
                    &print,
                    &print_preview,
                    &export_pdf,
                    &sep_file_2,
                    &close_window,
                ],
            )?
        }

        #[cfg(not(target_os = "macos"))]
        {
            Submenu::with_items(
                handle,
                "File",
                true,
                &[
                    &new,
                    &open,
                    &sep_file_1,
                    &save,
                    &save_as,
                    &print,
                    &print_preview,
                    &export_pdf,
                    &sep_file_2,
                    &close_window,
                    &sep_file_3,
                    &quit,
                ],
            )?
        }
    };

    let edit_menu = Submenu::with_items(
        handle,
        "Edit",
        true,
        &[
            &undo,
            &redo,
            &sep_edit_1,
            &cut,
            &copy,
            &paste,
            &paste_special,
            &sep_edit_2,
            &select_all,
        ],
    )?;

    let view_menu = {
        #[cfg(debug_assertions)]
        {
            let reload =
                MenuItem::with_id(handle, ITEM_RELOAD, "Reload", true, Some("CmdOrCtrl+R"))?;
            let toggle_devtools = MenuItem::with_id(
                handle,
                ITEM_TOGGLE_DEVTOOLS,
                "Toggle DevTools",
                true,
                Some("CmdOrCtrl+Alt+I"),
            )?;
            let sep_view = PredefinedMenuItem::separator(handle)?;
            let sep_zoom = PredefinedMenuItem::separator(handle)?;
            Submenu::with_items(
                handle,
                "View",
                true,
                &[
                    &reload,
                    &sep_view,
                    &zoom_in,
                    &zoom_out,
                    &zoom_reset,
                    &sep_zoom,
                    &toggle_devtools,
                ],
            )?
        }

        #[cfg(not(debug_assertions))]
        {
            Submenu::with_items(handle, "View", true, &[&zoom_in, &zoom_out, &zoom_reset])?
        }
    };

    let help_menu = {
        #[cfg(target_os = "macos")]
        {
            Submenu::with_items(handle, "Help", true, &[&check_updates, &open_release_page])?
        }

        #[cfg(not(target_os = "macos"))]
        {
            let sep_help = PredefinedMenuItem::separator(handle)?;
            Submenu::with_items(
                handle,
                "Help",
                true,
                &[&about, &sep_help, &check_updates, &open_release_page],
            )?
        }
    };

    #[cfg(target_os = "macos")]
    let menu = {
        // On macOS, the first submenu is treated as the "app menu".
        let sep_app = PredefinedMenuItem::separator(handle)?;
        let app_menu = Submenu::with_items(handle, app_name, true, &[&about, &sep_app, &quit])?;
        Menu::with_items(
            handle,
            &[&app_menu, &file_menu, &edit_menu, &view_menu, &help_menu],
        )?
    };

    #[cfg(not(target_os = "macos"))]
    let menu = Menu::with_items(handle, &[&file_menu, &edit_menu, &view_menu, &help_menu])?;

    Ok(menu)
}

fn show_main_window(app: &AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.show();
        let _ = window.set_focus();
    }
}
