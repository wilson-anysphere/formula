//! Guardrails to ensure privileged IPC commands enforce origin checks.
//!
//! These are source-level tests so they run in headless CI without the `desktop` feature
//! (which would pull in the system WebView toolchain on Linux).
//!
//! The goal is defense-in-depth: even if a WebView is navigated to untrusted content, privileged
//! `#[tauri::command]` functions should still validate the caller's window label and origin URL.

use std::fs;
use std::path::PathBuf;

const MAX_SCAN_LINES: usize = 300;

fn repo_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(relative)
}

fn read_repo_file(relative: &str) -> String {
    let path = repo_path(relative);
    fs::read_to_string(&path).unwrap_or_else(|err| panic!("failed to read {}: {err}", path.display()))
}

fn is_ident_char(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_'
}

fn find_fn_start(src: &str, fn_name: &str, file: &str) -> usize {
    let patterns = [
        format!("pub async fn {fn_name}"),
        format!("pub fn {fn_name}"),
        format!("async fn {fn_name}"),
        format!("fn {fn_name}"),
    ];

    for pat in patterns {
        let mut search_start = 0usize;
        while let Some(rel_idx) = src[search_start..].find(&pat) {
            let idx = search_start + rel_idx;
            let after_idx = idx + pat.len();
            let next = src.get(after_idx..).and_then(|rest| rest.chars().next());
            // Ensure we matched the full identifier name (avoid prefix matches like
            // `fn verify_ed25519_signature_payload` when we were looking for
            // `fn verify_ed25519_signature`).
            if next.map_or(true, |ch| !is_ident_char(ch)) {
                return idx;
            }

            // Skip past this prefix match and keep searching for the full identifier.
            search_start = after_idx;
        }
    }

    panic!("failed to find function `{fn_name}` in {file}");
}

fn end_index_by_line_limit(src: &str, max_lines: usize) -> usize {
    if max_lines == 0 {
        return 0;
    }

    let mut lines_seen = 0usize;
    for (idx, ch) in src.char_indices() {
        if ch == '\n' {
            lines_seen += 1;
            if lines_seen >= max_lines {
                return idx;
            }
        }
    }

    src.len()
}

fn function_scan_window<'a>(src: &'a str, fn_name: &str, file: &str) -> &'a str {
    let start = find_fn_start(src, fn_name, file);
    let rest = &src[start..];

    let mut end = rest.len();

    // Bound the scan to the next command attribute when it appears soon (prevents false positives
    // from later commands in the same file), but always cap by a line limit so functions near the
    // end of a "command block" don't accidentally scan thousands of lines.
    if let Some(idx) = rest.find("#[tauri::command]") {
        end = end.min(idx);
    }
    end = end.min(end_index_by_line_limit(rest, MAX_SCAN_LINES));

    &rest[..end]
}

fn contains_call(body: &str, fn_name: &str) -> bool {
    body.contains(&format!("{fn_name}(")) || body.contains(&format!("{fn_name} ("))
}

fn assert_ipc_origin_checks(src: &str, file: &str, fn_name: &str) {
    let window = function_scan_window(src, fn_name, file);

    let has_combined = contains_call(window, "ensure_main_window_and_trusted_origin")
        || contains_call(window, "ensure_main_window_and_stable_origin");
    let has_main = has_combined || contains_call(window, "ensure_main_window");
    let has_origin = has_combined
        || contains_call(window, "ensure_trusted_origin")
        || contains_call(window, "ensure_stable_origin");

    assert!(
        has_main,
        "{file}:{fn_name} is missing a main-window guard (expected `ipc_origin::ensure_main_window(...)` or `ipc_origin::ensure_main_window_and_stable_origin(...)` / `ipc_origin::ensure_main_window_and_trusted_origin(...)`)"
    );
    assert!(
        has_origin,
        "{file}:{fn_name} is missing an origin guard (expected `ipc_origin::ensure_stable_origin(...)` / `ipc_origin::ensure_trusted_origin(...)` or `ipc_origin::ensure_main_window_and_stable_origin(...)` / `ipc_origin::ensure_main_window_and_trusted_origin(...)`)"
    );
}

#[test]
fn commands_rs_privileged_commands_enforce_origin_guards() {
    let commands_rs = read_repo_file("src/commands.rs");
    let file = "src/commands.rs";

    let privileged_commands = [
        // Filesystem / workbook.
        "open_workbook",
        "save_workbook",
        "read_text_file",
        "read_binary_file",
        "read_binary_file_range",
        "stat_file",
        "list_dir",
        // Clipboard (command wrappers in commands.rs; `clipboard/mod.rs` has its own coverage).
        "read_clipboard",
        "write_clipboard",
        // External integration / network.
        "open_external_url",
        "network_fetch",
        "marketplace_search",
        "marketplace_get_extension",
        "marketplace_download_package",
        // Updater flows (network + privileged restart/exit).
        "check_for_updates",
        "quit_app",
        "restart_app",
        // Macro + Python execution / inspection.
        "get_macro_security_status",
        "set_macro_trust",
        "get_vba_project",
        "list_macros",
        "set_macro_ui_context",
        "validate_vba_migration",
        "fire_workbook_open",
        "fire_workbook_before_close",
        "fire_worksheet_change",
        "fire_selection_change",
        "run_macro",
        "run_python_script",
        // Power Query secret-bearing state.
        "power_query_cache_key_get_or_create",
        "power_query_credential_get",
        "power_query_credential_set",
        "power_query_credential_delete",
        "power_query_credential_list",
        "power_query_refresh_state_get",
        "power_query_refresh_state_set",
        "power_query_state_get",
        "power_query_state_set",
        // SQL queries can reach local databases and must remain origin-scoped.
        "sql_query",
        "sql_get_schema",
    ];

    for cmd in privileged_commands {
        assert_ipc_origin_checks(&commands_rs, file, cmd);
    }
}

#[test]
fn main_rs_privileged_commands_enforce_origin_guards() {
    let main_rs = read_repo_file("src/main.rs");
    let file = "src/main.rs";

    for cmd in ["show_system_notification", "oauth_loopback_listen"] {
        assert_ipc_origin_checks(&main_rs, file, cmd);
    }
}

#[test]
fn clipboard_mod_privileged_commands_enforce_origin_guards() {
    let clipboard_mod = read_repo_file("src/clipboard/mod.rs");
    let file = "src/clipboard/mod.rs";

    for cmd in ["clipboard_read", "clipboard_write"] {
        assert_ipc_origin_checks(&clipboard_mod, file, cmd);
    }
}

#[test]
fn updater_rs_privileged_commands_enforce_origin_guards() {
    let updater_rs = read_repo_file("src/updater.rs");
    let file = "src/updater.rs";

    for cmd in ["install_downloaded_update"] {
        assert_ipc_origin_checks(&updater_rs, file, cmd);
    }
}

#[test]
fn tray_status_rs_commands_enforce_origin_guards() {
    let tray_status_rs = read_repo_file("src/tray_status.rs");
    let file = "src/tray_status.rs";

    for cmd in ["set_tray_status"] {
        assert_ipc_origin_checks(&tray_status_rs, file, cmd);
    }
}

#[test]
fn ed25519_verifier_commands_enforce_origin_guards() {
    let verifier_rs = read_repo_file("src/ed25519_verifier.rs");
    let file = "src/ed25519_verifier.rs";

    for cmd in ["verify_ed25519_signature"] {
        assert_ipc_origin_checks(&verifier_rs, file, cmd);
    }
}
