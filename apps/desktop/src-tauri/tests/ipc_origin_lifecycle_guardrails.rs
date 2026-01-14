/// Guardrails to ensure high-privilege lifecycle commands enforce both main-window and
/// trusted-origin checks.
///
/// We use a bounded *text scan* (rather than parsing Rust syntax) so this test remains robust to
/// braces in string literals and runs headlessly without enabling the `desktop` feature.

const MAX_SCAN_LINES: usize = 250;

fn is_ident_char(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_'
}

fn find_fn_start(src: &str, name: &str) -> usize {
    let patterns = [format!("pub async fn {name}"), format!("pub fn {name}")];
    for pat in patterns {
        let mut search_start = 0usize;
        while let Some(rel_idx) = src[search_start..].find(&pat) {
            let idx = search_start + rel_idx;
            let after_idx = idx + pat.len();
            let next = src.get(after_idx..).and_then(|rest| rest.chars().next());
            if next.map_or(true, |ch| !is_ident_char(ch)) {
                return idx;
            }
            search_start = after_idx;
        }
    }
    panic!("failed to find function `{name}`");
}

fn end_index_by_line_limit(src: &str, max_lines: usize) -> usize {
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

fn function_scan_window<'a>(src: &'a str, name: &str) -> &'a str {
    let start = find_fn_start(src, name);
    let rest = &src[start..];

    // Prefer stopping at the next command attribute so we don't accidentally scan into the next
    // function, but always cap by a line limit.
    let mut end = rest.len();
    if let Some(idx) = rest.find("#[tauri::command]") {
        end = end.min(idx);
    }
    end = end.min(end_index_by_line_limit(rest, MAX_SCAN_LINES));
    &rest[..end]
}

fn assert_ipc_origin_guardrails(src: &str, file: &str, name: &str) {
    let fun = function_scan_window(src, name);
    assert!(
        fun.contains("ensure_main_window("),
        "{file}:{name} must enforce main-window IPC guardrails via `ensure_main_window(...)`"
    );
    assert!(
        fun.contains("ensure_trusted_origin("),
        "{file}:{name} must enforce trusted-origin IPC guardrails via `ensure_trusted_origin(...)`"
    );
}

#[test]
fn privileged_lifecycle_commands_enforce_ipc_origin() {
    let file = "../src/commands.rs";
    let src = include_str!("../src/commands.rs");
    for cmd in ["check_for_updates", "quit_app", "restart_app"] {
        assert_ipc_origin_guardrails(src, file, cmd);
    }
}

#[test]
fn updater_install_command_enforces_ipc_origin() {
    let file = "../src/updater.rs";
    let src = include_str!("../src/updater.rs");
    assert_ipc_origin_guardrails(src, file, "install_downloaded_update");
}

#[test]
fn tray_status_command_enforces_ipc_origin() {
    let file = "../src/tray_status.rs";
    let src = include_str!("../src/tray_status.rs");
    assert_ipc_origin_guardrails(src, file, "set_tray_status");
}
