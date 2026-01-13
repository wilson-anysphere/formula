fn extract_rust_fn(src: &str, name: &str) -> String {
    let patterns = [format!("pub fn {name}"), format!("pub async fn {name}")];
    let start = patterns
        .iter()
        .find_map(|pat| src.find(pat))
        .unwrap_or_else(|| panic!("failed to find function `{name}`"));

    let brace_start = src[start..]
        .find('{')
        .map(|idx| start + idx)
        .unwrap_or_else(|| panic!("failed to find opening brace for `{name}`"));

    let mut depth: i32 = 0;
    let mut end = None;
    for (idx, ch) in src[brace_start..].char_indices() {
        match ch {
            '{' => depth += 1,
            '}' => {
                depth -= 1;
                if depth == 0 {
                    end = Some(brace_start + idx + ch.len_utf8());
                    break;
                }
            }
            _ => {}
        }
    }

    let end = end.unwrap_or_else(|| panic!("failed to find closing brace for `{name}`"));
    src[start..end].to_string()
}

fn assert_ipc_origin_guardrails(src: &str, name: &str) {
    let fun = extract_rust_fn(src, name);
    assert!(
        fun.contains("ensure_main_window("),
        "`{name}` must enforce main-window IPC guardrails via `ensure_main_window(...)`"
    );
    assert!(
        fun.contains("ensure_trusted_origin("),
        "`{name}` must enforce trusted-origin IPC guardrails via `ensure_trusted_origin(...)`"
    );
}

#[test]
fn privileged_lifecycle_commands_enforce_ipc_origin() {
    let src = include_str!("../src/commands.rs");
    for cmd in ["check_for_updates", "quit_app", "restart_app"] {
        assert_ipc_origin_guardrails(src, cmd);
    }
}

#[test]
fn updater_install_command_enforces_ipc_origin() {
    let src = include_str!("../src/updater.rs");
    assert_ipc_origin_guardrails(src, "install_downloaded_update");
}

#[test]
fn tray_status_command_enforces_ipc_origin() {
    let src = include_str!("../src/tray_status.rs");
    assert_ipc_origin_guardrails(src, "set_tray_status");
}
