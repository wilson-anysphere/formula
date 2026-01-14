#[test]
fn updater_ui_ready_listener_requires_trusted_origin() {
    // This is a source-level guardrail: `src/main.rs` is only compiled with the `desktop` feature,
    // but we still want CI to catch regressions when running headless backend tests.
    let main_rs_src = include_str!("../src/main.rs");

    let start_marker = r#"listen("updater-ui-ready""#;
    let start = main_rs_src
        .find(start_marker)
        .unwrap_or_else(|| panic!("failed to find `{start_marker}` in src/main.rs"));

    let rest = &main_rs_src[start..];
    let spawn_marker = "updater::spawn_update_check";
    let spawn = rest
        .find(spawn_marker)
        .unwrap_or_else(|| panic!("failed to find `{spawn_marker}` after `{start_marker}`"));

    let listener_block = &rest[..spawn];
    assert!(
        listener_block.contains("ensure_stable_origin"),
        "`updater-ui-ready` listener must gate to trusted app origins via `ipc_origin::ensure_stable_origin`"
    );
}
