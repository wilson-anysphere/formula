//! Guardrails for sensitive event delivery to the WebView.
//!
//! These are source-level tests so they run headless without the `desktop` feature (which would
//! pull in the system WebView toolchain on Linux).

fn extract_brace_block<'a>(source: &'a str, anchor: &str) -> &'a str {
    let start = source
        .find(anchor)
        .unwrap_or_else(|| panic!("main.rs missing anchor {anchor:?}"));

    let brace_start = source[start..]
        .find('{')
        .map(|idx| start + idx)
        .unwrap_or_else(|| panic!("main.rs missing '{{' after anchor {anchor:?}"));

    let bytes = source.as_bytes();
    let mut depth: i32 = 0;
    for i in brace_start..bytes.len() {
        match bytes[i] {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    return &source[brace_start..=i];
                }
            }
            _ => {}
        }
    }

    panic!("main.rs contains unbalanced braces after anchor {anchor:?}");
}

fn assert_any_contains_in_order(haystack: &str, guards: &[&str], second: &str, context: &str) {
    let second_idx = haystack
        .find(second)
        .unwrap_or_else(|| panic!("{context} (missing {second:?})"));

    let (guard, guard_idx) = guards
        .iter()
        .filter_map(|g| haystack.find(g).map(|idx| (*g, idx)))
        .min_by_key(|(_, idx)| *idx)
        .unwrap_or_else(|| panic!("{context} (missing any of {guards:?})"));

    assert!(
        guard_idx < second_idx,
        "{context} (expected {guard:?} before {second:?})"
    );
}

#[test]
fn sensitive_ipc_events_require_trusted_origin() {
    let main_rs = include_str!("../src/main.rs");

    let emit_open_file_event = extract_brace_block(main_rs, "fn emit_open_file_event");
    assert_any_contains_in_order(
        emit_open_file_event,
        &["ensure_stable_origin", "is_trusted_app_origin"],
        "window.emit(OPEN_FILE_EVENT",
        "emit_open_file_event must verify the main window origin is trusted before emitting",
    );
    assert!(
        !emit_open_file_event.contains("app.emit(OPEN_FILE_EVENT"),
        "emit_open_file_event must not broadcast OPEN_FILE_EVENT via AppHandle::emit"
    );

    let emit_oauth_redirect_event = extract_brace_block(main_rs, "fn emit_oauth_redirect_event");
    assert_any_contains_in_order(
        emit_oauth_redirect_event,
        &["ensure_stable_origin", "is_trusted_app_origin"],
        "window.emit(OAUTH_REDIRECT_EVENT",
        "emit_oauth_redirect_event must verify the main window origin is trusted before emitting",
    );
    assert!(
        !emit_oauth_redirect_event.contains("app.emit(OAUTH_REDIRECT_EVENT"),
        "emit_oauth_redirect_event must not broadcast OAUTH_REDIRECT_EVENT via AppHandle::emit"
    );

    let open_file_ready_listener = extract_brace_block(main_rs, "listen(OPEN_FILE_READY_EVENT");
    assert_any_contains_in_order(
        open_file_ready_listener,
        &["ensure_stable_origin", "is_trusted_app_origin"],
        "mark_ready_and_drain",
        "OPEN_FILE_READY_EVENT handler must verify the main window origin is trusted before draining",
    );

    let oauth_redirect_ready_listener =
        extract_brace_block(main_rs, "listen(OAUTH_REDIRECT_READY_EVENT");
    assert_any_contains_in_order(
        oauth_redirect_ready_listener,
        &["ensure_stable_origin", "is_trusted_app_origin"],
        "mark_ready_and_drain",
        "OAUTH_REDIRECT_READY_EVENT handler must verify the main window origin is trusted before draining",
    );
}
