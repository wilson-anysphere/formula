//! Guardrails for the OAuth redirect IPC readiness handshake.
//!
//! Tauri does not guarantee that emitted events are queued before JS listeners are registered.
//! To avoid dropping `oauth-redirect` deep links on cold start, the Rust backend queues incoming
//! redirect URLs until the frontend emits `oauth-redirect-ready`, then flushes the queued URLs.
//!
//! `src/main.rs` is only compiled when the `desktop` feature is enabled (it depends on the
//! system WebView toolchain on Linux). These are intentionally source-level tests so they run in
//! headless CI without that feature.

fn slice_between<'a>(
    src: &'a str,
    start_marker: &str,
    end_marker: &str,
    context: &str,
) -> &'a str {
    let start = src
        .find(start_marker)
        .unwrap_or_else(|| panic!("{context}: failed to find start marker `{start_marker}`"));
    let end = src[start..]
        .find(end_marker)
        .map(|idx| start + idx)
        .unwrap_or_else(|| panic!("{context}: failed to find end marker `{end_marker}`"));
    &src[start..end]
}

#[test]
fn tauri_main_wires_oauth_redirect_ready_handshake() {
    let main_rs = include_str!(concat!(env!("CARGO_MANIFEST_DIR"), "/src/main.rs"));

    // Ensure cold-start argv OAuth redirects are queued too.
    let init_start = main_rs.find("let initial_oauth_urls").expect(
        "desktop main.rs missing initial_oauth_urls extraction; \
         cold-start OAuth redirects must be queued until the frontend is ready",
    );
    let init_window = main_rs
        .get(init_start..init_start.saturating_add(700))
        .unwrap_or(&main_rs[init_start..]);
    assert!(
        init_window.contains(".queue_or_emit("),
        "desktop main.rs must queue initial argv OAuth redirects via OauthRedirectState::queue_or_emit \
         so redirects aren't dropped on cold start"
    );

    // --- 1) Ensure there is an `oauth-redirect-ready` listener registered.
    let listener_start = main_rs
        .find("listen(OAUTH_REDIRECT_READY_EVENT")
        .expect(
            "desktop main.rs must listen for OAUTH_REDIRECT_READY_EVENT (oauth-redirect-ready) so \
             queued OAuth redirect deep links aren't dropped on cold start",
        );
    let listener_after = &main_rs[listener_start..];
    let listener_end = listener_after
        .find("});")
        .map(|idx| idx + 3)
        .expect(
            "failed to locate end of OAUTH_REDIRECT_READY_EVENT listener (expected `});` after window.listen(...))",
        );
    let listener_body = &listener_after[..listener_end];

    // --- 2) Ensure the listener flips readiness exactly once, drains pending URLs, and emits.
    //
    // With the dedicated `OauthRedirectState` state machine, idempotent flush-once behavior is
    // encapsulated in `mark_ready_and_drain`.
    assert!(
        listener_body.contains(".mark_ready_and_drain("),
        "OAUTH_REDIRECT_READY_EVENT listener must call OauthRedirectState::mark_ready_and_drain \
         so queued OAuth redirect URLs are flushed exactly once"
    );
    let ready_calls_in_listener = listener_body.matches(".mark_ready_and_drain(").count();
    assert_eq!(
        ready_calls_in_listener, 1,
        "expected exactly one mark_ready_and_drain call inside OAUTH_REDIRECT_READY_EVENT listener, found {ready_calls_in_listener}"
    );

    let emits_oauth_redirect = listener_body.contains("emit_oauth_redirect_event")
        || listener_body.contains(".emit(OAUTH_REDIRECT_EVENT")
        || listener_body.contains(".emit(\"oauth-redirect\"");
    assert!(
        emits_oauth_redirect,
        "OAUTH_REDIRECT_READY_EVENT listener must flush queued URLs by emitting `oauth-redirect` \
         events. Without this, OAuth redirects delivered before JS installs listeners will be \
         dropped on cold start."
    );

    let iterates_pending = listener_body.contains("for url in pending")
        || listener_body.contains("for url in pending_urls")
        || listener_body.contains(".for_each(");
    assert!(
        iterates_pending,
        "OAUTH_REDIRECT_READY_EVENT listener must emit an `oauth-redirect` event for *each* drained \
         URL (e.g. loop over the drained list). Emitting only one URL would drop additional \
         redirects."
    );

    // --- 3) Ensure `handle_oauth_redirect_request` queues when not ready.
    let handler_body = slice_between(
        main_rs,
        "fn handle_oauth_redirect_request",
        "fn extract_open_file_paths",
        "oauth redirect handler wiring",
    );

    assert!(
        handler_body.contains(".queue_or_emit("),
        "handle_oauth_redirect_request must route incoming URLs through OauthRedirectState::queue_or_emit \
         so redirects are either queued (before ready) or emitted immediately (after ready)"
    );
}
