//! Guardrails for the OAuth redirect IPC readiness handshake.
//!
//! Tauri does not guarantee that emitted events are queued before JS listeners are registered.
//! To avoid dropping `oauth-redirect` deep links on cold start, the Rust backend queues incoming
//! redirect URLs until the frontend emits `oauth-redirect-ready`, then flushes the queued URLs.
//!
//! `src/main.rs` is only compiled when the `desktop` feature is enabled (it depends on the
//! system WebView toolchain on Linux). These are intentionally source-level tests so they run in
//! headless CI without that feature.

fn extract_brace_block<'a>(source: &'a str, anchor: &str, context: &str) -> &'a str {
    let start = source
        .find(anchor)
        .unwrap_or_else(|| panic!("{context}: missing anchor {anchor:?}"));

    let brace_start = source[start..]
        .find('{')
        .map(|idx| start + idx)
        .unwrap_or_else(|| panic!("{context}: missing '{{' after anchor {anchor:?}"));

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

    panic!("{context}: unbalanced braces after anchor {anchor:?}");
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
    let listener_body = extract_brace_block(
        main_rs,
        "listen(OAUTH_REDIRECT_READY_EVENT",
        "oauth redirect ready listener wiring",
    );

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

    // Extra guardrail: the backend should only flip *oauth-redirect* readiness in response to the
    // frontend readiness signal. If `mark_ready_and_drain` starts getting called elsewhere (e.g.
    // during startup), cold-start redirects can be emitted before the JS listener exists.
    let mut oauth_redirect_ready_calls = 0;
    for (idx, _) in main_rs.match_indices("state::<SharedOauthRedirectState>") {
        let window = main_rs
            // Keep the scan bounded but allow some room for origin checks / logging between the
            // state lookup and the eventual flush call.
            .get(idx..idx.saturating_add(800))
            .unwrap_or(&main_rs[idx..]);
        oauth_redirect_ready_calls += window.matches(".mark_ready_and_drain(").count();
    }
    assert_eq!(
        oauth_redirect_ready_calls, 1,
        "expected exactly one mark_ready_and_drain call associated with SharedOauthRedirectState in desktop main.rs (readiness should only flip in response to oauth-redirect-ready), found {oauth_redirect_ready_calls}"
    );

    // Guardrail: the ready listener must verify the webview origin before flushing pending URLs.
    let mark_ready_idx = listener_body.find(".mark_ready_and_drain(").expect(
        "OAUTH_REDIRECT_READY_EVENT listener must call mark_ready_and_drain to flush queued URLs",
    );
    let state_idx = listener_body.find("state::<SharedOauthRedirectState>").expect(
        "OAUTH_REDIRECT_READY_EVENT listener must reference SharedOauthRedirectState before flushing queued URLs",
    );
    assert!(
        state_idx < mark_ready_idx,
        "expected SharedOauthRedirectState lookup to occur before mark_ready_and_drain inside OAUTH_REDIRECT_READY_EVENT listener"
    );

    let trusted_origin_idx = listener_body.find("is_trusted_app_origin").expect(
        "OAUTH_REDIRECT_READY_EVENT listener must reject untrusted origins before flushing queued URLs",
    );
    assert!(
        trusted_origin_idx < mark_ready_idx,
        "expected trusted origin check to occur before mark_ready_and_drain inside OAUTH_REDIRECT_READY_EVENT listener"
    );

    let stable_origin_idx = listener_body.find("ensure_stable_origin").expect(
        "OAUTH_REDIRECT_READY_EVENT listener must enforce stable origin before flushing queued URLs",
    );
    assert!(
        stable_origin_idx < mark_ready_idx,
        "expected stable origin check to occur before mark_ready_and_drain inside OAUTH_REDIRECT_READY_EVENT listener"
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
    let handler_body = extract_brace_block(
        main_rs,
        "fn handle_oauth_redirect_request",
        "oauth redirect handler wiring",
    );

    assert!(
        handler_body.contains(".queue_or_emit("),
        "handle_oauth_redirect_request must route incoming URLs through OauthRedirectState::queue_or_emit \
         so redirects are either queued (before ready) or emitted immediately (after ready)"
    );
}
