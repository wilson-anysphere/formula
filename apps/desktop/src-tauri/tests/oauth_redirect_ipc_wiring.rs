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
    let ready_assigns_in_listener = listener_body.matches(".ready = true").count();
    assert_eq!(
        ready_assigns_in_listener, 1,
        "OAUTH_REDIRECT_READY_EVENT listener must set OauthRedirectState.ready = true exactly once \
         (idempotent flush). Without this guard, the frontend can trigger multiple flushes and \
         re-emit stale OAuth redirects."
    );

    // Extra guardrail (mirrors the open-file test): readiness should only be flipped in response
    // to the explicit frontend readiness event.
    let total_ready_assigns = main_rs.matches(".ready = true").count();
    assert_eq!(
        total_ready_assigns, 1,
        "expected exactly one `.ready = true` assignment in desktop main.rs (inside the \
         OAUTH_REDIRECT_READY_EVENT listener). If readiness is flipped elsewhere, cold-start OAuth \
         redirects can be emitted before the JS listener exists."
    );

    assert!(
        listener_body.contains("if guard.ready") || listener_body.contains("if state.ready"),
        "OAUTH_REDIRECT_READY_EVENT listener should return early if already ready so the flush \
         happens at most once"
    );

    let drains_pending = listener_body.contains("take(&mut guard.pending_urls)")
        || listener_body.contains("take(&mut state.pending_urls)")
        || listener_body.contains("pending_urls.drain(");
    assert!(
        drains_pending,
        "OAUTH_REDIRECT_READY_EVENT listener must drain queued OAuth redirect URLs \
         (e.g. `std::mem::take(&mut guard.pending_urls)`). Without draining, redirects captured \
         before the frontend is ready may be lost or re-emitted unexpectedly."
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

    let has_ready_branch = handler_body.contains("if state.ready") || handler_body.contains("if !state.ready");
    assert!(
        has_ready_branch,
        "handle_oauth_redirect_request must branch on OauthRedirectState.ready so it can either \
         emit immediately (ready) or queue (not ready). Without this, OAuth redirects can be \
         emitted before the JS listener exists and get dropped."
    );

    assert!(
        handler_body.contains("pending_urls.extend("),
        "handle_oauth_redirect_request must queue incoming URLs when the frontend isn't ready yet \
         (e.g. `state.pending_urls.extend(urls)`). If this queueing is removed, OAuth redirects \
         arriving during startup will be lost."
    );

    // Heuristic sanity check: in the current implementation the `extend` call should occur after
    // the readiness check (in the not-ready branch). Keep this check lax so refactors that keep
    // the semantics intact can still pass.
    if let (Some(ready_idx), Some(extend_idx)) =
        (handler_body.find("state.ready"), handler_body.find("pending_urls.extend"))
    {
        assert!(
            extend_idx > ready_idx,
            "expected `pending_urls.extend(...)` to appear after the readiness check in \
             handle_oauth_redirect_request (queueing should be conditional on not-ready)"
        );
    }
}

