//! Guardrails for the RFC 8252 loopback OAuth listener wiring.
//!
//! These are source-level tests so they run headless without the `desktop` feature (which would
//! pull in the system WebView toolchain on Linux).

const MAX_SCAN_LINES: usize = 400;

fn is_ident_char(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_'
}

fn find_fn_start(src: &str, fn_name: &str) -> usize {
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
            if next.map_or(true, |ch| !is_ident_char(ch)) {
                return idx;
            }

            search_start = after_idx;
        }
    }

    panic!("failed to find function `{fn_name}` in src/main.rs");
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

fn function_scan_window<'a>(src: &'a str, fn_name: &str) -> &'a str {
    let start = find_fn_start(src, fn_name);
    let rest = &src[start..];
    let end = end_index_by_line_limit(rest, MAX_SCAN_LINES);
    &rest[..end]
}

fn slice_after<'a>(src: &'a str, marker: &str, len: usize, context: &str) -> &'a str {
    let start = src
        .find(marker)
        .unwrap_or_else(|| panic!("{context}: missing marker {marker:?}"));
    src.get(start..start.saturating_add(len))
        .unwrap_or(&src[start..])
}

#[test]
fn oauth_loopback_listen_is_centralized_capped_and_dual_stack_for_localhost() {
    let main_rs = include_str!("../src/main.rs");
    let window = function_scan_window(main_rs, "oauth_loopback_listen");

    assert!(
        window.contains("parse_loopback_redirect_uri"),
        "oauth_loopback_listen must validate redirect_uri via parse_loopback_redirect_uri so host/scheme/port checks stay centralized"
    );

    assert!(
        window.contains("acquire_oauth_loopback_listener"),
        "oauth_loopback_listen must call acquire_oauth_loopback_listener so the active-listener cap is enforced"
    );

    // `localhost` can resolve to either 127.0.0.1 or ::1. The listener should bind both address
    // families (best-effort) to avoid platform resolver differences breaking OAuth.
    assert!(
        window.contains("Ipv4Addr::LOCALHOST"),
        "oauth_loopback_listen must attempt to bind Ipv4Addr::LOCALHOST when handling loopback redirect URIs"
    );
    assert!(
        window.contains("Ipv6Addr::LOCALHOST"),
        "oauth_loopback_listen must attempt to bind Ipv6Addr::LOCALHOST when handling loopback redirect URIs"
    );

    let wants_ipv4 = slice_after(
        window,
        "let wants_ipv4",
        300,
        "oauth_loopback_listen wants_ipv4 wiring",
    );
    assert!(
        wants_ipv4.contains("LoopbackHostKind::Localhost"),
        "oauth_loopback_listen wants_ipv4 must include LoopbackHostKind::Localhost so localhost redirect URIs bind IPv4"
    );

    let wants_ipv6 = slice_after(
        window,
        "let wants_ipv6",
        300,
        "oauth_loopback_listen wants_ipv6 wiring",
    );
    assert!(
        wants_ipv6.contains("LoopbackHostKind::Localhost"),
        "oauth_loopback_listen wants_ipv6 must include LoopbackHostKind::Localhost so localhost redirect URIs bind IPv6"
    );
}

