use std::collections::HashSet;

use crate::oauth_redirect_ipc::{
    MAX_PENDING_BYTES as MAX_OAUTH_REDIRECT_PENDING_BYTES,
    MAX_PENDING_URLS as MAX_OAUTH_REDIRECT_PENDING_URLS,
};

/// Extract OAuth redirect URL candidates from a process argv list.
///
/// Some platforms deliver custom scheme deep links (e.g. `formula://...`) via argv on cold start or
/// via the single-instance plugin on warm start. Treat argv as untrusted: it can be arbitrarily
/// large if a malicious sender invokes the app with a huge argument list.
///
/// This helper is intentionally bounded to keep allocations deterministic. When the cap is
/// exceeded, we drop the **oldest** entries and keep the most recent ones ("latest user action
/// wins"). The caps are aligned with the pending oauth-redirect IPC queue enforced by
/// [`crate::oauth_redirect_ipc::OauthRedirectState`].
pub fn extract_oauth_redirect_urls_from_argv(argv: &[String]) -> Vec<String> {
    let mut out_rev = Vec::with_capacity(MAX_OAUTH_REDIRECT_PENDING_URLS.min(argv.len()));
    let mut bytes = 0usize;

    // Walk backwards so we keep the most recent entries, then reverse at the end to preserve the
    // original order among kept URLs.
    for arg in argv.iter().rev() {
        if out_rev.len() >= MAX_OAUTH_REDIRECT_PENDING_URLS {
            break;
        }

        let trimmed = arg.trim().trim_matches('"');
        if trimmed.is_empty() {
            continue;
        }

        // Deep links are delivered via argv as raw URL strings. Filter down to configured app
        // schemes so we don't attempt to parse every argv entry as a URL.
        if !crate::deep_link_schemes::is_deep_link_url(trimmed) {
            continue;
        }

        if trimmed.len() > MAX_OAUTH_REDIRECT_PENDING_BYTES {
            // Single oversized entry; skip rather than exceeding the deterministic cap.
            continue;
        }

        if bytes.saturating_add(trimmed.len()) > MAX_OAUTH_REDIRECT_PENDING_BYTES {
            // Adding this (older) entry would exceed the byte cap; keep scanning in case smaller
            // ones still fit.
            continue;
        }

        bytes += trimmed.len();
        out_rev.push(trimmed.to_string());
    }

    out_rev.reverse();
    out_rev
}

/// Normalize OAuth redirect URLs passed to the desktop host.
///
/// This filters and de-dupes a list of URL strings (argv, deep link plugin, etc) and returns only
/// URLs that are safe to forward to the frontend OAuth broker.
///
/// Accepted:
/// - Custom scheme deep links registered by the desktop app (from
///   `tauri.conf.json` → `plugins.deep-link.desktop.schemes`, e.g. `formula://...`)
/// - RFC 8252 loopback redirects:
///   - `http://127.0.0.1:<port>/...`
///   - `http://localhost:<port>/...`
///   - `http://[::1]:<port>/...`
///
/// SECURITY: loopback URLs are accepted only when the scheme is `http`, an explicit non-zero port
/// is present, and the host is a loopback host.
pub fn normalize_oauth_redirect_request_urls(urls: Vec<String>) -> Vec<String> {
    normalize_oauth_redirect_request_urls_with_schemes(urls, crate::deep_link_schemes::configured_schemes())
}

fn has_scheme_prefix(value: &str, scheme: &str) -> bool {
    let scheme = scheme.trim();
    if scheme.is_empty() {
        return false;
    }

    // Fast, bounded prefix check: we only need to inspect `scheme.len() + 1` bytes (`<scheme>:`).
    let scheme_len = scheme.len();
    if value.len() < scheme_len + 1 {
        return false;
    }
    if value.as_bytes().get(scheme_len) != Some(&b':') {
        return false;
    }

    value
        .get(..scheme_len)
        .map_or(false, |prefix| prefix.eq_ignore_ascii_case(scheme))
}

fn raw_url_has_userinfo(raw: &str) -> bool {
    let Some(after_scheme) = raw.splitn(2, "://").nth(1) else {
        // URLs without an authority (`scheme:...`) cannot contain userinfo.
        return false;
    };
    let authority_end = after_scheme
        .find(&['/', '?', '#'][..])
        .unwrap_or(after_scheme.len());
    let authority = &after_scheme[..authority_end];
    authority.contains('@')
}

fn starts_with_any_scheme(trimmed: &str, schemes: &[String]) -> bool {
    for scheme in schemes {
        if has_scheme_prefix(trimmed, scheme) {
            return true;
        }
    }
    false
}

pub fn normalize_oauth_redirect_request_urls_with_schemes(
    urls: Vec<String>,
    schemes: &[String],
) -> Vec<String> {
    // Keep this bounded; argv / OS-delivered deep links should be treated as untrusted.
    let mut seen =
        HashSet::<String>::with_capacity(MAX_OAUTH_REDIRECT_PENDING_URLS.min(urls.len()));
    let mut out_rev = Vec::with_capacity(MAX_OAUTH_REDIRECT_PENDING_URLS.min(urls.len()));
    let mut bytes = 0usize;

    // Walk backwards so we keep the most recent entries, then reverse at the end to preserve the
    // original order among kept URLs.
    for url in urls.into_iter().rev() {
        let trimmed = url.trim().trim_matches('"');
        if trimmed.is_empty() {
            continue;
        }

        if trimmed.len() > MAX_OAUTH_REDIRECT_PENDING_BYTES {
            // Single oversized entry; skip rather than exceeding the deterministic cap.
            continue;
        }

        if seen.contains(trimmed) {
            continue;
        }

        let is_custom_scheme = starts_with_any_scheme(trimmed, schemes);
        if is_custom_scheme && raw_url_has_userinfo(trimmed) {
            continue;
        }

        let is_loopback = if !is_custom_scheme {
            crate::oauth_loopback::parse_loopback_redirect_uri(trimmed).is_ok()
        } else {
            false
        };

        if !is_custom_scheme && !is_loopback {
            continue;
        }

        if out_rev.len() >= MAX_OAUTH_REDIRECT_PENDING_URLS {
            break;
        }

        if bytes.saturating_add(trimmed.len()) > MAX_OAUTH_REDIRECT_PENDING_BYTES {
            // Adding this (older) entry would exceed the byte cap; keep scanning in case smaller
            // ones still fit.
            continue;
        }

        let normalized = trimmed.to_string();
        bytes += normalized.len();
        seen.insert(normalized.clone());
        out_rev.push(normalized);
    }

    out_rev.reverse();
    out_rev
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_formula_scheme_urls() {
        let out = normalize_oauth_redirect_request_urls(vec![
            "formula://oauth/callback?code=123".to_string(),
            "FORMULA://oauth/callback?code=456".to_string(),
        ]);

        assert_eq!(
            out,
            vec![
                "formula://oauth/callback?code=123".to_string(),
                "FORMULA://oauth/callback?code=456".to_string(),
            ]
        );
    }

    #[test]
    fn accepts_additional_custom_scheme_urls_when_configured() {
        let schemes = vec!["formula".to_string(), "formula-extra".to_string()];
        let out = normalize_oauth_redirect_request_urls_with_schemes(
            vec![
                "formula-extra://oauth/callback?code=123".to_string(),
                "FORMULA-EXTRA://oauth/callback?code=456".to_string(),
            ],
            &schemes,
        );

        assert_eq!(
            out,
            vec![
                "formula-extra://oauth/callback?code=123".to_string(),
                "FORMULA-EXTRA://oauth/callback?code=456".to_string(),
            ]
        );
    }

    #[test]
    fn accepts_rfc8252_loopback_redirects_for_ipv4_ipv6_and_localhost() {
        let out = normalize_oauth_redirect_request_urls(vec![
            "http://127.0.0.1:8080/callback?code=1".to_string(),
            "http://localhost:8081/callback?code=2".to_string(),
            "http://[::1]:8082/callback?code=3".to_string(),
        ]);

        assert_eq!(
            out,
            vec![
                "http://127.0.0.1:8080/callback?code=1".to_string(),
                "http://localhost:8081/callback?code=2".to_string(),
                "http://[::1]:8082/callback?code=3".to_string(),
            ]
        );
    }

    #[test]
    fn rejects_non_loopback_or_insecure_redirects() {
        let out = normalize_oauth_redirect_request_urls(vec![
            // Wrong scheme.
            "https://127.0.0.1:8080/callback?code=1".to_string(),
            // Missing port.
            "http://127.0.0.1/callback?code=2".to_string(),
            // Explicit port 0 is forbidden.
            "http://127.0.0.1:0/callback?code=3".to_string(),
            // Not a loopback host.
            "http://example.com:8080/callback?code=4".to_string(),
            // Not an allowed loopback IP.
            "http://127.0.0.2:8080/callback?code=5".to_string(),
            // Not an allowed loopback IP.
            "http://[::2]:8080/callback?code=6".to_string(),
        ]);

        assert!(out.is_empty());
    }

    #[test]
    fn trims_quotes_de_dupes_and_drops_empty_entries() {
        let out = normalize_oauth_redirect_request_urls(vec![
            "  ".to_string(),
            "\"formula://oauth/callback?code=123\"".to_string(),
            "formula://oauth/callback?code=123".to_string(),
        ]);

        assert_eq!(out, vec!["formula://oauth/callback?code=123".to_string()]);
    }

    #[test]
    fn rejects_custom_scheme_urls_with_userinfo() {
        let out = normalize_oauth_redirect_request_urls(vec![
            "formula://user@oauth/callback?code=123".to_string(),
            "FORMULA://user:pass@oauth/callback?code=456".to_string(),
            "formula://oauth/callback?code=789".to_string(),
        ]);

        assert_eq!(out, vec!["formula://oauth/callback?code=789".to_string()]);
    }

    #[test]
    fn caps_oauth_redirect_urls_by_count_dropping_oldest() {
        let urls: Vec<String> = (0..(MAX_OAUTH_REDIRECT_PENDING_URLS + 3))
            .map(|idx| format!("formula://u{idx}"))
            .collect();

        let out = normalize_oauth_redirect_request_urls(urls);
        assert_eq!(out.len(), MAX_OAUTH_REDIRECT_PENDING_URLS);

        let expected: Vec<String> = (3..(MAX_OAUTH_REDIRECT_PENDING_URLS + 3))
            .map(|idx| format!("formula://u{idx}"))
            .collect();
        assert_eq!(out, expected);
    }

    #[test]
    fn caps_oauth_redirect_urls_by_total_bytes_dropping_oldest_deterministically() {
        // Use fixed-size strings so the expected trim point is deterministic.
        let entry_len = 4096;
        let prefix_len = "formula://000-".len();
        let payload = "x".repeat(entry_len - prefix_len);
        let urls: Vec<String> = (0..MAX_OAUTH_REDIRECT_PENDING_URLS)
            .map(|i| format!("formula://{i:03}-{payload}"))
            .collect();

        assert_eq!(urls.len(), MAX_OAUTH_REDIRECT_PENDING_URLS);
        assert_eq!(urls[0].len(), entry_len);

        let out = normalize_oauth_redirect_request_urls(urls.clone());

        let expected_len = MAX_OAUTH_REDIRECT_PENDING_BYTES / entry_len;
        assert_eq!(out.len(), expected_len);

        let total_bytes: usize = out.iter().map(|u| u.len()).sum();
        assert!(
            total_bytes <= MAX_OAUTH_REDIRECT_PENDING_BYTES,
            "normalized bytes {total_bytes} exceeded cap {MAX_OAUTH_REDIRECT_PENDING_BYTES}"
        );

        let expected = urls[MAX_OAUTH_REDIRECT_PENDING_URLS - expected_len..].to_vec();
        assert_eq!(out, expected);
    }

    #[test]
    fn extract_oauth_redirect_urls_from_argv_caps_by_count_dropping_oldest() {
        let mut argv = vec!["formula-desktop".to_string()];
        for idx in 0..(MAX_OAUTH_REDIRECT_PENDING_URLS + 3) {
            argv.push(format!("formula://u{idx}"));
        }

        let out = extract_oauth_redirect_urls_from_argv(&argv);
        assert_eq!(out.len(), MAX_OAUTH_REDIRECT_PENDING_URLS);

        let expected: Vec<String> = (3..(MAX_OAUTH_REDIRECT_PENDING_URLS + 3))
            .map(|idx| format!("formula://u{idx}"))
            .collect();
        assert_eq!(out, expected);
    }

    #[test]
    fn extract_oauth_redirect_urls_from_argv_caps_by_total_bytes_dropping_oldest_deterministically() {
        // Use fixed-size strings so the expected trim point is deterministic.
        let entry_len = 4096;
        let prefix_len = "formula://000-".len();
        let payload = "x".repeat(entry_len - prefix_len);

        let mut argv = vec!["formula-desktop".to_string()];
        for i in 0..MAX_OAUTH_REDIRECT_PENDING_URLS {
            argv.push(format!("formula://{i:03}-{payload}"));
        }

        let out = extract_oauth_redirect_urls_from_argv(&argv);
        let expected_len = MAX_OAUTH_REDIRECT_PENDING_BYTES / entry_len;
        assert_eq!(out.len(), expected_len);

        let total_bytes: usize = out.iter().map(|u| u.len()).sum();
        assert!(
            total_bytes <= MAX_OAUTH_REDIRECT_PENDING_BYTES,
            "extracted bytes {total_bytes} exceeded cap {MAX_OAUTH_REDIRECT_PENDING_BYTES}"
        );

        let expected = argv[argv.len() - expected_len..].to_vec();
        assert_eq!(out, expected);
    }
}
