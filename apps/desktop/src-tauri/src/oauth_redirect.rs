use std::collections::HashSet;

/// Normalize OAuth redirect URLs passed to the desktop host.
///
/// This filters and de-dupes a list of URL strings (argv, deep link plugin, etc) and returns only
/// URLs that are safe to forward to the frontend OAuth broker.
///
/// Accepted:
/// - Custom scheme deep links: `formula://...`
/// - RFC 8252 loopback redirects:
///   - `http://127.0.0.1:<port>/...`
///   - `http://localhost:<port>/...`
///   - `http://[::1]:<port>/...`
///
/// SECURITY: loopback URLs are accepted only when the scheme is `http`, an explicit non-zero port
/// is present, and the host is a loopback host.
pub fn normalize_oauth_redirect_request_urls(urls: Vec<String>) -> Vec<String> {
    let mut seen = HashSet::<String>::new();
    let mut out = Vec::new();

    for url in urls {
        let trimmed = url.trim().trim_matches('"');
        if trimmed.is_empty() {
            continue;
        }

        let is_formula = trimmed
            .get(..8)
            .map_or(false, |prefix| prefix.eq_ignore_ascii_case("formula:"));

        let is_loopback = if !is_formula {
            crate::oauth_loopback::parse_loopback_redirect_uri(trimmed).is_ok()
        } else {
            false
        };

        if !is_formula && !is_loopback {
            continue;
        }

        let normalized = trimmed.to_string();
        if seen.insert(normalized.clone()) {
            out.push(normalized);
        }
    }

    out
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
}
