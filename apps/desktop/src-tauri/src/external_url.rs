use url::Url;

/// Parse + validate a user-provided URL string intended for the `open_external_url` IPC command.
///
/// SECURITY: This is a security boundary between a potentially compromised webview and the host.
/// The host must not become an "open arbitrary protocol" primitive.
///
/// The validation rules are intentionally strict:
/// - Allow: `http`, `https`, `mailto`
/// - Block: `javascript`, `data`, `file`
/// - Reject: all other schemes
pub fn validate_external_url(url: &str) -> Result<Url, String> {
    let trimmed = url.trim();
    let parsed = Url::parse(trimmed).map_err(|err| format!("Invalid URL: {err}"))?;

    match parsed.scheme() {
        "http" | "https" | "mailto" => {
            // Reject userinfo (`https://user:pass@host/...`). This is rarely needed, is deprecated
            // in modern browsers, and can be used to construct misleading URLs (e.g.
            // `https://trusted.com@evil.com/...`).
            if !parsed.username().is_empty() || parsed.password().is_some() {
                return Err("Refusing to open URL containing a username/password".to_string());
            }
            Ok(parsed)
        }
        "javascript" | "data" | "file" => Err(format!(
            "Refusing to open URL with blocked scheme \"{}:\"",
            parsed.scheme()
        )),
        other => Err(format!(
            "Refusing to open URL with unsupported scheme \"{other}:\" (allowed: http, https, mailto)"
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::validate_external_url;

    #[test]
    fn allows_http_https_and_mailto() {
        let https = validate_external_url("https://example.com").expect("https allowed");
        assert_eq!(https.scheme(), "https");
        assert_eq!(https.as_str(), "https://example.com/");

        let http = validate_external_url("http://localhost:1234").expect("http allowed");
        assert_eq!(http.scheme(), "http");
        assert_eq!(http.as_str(), "http://localhost:1234/");

        let mailto = validate_external_url("mailto:test@example.com").expect("mailto allowed");
        assert_eq!(mailto.scheme(), "mailto");
        assert_eq!(mailto.as_str(), "mailto:test@example.com");
    }

    #[test]
    fn rejects_blocked_schemes() {
        let err = validate_external_url("javascript:alert(1)").expect_err("javascript blocked");
        assert_eq!(
            err,
            "Refusing to open URL with blocked scheme \"javascript:\""
        );

        let err = validate_external_url("data:text/plain,hi").expect_err("data blocked");
        assert_eq!(err, "Refusing to open URL with blocked scheme \"data:\"");

        let err = validate_external_url("file:///etc/passwd").expect_err("file blocked");
        assert_eq!(err, "Refusing to open URL with blocked scheme \"file:\"");
    }

    #[test]
    fn rejects_unknown_scheme() {
        let err = validate_external_url("ssh://host").expect_err("unknown scheme rejected");
        assert_eq!(
            err,
            "Refusing to open URL with unsupported scheme \"ssh:\" (allowed: http, https, mailto)"
        );
    }

    #[test]
    fn rejects_urls_with_userinfo() {
        let err = validate_external_url("https://user:pass@example.com")
            .expect_err("userinfo rejected");
        assert_eq!(err, "Refusing to open URL containing a username/password");

        let err =
            validate_external_url("http://user@example.com").expect_err("userinfo rejected");
        assert_eq!(err, "Refusing to open URL containing a username/password");
    }

    #[test]
    fn trims_whitespace() {
        let url =
            validate_external_url(" \n\t https://example.com \n ").expect("whitespace trimmed");
        assert_eq!(url.as_str(), "https://example.com/");
    }
}
