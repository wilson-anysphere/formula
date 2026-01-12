use url::Url;

/// Returns whether a WebView origin is considered trusted for privileged IPC commands.
///
/// This is a defense-in-depth guard: even though Tauri's security model should prevent remote
/// origins from accessing the invoke API by default, we keep privileged commands resilient in case
/// remote content is ever loaded into a WebView.
///
/// We treat app-local content as trusted:
/// - packaged builds typically run on an internal `*.localhost` origin
/// - dev builds run on `http://localhost:<port>`
///
/// Loopback variants (`127.0.0.1`, `::1`) are also allowed, and `file://` is included as a
/// best-effort compatibility fallback.
pub fn is_trusted_app_origin(url: &Url) -> bool {
    match url.scheme() {
        "data" => return false,
        // Best-effort compatibility fallback.
        "file" => return true,
        _ => {}
    }

    match url.host() {
        Some(url::Host::Domain(host)) => host == "localhost" || host.ends_with(".localhost"),
        Some(url::Host::Ipv4(ip)) => ip == std::net::Ipv4Addr::LOCALHOST,
        Some(url::Host::Ipv6(ip)) => ip == std::net::Ipv6Addr::LOCALHOST,
        None => false,
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Verb {
    Is,
    Are,
}

impl Verb {
    fn as_str(self) -> &'static str {
        match self {
            Verb::Is => "is",
            Verb::Are => "are",
        }
    }
}

pub fn ensure_main_window(label: &str, subject: &str, verb: Verb) -> Result<(), String> {
    if label != "main" {
        return Err(format!(
            "{subject} {} only allowed from the main window",
            verb.as_str()
        ));
    }
    Ok(())
}

pub fn ensure_trusted_origin(url: &Url, subject: &str, verb: Verb) -> Result<(), String> {
    if !is_trusted_app_origin(url) {
        return Err(format!(
            "{subject} {} not allowed from this origin",
            verb.as_str()
        ));
    }
    Ok(())
}

pub fn ensure_main_window_and_trusted_origin(
    label: &str,
    url: &Url,
    subject: &str,
    verb: Verb,
) -> Result<(), String> {
    ensure_main_window(label, subject, verb)?;
    ensure_trusted_origin(url, subject, verb)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::is_trusted_app_origin;
    use url::Url;

    #[test]
    fn allows_localhost() {
        let url = Url::parse("http://localhost:1420/").unwrap();
        assert!(is_trusted_app_origin(&url));
    }

    #[test]
    fn allows_localhost_subdomains() {
        let url = Url::parse("https://tauri.localhost/").unwrap();
        assert!(is_trusted_app_origin(&url));

        let url = Url::parse("https://foo.bar.localhost/some/path").unwrap();
        assert!(is_trusted_app_origin(&url));
    }

    #[test]
    fn allows_loopback_hosts() {
        let url = Url::parse("http://127.0.0.1:1234/").unwrap();
        assert!(is_trusted_app_origin(&url));

        let url = Url::parse("http://[::1]:1234/").unwrap();
        assert!(is_trusted_app_origin(&url));
    }

    #[test]
    fn allows_file_scheme_fallback() {
        let url = Url::parse("file:///foo/bar.txt").unwrap();
        assert!(is_trusted_app_origin(&url));
    }

    #[test]
    fn denies_data_scheme() {
        let url = Url::parse("data:text/plain,hello").unwrap();
        assert!(!is_trusted_app_origin(&url));
    }

    #[test]
    fn denies_remote_hosts() {
        let url = Url::parse("https://example.com/").unwrap();
        assert!(!is_trusted_app_origin(&url));
    }

    #[test]
    fn denies_empty_host() {
        // Valid URL with a scheme but no host.
        let url = Url::parse("http:///path").unwrap();
        assert!(!is_trusted_app_origin(&url));
    }
}
