use std::collections::HashSet;
use std::net::{IpAddr, Ipv4Addr, Ipv6Addr, SocketAddr};

use url::Url;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LoopbackHost {
    Ipv4,
    Ipv6,
    Localhost,
}

#[derive(Debug, thiserror::Error, PartialEq, Eq)]
pub enum LoopbackRedirectUriError {
    #[error("Invalid OAuth redirect URI {uri:?}: {source}")]
    Parse {
        uri: String,
        #[source]
        source: url::ParseError,
    },

    #[error("Invalid loopback OAuth redirect URI {uri:?}: scheme must be http:// (got {scheme:?})")]
    UnsupportedScheme { uri: String, scheme: String },

    #[error(
        "Invalid loopback OAuth redirect URI {uri:?}: host must be 127.0.0.1, localhost, or ::1 (got {host:?})"
    )]
    UnsupportedHost { uri: String, host: String },

    #[error("Invalid loopback OAuth redirect URI {uri:?}: must include an explicit port")]
    MissingPort { uri: String },

    #[error("Invalid loopback OAuth redirect URI {uri:?}: port must not be 0")]
    ZeroPort { uri: String },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LoopbackRedirectUri {
    pub parsed: Url,
    pub host: LoopbackHost,
    pub port: u16,
    pub expected_path: String,
    /// Loopback addresses the host should bind to in order to capture redirects.
    ///
    /// - `127.0.0.1` -> [`Ipv4Addr::LOCALHOST`]
    /// - `::1` -> [`Ipv6Addr::LOCALHOST`]
    /// - `localhost` -> *both* IPv4 and IPv6 loopback (best-effort)
    pub bind_addrs: Vec<SocketAddr>,
}

impl LoopbackRedirectUri {
    pub fn parse(raw: &str) -> Result<Self, LoopbackRedirectUriError> {
        let raw = raw.trim();
        if raw.is_empty() {
            return Err(LoopbackRedirectUriError::Parse {
                uri: raw.to_string(),
                source: url::ParseError::EmptyHost,
            });
        }

        let parsed = Url::parse(raw).map_err(|source| LoopbackRedirectUriError::Parse {
            uri: raw.to_string(),
            source,
        })?;

        if parsed.scheme() != "http" {
            return Err(LoopbackRedirectUriError::UnsupportedScheme {
                uri: raw.to_string(),
                scheme: parsed.scheme().to_string(),
            });
        }

        let host = match parsed.host() {
            Some(url::Host::Domain(host)) if host == "localhost" => LoopbackHost::Localhost,
            Some(url::Host::Ipv4(ip)) if ip == Ipv4Addr::LOCALHOST => LoopbackHost::Ipv4,
            Some(url::Host::Ipv6(ip)) if ip == Ipv6Addr::LOCALHOST => LoopbackHost::Ipv6,
            Some(other) => {
                return Err(LoopbackRedirectUriError::UnsupportedHost {
                    uri: raw.to_string(),
                    host: other.to_string(),
                })
            }
            None => {
                return Err(LoopbackRedirectUriError::UnsupportedHost {
                    uri: raw.to_string(),
                    host: "".to_string(),
                })
            }
        };

        let port = parsed
            .port()
            .ok_or_else(|| LoopbackRedirectUriError::MissingPort {
                uri: raw.to_string(),
            })?;
        if port == 0 {
            return Err(LoopbackRedirectUriError::ZeroPort {
                uri: raw.to_string(),
            });
        }

        let expected_path = parsed.path().to_string();
        let bind_addrs = match host {
            LoopbackHost::Ipv4 => vec![SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), port)],
            LoopbackHost::Ipv6 => vec![SocketAddr::new(IpAddr::V6(Ipv6Addr::LOCALHOST), port)],
            LoopbackHost::Localhost => vec![
                SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), port),
                SocketAddr::new(IpAddr::V6(Ipv6Addr::LOCALHOST), port),
            ],
        };

        Ok(Self {
            parsed,
            host,
            port,
            expected_path,
            bind_addrs,
        })
    }
}

pub fn is_loopback_http_redirect_uri(url: &Url) -> bool {
    if url.scheme() != "http" {
        return false;
    }
    let Some(port) = url.port() else {
        return false;
    };
    if port == 0 {
        return false;
    }
    match url.host() {
        Some(url::Host::Domain(host)) => host == "localhost",
        Some(url::Host::Ipv4(ip)) => ip == Ipv4Addr::LOCALHOST,
        Some(url::Host::Ipv6(ip)) => ip == Ipv6Addr::LOCALHOST,
        None => false,
    }
}

/// Normalize OAuth redirect URLs passed to the desktop host.
///
/// This function filters and de-dupes a list of URL strings (argv, deep link plugin, etc) and
/// returns only URLs that are safe to forward to the frontend OAuth broker:
///
/// - Custom scheme deep links: `formula://...`
/// - RFC 8252 loopback redirects:
///   - `http://127.0.0.1:<port>/...`
///   - `http://localhost:<port>/...`
///   - `http://[::1]:<port>/...`
///
/// Security note: loopback URLs are accepted only when the scheme is `http` and an explicit,
/// non-zero port is present.
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
            Url::parse(trimmed)
                .ok()
                .is_some_and(|url| is_loopback_http_redirect_uri(&url))
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
    fn normalize_accepts_formula_urls_unchanged() {
        let urls = vec![
            "formula://oauth/callback?code=123".to_string(),
            "\"formula://oauth/callback?code=123\"".to_string(),
        ];
        let normalized = normalize_oauth_redirect_request_urls(urls);
        assert_eq!(
            normalized,
            vec!["formula://oauth/callback?code=123".to_string()]
        );
    }

    #[test]
    fn normalize_accepts_loopback_hosts_with_explicit_nonzero_port() {
        let urls = vec![
            "http://127.0.0.1:1234/callback?code=abc".to_string(),
            "http://localhost:1234/callback?code=def".to_string(),
            "http://[::1]:1234/callback?code=ghi".to_string(),
        ];
        let normalized = normalize_oauth_redirect_request_urls(urls.clone());
        assert_eq!(normalized, urls);
    }

    #[test]
    fn normalize_rejects_non_loopback_hosts() {
        let urls = vec![
            "http://example.com:1234/callback?code=abc".to_string(),
            "http://127.0.0.2:1234/callback?code=def".to_string(),
            "http://[::2]:1234/callback?code=ghi".to_string(),
        ];
        let normalized = normalize_oauth_redirect_request_urls(urls);
        assert!(normalized.is_empty());
    }

    #[test]
    fn normalize_rejects_missing_port_or_zero_port() {
        let urls = vec![
            "http://127.0.0.1/callback?code=abc".to_string(),
            "http://localhost/callback?code=def".to_string(),
            "http://[::1]/callback?code=ghi".to_string(),
            "http://127.0.0.1:0/callback?code=jkl".to_string(),
            "http://localhost:0/callback?code=mno".to_string(),
            "http://[::1]:0/callback?code=pqr".to_string(),
        ];
        let normalized = normalize_oauth_redirect_request_urls(urls);
        assert!(normalized.is_empty());
    }

    #[test]
    fn loopback_redirect_parse_accepts_localhost_and_includes_both_bind_addrs() {
        let parsed = LoopbackRedirectUri::parse("http://localhost:4242/callback").unwrap();
        assert_eq!(parsed.host, LoopbackHost::Localhost);
        assert_eq!(parsed.port, 4242);
        assert_eq!(parsed.expected_path, "/callback");
        assert!(parsed
            .bind_addrs
            .contains(&SocketAddr::new(IpAddr::V4(Ipv4Addr::LOCALHOST), 4242)));
        assert!(parsed
            .bind_addrs
            .contains(&SocketAddr::new(IpAddr::V6(Ipv6Addr::LOCALHOST), 4242)));
    }

    #[test]
    fn loopback_redirect_parse_rejects_missing_port_and_zero_port() {
        assert_eq!(
            LoopbackRedirectUri::parse("http://localhost/callback").unwrap_err(),
            LoopbackRedirectUriError::MissingPort {
                uri: "http://localhost/callback".to_string()
            }
        );
        assert_eq!(
            LoopbackRedirectUri::parse("http://localhost:0/callback").unwrap_err(),
            LoopbackRedirectUriError::ZeroPort {
                uri: "http://localhost:0/callback".to_string()
            }
        );
    }

    #[test]
    fn loopback_redirect_parse_rejects_non_loopback_host() {
        let err = LoopbackRedirectUri::parse("http://example.com:1234/callback").unwrap_err();
        assert_eq!(
            err,
            LoopbackRedirectUriError::UnsupportedHost {
                uri: "http://example.com:1234/callback".to_string(),
                host: "example.com".to_string()
            }
        );
    }
}
