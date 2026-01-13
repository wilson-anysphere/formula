use std::net::{Ipv4Addr, Ipv6Addr};
use url::{Host, Url};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum LoopbackHostKind {
    Ipv4Loopback,
    Ipv6Loopback,
    Localhost,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct LoopbackRedirectUri {
    pub host_kind: LoopbackHostKind,
    pub port: u16,
    pub path: String,
    pub normalized_redirect_uri: String,
}

pub fn parse_loopback_redirect_uri(redirect_uri: &str) -> Result<LoopbackRedirectUri, String> {
    let parsed = Url::parse(redirect_uri.trim())
        .map_err(|err| format!("Invalid OAuth redirect URI: {err}"))?;

    if parsed.scheme() != "http" {
        return Err("Loopback OAuth redirect capture requires an http:// redirect URI".to_string());
    }

    let port = parsed
        .port()
        .ok_or_else(|| "Loopback OAuth redirect URI must include an explicit port".to_string())?;
    if port == 0 {
        return Err("Loopback OAuth redirect URI must not use port 0".to_string());
    }

    let host_kind = match parsed.host() {
        Some(Host::Ipv4(addr)) if addr == Ipv4Addr::LOCALHOST => LoopbackHostKind::Ipv4Loopback,
        Some(Host::Ipv6(addr)) if addr == Ipv6Addr::LOCALHOST => LoopbackHostKind::Ipv6Loopback,
        Some(Host::Domain(domain)) if domain.eq_ignore_ascii_case("localhost") => {
            LoopbackHostKind::Localhost
        }
        _ => {
            return Err(
                "Loopback OAuth redirect capture supports only 127.0.0.1, localhost, and [::1]"
                    .to_string(),
            );
        }
    };

    Ok(LoopbackRedirectUri {
        host_kind,
        port,
        path: parsed.path().to_string(),
        normalized_redirect_uri: parsed.to_string(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_ipv4_loopback() {
        let parsed = parse_loopback_redirect_uri("http://127.0.0.1:4242/callback").unwrap();
        assert_eq!(parsed.host_kind, LoopbackHostKind::Ipv4Loopback);
        assert_eq!(parsed.port, 4242);
        assert_eq!(parsed.path, "/callback");
        assert_eq!(parsed.normalized_redirect_uri, "http://127.0.0.1:4242/callback");
    }

    #[test]
    fn parses_localhost() {
        let parsed = parse_loopback_redirect_uri("http://localhost:4242/callback").unwrap();
        assert_eq!(parsed.host_kind, LoopbackHostKind::Localhost);
        assert_eq!(parsed.port, 4242);
        assert_eq!(parsed.path, "/callback");
        assert_eq!(parsed.normalized_redirect_uri, "http://localhost:4242/callback");
    }

    #[test]
    fn parses_ipv6_loopback() {
        let parsed = parse_loopback_redirect_uri("http://[::1]:4242/callback").unwrap();
        assert_eq!(parsed.host_kind, LoopbackHostKind::Ipv6Loopback);
        assert_eq!(parsed.port, 4242);
        assert_eq!(parsed.path, "/callback");
        assert_eq!(parsed.normalized_redirect_uri, "http://[::1]:4242/callback");
    }

    #[test]
    fn rejects_non_http_scheme() {
        let err = parse_loopback_redirect_uri("https://127.0.0.1:4242/callback").unwrap_err();
        assert!(err.contains("http://"));
    }

    #[test]
    fn rejects_non_loopback_hosts() {
        let err = parse_loopback_redirect_uri("http://example.com:4242/callback").unwrap_err();
        assert!(err.contains("127.0.0.1"));
    }

    #[test]
    fn rejects_missing_port() {
        let err = parse_loopback_redirect_uri("http://127.0.0.1/callback").unwrap_err();
        assert!(err.contains("explicit port"));
    }

    #[test]
    fn rejects_port_zero() {
        let err = parse_loopback_redirect_uri("http://127.0.0.1:0/callback").unwrap_err();
        assert!(err.contains("port 0"));
    }
}

