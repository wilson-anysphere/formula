use std::collections::HashSet;
use std::net::{Ipv4Addr, Ipv6Addr};
use std::sync::{Arc, Mutex};

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

/// Hard cap on concurrently-active RFC 8252 loopback listeners.
///
/// This exists to prevent a compromised webview from repeatedly invoking the
/// `oauth_loopback_listen` command with many distinct redirect URIs/ports and exhausting OS
/// resources (file descriptors, tasks, etc).
///
/// Behavior when the cap is exceeded: **reject** the request with a
/// `"Too many active OAuth listeners"` error rather than evicting an existing listener (eviction
/// would be surprising and could break an in-flight sign-in flow).
pub const OAUTH_LOOPBACK_MAX_ACTIVE_LISTENERS: usize = 3;

#[derive(Debug, Default)]
pub struct OauthLoopbackState {
    active_redirect_uris: HashSet<String>,
}

impl OauthLoopbackState {
    pub fn active_count(&self) -> usize {
        self.active_redirect_uris.len()
    }

    pub fn is_active(&self, redirect_uri: &str) -> bool {
        self.active_redirect_uris.contains(redirect_uri)
    }
}

pub type SharedOauthLoopbackState = Arc<Mutex<OauthLoopbackState>>;

#[derive(Debug)]
pub enum AcquireOauthLoopbackListener {
    /// A listener for the given redirect URI is already active; no-op.
    AlreadyActive,
    /// A new listener was registered; hold the guard for as long as the listener is active.
    Acquired(OauthLoopbackListenerGuard),
}

/// RAII guard that unregisters an active loopback listener on drop.
///
/// This is used to guarantee cleanup on *all* exit paths: normal task completion, early errors
/// after reserving the slot, panics inside the listener task, etc.
#[derive(Debug)]
pub struct OauthLoopbackListenerGuard {
    state: SharedOauthLoopbackState,
    redirect_uri: String,
}

impl Drop for OauthLoopbackListenerGuard {
    fn drop(&mut self) {
        if let Ok(mut guard) = self.state.lock() {
            guard.active_redirect_uris.remove(&self.redirect_uri);
        }
    }
}

pub fn acquire_oauth_loopback_listener(
    state: &SharedOauthLoopbackState,
    redirect_uri: String,
) -> Result<AcquireOauthLoopbackListener, String> {
    acquire_oauth_loopback_listener_with_cap(state, redirect_uri, OAUTH_LOOPBACK_MAX_ACTIVE_LISTENERS)
}

pub fn acquire_oauth_loopback_listener_with_cap(
    state: &SharedOauthLoopbackState,
    redirect_uri: String,
    max_active: usize,
) -> Result<AcquireOauthLoopbackListener, String> {
    let mut guard = state.lock().unwrap();

    if guard.active_redirect_uris.contains(&redirect_uri) {
        return Ok(AcquireOauthLoopbackListener::AlreadyActive);
    }

    if guard.active_redirect_uris.len() >= max_active {
        return Err("Too many active OAuth listeners".to_string());
    }

    guard.active_redirect_uris.insert(redirect_uri.clone());
    drop(guard);

    Ok(AcquireOauthLoopbackListener::Acquired(
        OauthLoopbackListenerGuard {
            state: state.clone(),
            redirect_uri,
        },
    ))
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

    #[test]
    fn enforces_cap_and_allows_idempotent_reuse() {
        let state: SharedOauthLoopbackState = Arc::new(Mutex::new(OauthLoopbackState::default()));

        let uri1 = "http://127.0.0.1:1234/callback".to_string();
        let uri2 = "http://127.0.0.1:2345/callback".to_string();
        let uri3 = "http://127.0.0.1:3456/callback".to_string();

        let guard1 = match acquire_oauth_loopback_listener_with_cap(&state, uri1.clone(), 2).unwrap()
        {
            AcquireOauthLoopbackListener::Acquired(guard) => guard,
            AcquireOauthLoopbackListener::AlreadyActive => panic!("expected new guard"),
        };

        // Re-registering the same URI should not consume another slot.
        assert!(matches!(
            acquire_oauth_loopback_listener_with_cap(&state, uri1.clone(), 2).unwrap(),
            AcquireOauthLoopbackListener::AlreadyActive
        ));

        let guard2 = match acquire_oauth_loopback_listener_with_cap(&state, uri2.clone(), 2).unwrap()
        {
            AcquireOauthLoopbackListener::Acquired(guard) => guard,
            AcquireOauthLoopbackListener::AlreadyActive => panic!("expected new guard"),
        };

        // Cap reached; a distinct new URI should be rejected.
        let err = acquire_oauth_loopback_listener_with_cap(&state, uri3.clone(), 2).unwrap_err();
        assert_eq!(err, "Too many active OAuth listeners");
        assert_eq!(state.lock().unwrap().active_count(), 2);

        drop(guard1);
        assert_eq!(state.lock().unwrap().active_count(), 1);

        // After cleanup, we should be able to acquire a new URI again.
        let guard3 = match acquire_oauth_loopback_listener_with_cap(&state, uri3.clone(), 2).unwrap()
        {
            AcquireOauthLoopbackListener::Acquired(guard) => guard,
            AcquireOauthLoopbackListener::AlreadyActive => panic!("expected new guard"),
        };

        assert_eq!(state.lock().unwrap().active_count(), 2);
        drop(guard2);
        drop(guard3);
        assert_eq!(state.lock().unwrap().active_count(), 0);
    }

    #[test]
    fn guard_drop_cleans_up_on_early_error_paths() {
        let state: SharedOauthLoopbackState = Arc::new(Mutex::new(OauthLoopbackState::default()));
        let uri = "http://127.0.0.1:5555/callback".to_string();

        let guard = match acquire_oauth_loopback_listener_with_cap(&state, uri.clone(), 1).unwrap()
        {
            AcquireOauthLoopbackListener::Acquired(guard) => guard,
            AcquireOauthLoopbackListener::AlreadyActive => panic!("expected new guard"),
        };
        assert!(state.lock().unwrap().is_active(&uri));

        // Simulate an early error (e.g. bind failure) by dropping the guard immediately.
        drop(guard);
        assert!(!state.lock().unwrap().is_active(&uri));
    }
}

