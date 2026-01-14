use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct NetworkFetchResult {
    pub ok: bool,
    pub status: u16,
    pub status_text: String,
    pub url: String,
    pub headers: Vec<(String, String)>,
    pub body_text: String,
}

fn normalize_header_value(value: &JsonValue) -> String {
    match value {
        JsonValue::String(s) => s.to_string(),
        JsonValue::Number(n) => n.to_string(),
        JsonValue::Bool(b) => b.to_string(),
        JsonValue::Null => String::new(),
        other => other.to_string(),
    }
}

fn apply_request_init(
    mut builder: reqwest::RequestBuilder,
    init: &JsonValue,
) -> Result<reqwest::RequestBuilder, String> {
    if let Some(headers_val) = init.get("headers") {
        if let Some(map) = headers_val.as_object() {
            for (k, v) in map {
                builder = builder.header(k, normalize_header_value(v));
            }
        } else if let Some(arr) = headers_val.as_array() {
            for entry in arr {
                let Some(pair) = entry.as_array() else {
                    continue;
                };
                if pair.len() < 2 {
                    continue;
                }
                let key = pair[0].as_str().unwrap_or("").to_string();
                if key.trim().is_empty() {
                    continue;
                }
                builder = builder.header(key, normalize_header_value(&pair[1]));
            }
        }
    }

    if let Some(body_val) = init.get("body") {
        if !body_val.is_null() {
            if let Some(body_str) = body_val.as_str() {
                builder = builder.body(body_str.to_string());
            } else {
                // Best-effort: serialize non-string bodies (e.g. objects) as JSON.
                builder = builder.body(body_val.to_string());
            }
        }
    }

    Ok(builder)
}

/// Core implementation of the `network_fetch` IPC command.
///
/// This is intentionally kept independent of Tauri so it can be unit tested without the `desktop`
/// feature enabled.
pub async fn network_fetch_impl(url: &str, init: &JsonValue) -> Result<NetworkFetchResult, String> {
    network_fetch_impl_with_debug_assertions(url, init, cfg!(debug_assertions)).await
}

async fn network_fetch_impl_with_debug_assertions(
    url: &str,
    init: &JsonValue,
    debug_assertions: bool,
) -> Result<NetworkFetchResult, String> {
    use reqwest::header::LOCATION;
    use reqwest::Method;

    let parsed_url = reqwest::Url::parse(url).map_err(|e| format!("Invalid url: {e}"))?;
    crate::commands::ensure_ipc_network_url_allowed(&parsed_url, "network_fetch", debug_assertions)?;

    let method = init
        .get("method")
        .and_then(|v| v.as_str())
        .unwrap_or("GET")
        .to_uppercase();
    let method = Method::from_bytes(method.as_bytes())
        .map_err(|_| format!("Unsupported method: {method}"))?;

    let client = reqwest::Client::builder()
        .redirect(reqwest::redirect::Policy::custom(move |attempt| {
            // Keep redirect behavior aligned with browser `fetch` (follow redirects), but enforce
            // the same scheme/host allowlist for every hop to prevent `https -> http` downgrade
            // redirects from bypassing the IPC URL policy in release builds.
            if attempt.previous().len() >= 10 {
                return attempt.stop();
            }

            match crate::commands::ensure_ipc_network_url_allowed(
                attempt.url(),
                "network_fetch redirect",
                debug_assertions,
            ) {
                Ok(()) => attempt.follow(),
                Err(_) => attempt.stop(),
            }
        }))
        .build()
        .map_err(|e| e.to_string())?;

    let mut req = client.request(method, parsed_url);
    req = apply_request_init(req, init)?;

    let mut response = req.send().await.map_err(|e| e.to_string())?;
    let status = response.status();

    // If redirect following was stopped (e.g. because a hop would violate the IPC URL policy), make
    // the failure explicit rather than returning the raw 3xx response. This keeps behavior aligned
    // with browser `fetch` (which follows redirects by default) while still enforcing our stricter
    // release-mode http allowlist.
    if status.is_redirection() {
        if let Some(location) = response.headers().get(LOCATION).and_then(|v| v.to_str().ok()) {
            if let Ok(target) = response.url().join(location) {
                if let Err(err) = crate::commands::ensure_ipc_network_url_allowed(
                    &target,
                    "network_fetch redirect",
                    debug_assertions,
                ) {
                    return Err(err);
                }
            }
        }
    }

    let status_text = status.canonical_reason().unwrap_or("").to_string();
    let final_url = response.url().to_string();

    let headers = response
        .headers()
        .iter()
        .map(|(k, v)| {
            (
                k.as_str().to_string(),
                v.to_str().unwrap_or_default().to_string(),
            )
        })
        .collect::<Vec<_>>();

    let bytes = crate::network_limits::read_response_body_with_limit(
        &mut response,
        crate::network_limits::NETWORK_FETCH_MAX_BODY_BYTES,
        "network_fetch",
    )
    .await?;
    let body_text = String::from_utf8_lossy(&bytes).to_string();

    Ok(NetworkFetchResult {
        ok: status.is_success(),
        status: status.as_u16(),
        status_text,
        url: final_url,
        headers,
        body_text,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use tokio::io::{AsyncReadExt, AsyncWriteExt};
    use tokio::net::TcpListener;

    async fn spawn_server(body: Vec<u8>, include_content_length: bool) -> String {
        let listener = TcpListener::bind(("127.0.0.1", 0)).await.unwrap();
        let addr = listener.local_addr().unwrap();

        tokio::spawn(async move {
            let Ok((mut socket, _peer)) = listener.accept().await else {
                return;
            };

            // Best-effort: read request headers so the client doesn't get a connection reset.
            let mut buf = [0u8; 1024];
            let mut req = Vec::new();
            loop {
                match socket.read(&mut buf).await {
                    Ok(0) => break,
                    Ok(n) => {
                        req.extend_from_slice(&buf[..n]);
                        if req.windows(4).any(|w| w == b"\r\n\r\n") {
                            break;
                        }
                        if req.len() > 16 * 1024 {
                            break;
                        }
                    }
                    Err(_) => break,
                }
            }

            let mut headers = String::from("HTTP/1.1 200 OK\r\n");
            headers.push_str("Content-Type: text/plain\r\n");
            headers.push_str("Connection: close\r\n");
            if include_content_length {
                headers.push_str(&format!("Content-Length: {}\r\n", body.len()));
            }
            headers.push_str("\r\n");

            let _ = socket.write_all(headers.as_bytes()).await;
            let _ = socket.write_all(&body).await;
            let _ = socket.shutdown().await;
        });

        format!("http://{addr}/")
    }

    async fn spawn_redirect_server() -> String {
        let listener = TcpListener::bind(("127.0.0.1", 0)).await.unwrap();
        let addr = listener.local_addr().unwrap();

        tokio::spawn(async move {
            let mut served = 0usize;
            while served < 2 {
                let Ok((mut socket, _peer)) = listener.accept().await else {
                    break;
                };

                // Best-effort: read request headers so the client doesn't get a connection reset.
                let mut buf = [0u8; 1024];
                let mut req = Vec::new();
                loop {
                    match socket.read(&mut buf).await {
                        Ok(0) => break,
                        Ok(n) => {
                            req.extend_from_slice(&buf[..n]);
                            if req.windows(4).any(|w| w == b"\r\n\r\n") {
                                break;
                            }
                            if req.len() > 16 * 1024 {
                                break;
                            }
                        }
                        Err(_) => break,
                    }
                }

                let req_line = req
                    .split(|b| *b == b'\n')
                    .next()
                    .map(|l| String::from_utf8_lossy(l).to_string())
                    .unwrap_or_default();
                let path = req_line
                    .split_whitespace()
                    .nth(1)
                    .unwrap_or("/")
                    .trim()
                    .to_string();

                if path == "/" {
                    let response =
                        "HTTP/1.1 302 Found\r\nLocation: /final\r\nConnection: close\r\nContent-Length: 0\r\n\r\n";
                    let _ = socket.write_all(response.as_bytes()).await;
                } else {
                    let body = b"redirect ok";
                    let headers = format!(
                        "HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\nContent-Length: {}\r\nConnection: close\r\n\r\n",
                        body.len()
                    );
                    let _ = socket.write_all(headers.as_bytes()).await;
                    let _ = socket.write_all(body).await;
                }

                let _ = socket.shutdown().await;
                served += 1;
            }
        });

        format!("http://{addr}/")
    }

    #[tokio::test]
    async fn follows_allowed_redirects() {
        let url = spawn_redirect_server().await;
        let result = network_fetch_impl(&url, &JsonValue::Null).await.unwrap();
        assert!(result.ok);
        assert_eq!(result.status, 200);
        assert_eq!(result.body_text, "redirect ok");
        assert!(
            result.url.ends_with("/final"),
            "expected final url to end with /final, got: {}",
            result.url
        );
    }

    #[tokio::test]
    async fn rejects_redirect_to_disallowed_scheme() {
        let listener = TcpListener::bind(("127.0.0.1", 0)).await.unwrap();
        let addr = listener.local_addr().unwrap();

        let server = tokio::spawn(async move {
            let Ok((mut socket, _peer)) = listener.accept().await else {
                return;
            };
            let mut buf = [0u8; 1024];
            let _ = socket.read(&mut buf).await;

            let response = "HTTP/1.1 302 Found\r\nLocation: ftp://example.com/\r\nConnection: close\r\nContent-Length: 0\r\n\r\n";
            let _ = socket.write_all(response.as_bytes()).await;
            let _ = socket.shutdown().await;
        });

        let url = format!("http://{addr}/");
        let err = network_fetch_impl(&url, &JsonValue::Null).await.unwrap_err();
        assert!(
            err.contains("network_fetch redirect") && err.contains("only http/https allowed"),
            "unexpected error: {err}"
        );

        server.await.expect("server task");
    }

    #[tokio::test]
    async fn rejects_remote_http_in_release_mode() {
        let err = network_fetch_impl_with_debug_assertions(
            "http://example.com/",
            &JsonValue::Null,
            false,
        )
        .await
        .unwrap_err();
        assert!(
            err.contains("network_fetch: http URLs are only allowed for localhost in release builds")
                && err.contains("example.com"),
            "unexpected error: {err}"
        );
    }

    #[tokio::test]
    async fn rejects_redirect_to_remote_http_in_release_mode() {
        let listener = TcpListener::bind(("127.0.0.1", 0)).await.unwrap();
        let addr = listener.local_addr().unwrap();

        let server = tokio::spawn(async move {
            let Ok((mut socket, _peer)) = listener.accept().await else {
                return;
            };
            let mut buf = [0u8; 1024];
            let _ = socket.read(&mut buf).await;

            let response = "HTTP/1.1 302 Found\r\nLocation: http://example.com/\r\nConnection: close\r\nContent-Length: 0\r\n\r\n";
            let _ = socket.write_all(response.as_bytes()).await;
            let _ = socket.shutdown().await;
        });

        let url = format!("http://{addr}/");
        let err = network_fetch_impl_with_debug_assertions(&url, &JsonValue::Null, false)
            .await
            .unwrap_err();
        assert!(
            err.contains("network_fetch redirect: http URLs are only allowed for localhost in release builds")
                && err.contains("example.com"),
            "unexpected error: {err}"
        );

        server.await.expect("server task");
    }

    #[tokio::test]
    async fn rejects_large_response_without_content_length() {
        let body = vec![b'a'; crate::network_limits::NETWORK_FETCH_MAX_BODY_BYTES + 1];
        let url = spawn_server(body, false).await;

        let err = network_fetch_impl(&url, &JsonValue::Null).await.unwrap_err();
        assert!(
            err.contains("Response body too large"),
            "expected error to mention response too large; got: {err}"
        );
        assert!(
            err.contains(&crate::network_limits::NETWORK_FETCH_MAX_BODY_BYTES.to_string()),
            "expected error to include byte limit; got: {err}"
        );
        assert!(
            err.contains(&(crate::network_limits::NETWORK_FETCH_MAX_BODY_BYTES + 1).to_string()),
            "expected error to include observed size; got: {err}"
        );
    }

    #[tokio::test]
    async fn returns_small_text_body() {
        let url = spawn_server(b"hello".to_vec(), true).await;

        let result = network_fetch_impl(&url, &JsonValue::Null).await.unwrap();
        assert!(result.ok);
        assert_eq!(result.status, 200);
        assert_eq!(result.status_text, "OK");
        assert_eq!(result.body_text, "hello");
        assert!(!result.url.is_empty());
        assert!(
            result
                .headers
                .iter()
                .any(|(k, _)| k.eq_ignore_ascii_case("content-type")),
            "expected content-type header in response headers"
        );
    }
}
