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
    use reqwest::Method;

    let parsed_url = reqwest::Url::parse(url).map_err(|e| format!("Invalid url: {e}"))?;
    match parsed_url.scheme() {
        "http" | "https" => {}
        other => {
            return Err(format!(
                "Unsupported url scheme for network_fetch: {other} (only http/https allowed)"
            ));
        }
    }

    let method = init
        .get("method")
        .and_then(|v| v.as_str())
        .unwrap_or("GET")
        .to_uppercase();
    let method = Method::from_bytes(method.as_bytes())
        .map_err(|_| format!("Unsupported method: {method}"))?;

    let client = reqwest::Client::builder()
        .redirect(reqwest::redirect::Policy::limited(10))
        .build()
        .map_err(|e| e.to_string())?;

    let mut req = client.request(method, parsed_url);
    req = apply_request_init(req, init)?;

    let mut response = req.send().await.map_err(|e| e.to_string())?;
    let status = response.status();
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
