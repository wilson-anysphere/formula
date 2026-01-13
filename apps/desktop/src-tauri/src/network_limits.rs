//! Limits for Rust-mediated HTTP(S) responses used by desktop IPC commands.
//!
//! The Tauri WebView is treated as untrusted input: a compromised extension (or a buggy/malicious
//! marketplace endpoint) could otherwise force the backend process to allocate unbounded memory by
//! downloading arbitrarily large responses.

/// Maximum response payload size for `network_fetch` (before any base64 encoding).
pub const NETWORK_FETCH_MAX_BODY_BYTES: usize = 10 * 1024 * 1024; // 10 MiB

/// Maximum response payload size for marketplace JSON endpoints (`marketplace_search`,
/// `marketplace_get_extension`).
pub const MARKETPLACE_JSON_MAX_BODY_BYTES: usize = 5 * 1024 * 1024; // 5 MiB

/// Read a `reqwest::Response` body into memory, enforcing a hard maximum size.
///
/// Defense-in-depth:
/// - If `Content-Length` is present and exceeds `limit_bytes`, the call fails before downloading
///   the body.
/// - Otherwise, the body is streamed and the call fails once `limit_bytes + 1` bytes have been
///   observed.
///
/// Error messages are intentionally stable and include both the configured limit and the observed
/// (or advertised) size.
pub async fn read_response_body_with_limit(
    response: &mut reqwest::Response,
    limit_bytes: usize,
    context: &str,
) -> Result<Vec<u8>, String> {
    if let Some(content_length) = response.content_length() {
        if content_length > limit_bytes as u64 {
            return Err(format!(
                "Response body too large for {context} (limit {limit_bytes} bytes, Content-Length {content_length} bytes)"
            ));
        }
    }

    let max_bytes = limit_bytes.saturating_add(1);
    let mut out: Vec<u8> = Vec::new();

    loop {
        let Some(chunk) = response.chunk().await.map_err(|e| e.to_string())? else {
            break;
        };

        if out.len().saturating_add(chunk.len()) >= max_bytes {
            let remaining = max_bytes.saturating_sub(out.len());
            if remaining > 0 {
                out.extend_from_slice(&chunk[..remaining.min(chunk.len())]);
            }

            return Err(format!(
                "Response body too large for {context} (limit {limit_bytes} bytes, received {max_bytes} bytes)"
            ));
        }

        out.extend_from_slice(&chunk);
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use tokio::io::{AsyncReadExt, AsyncWriteExt};
    use tokio::net::TcpListener;

    async fn serve_once(body: Vec<u8>, content_length: Option<usize>) -> (String, tokio::task::JoinHandle<()>) {
        let listener = TcpListener::bind("127.0.0.1:0")
            .await
            .expect("bind test http listener");
        let addr = listener.local_addr().expect("listener addr");

        let handle = tokio::spawn(async move {
            let (mut socket, _) = listener.accept().await.expect("accept");

            // Best-effort: read until end of headers so the client doesn't see an early close on
            // some platforms.
            let mut buf = [0u8; 1024];
            let mut request = Vec::new();
            loop {
                let n = socket.read(&mut buf).await.expect("read request");
                if n == 0 {
                    break;
                }
                request.extend_from_slice(&buf[..n]);
                if request.windows(4).any(|w| w == b"\r\n\r\n") || request.len() > 16 * 1024 {
                    break;
                }
            }

            let mut headers = String::from("HTTP/1.1 200 OK\r\n");
            if let Some(len) = content_length {
                headers.push_str(&format!("Content-Length: {len}\r\n"));
            }
            headers.push_str("Connection: close\r\n\r\n");

            // The client may intentionally stop reading early once it hits the size cap, so ignore
            // broken pipe / connection reset errors while writing the response.
            let _ = socket.write_all(headers.as_bytes()).await;
            let _ = socket.write_all(&body).await;
        });

        (format!("http://{addr}/"), handle)
    }

    #[tokio::test]
    async fn read_response_body_with_limit_allows_small_bodies() {
        let (url, handle) = serve_once(b"hello".to_vec(), Some(5)).await;

        let client = reqwest::Client::builder()
            .redirect(reqwest::redirect::Policy::none())
            .build()
            .expect("build client");
        let mut response = client.get(url).send().await.expect("send");

        let body = read_response_body_with_limit(&mut response, 10, "test").await.unwrap();
        assert_eq!(body, b"hello");

        handle.await.expect("server task");
    }

    #[tokio::test]
    async fn read_response_body_with_limit_rejects_oversized_content_length() {
        let (url, handle) = serve_once(Vec::new(), Some(11)).await;

        let client = reqwest::Client::builder()
            .redirect(reqwest::redirect::Policy::none())
            .build()
            .expect("build client");
        let mut response = client.get(url).send().await.expect("send");

        let err = read_response_body_with_limit(&mut response, 10, "test")
            .await
            .expect_err("expected limit error");
        assert!(
            err.contains("limit 10") && err.contains("Content-Length 11"),
            "unexpected error: {err}"
        );

        handle.await.expect("server task");
    }

    #[tokio::test]
    async fn read_response_body_with_limit_streams_until_limit_plus_one() {
        // No Content-Length header, so the implementation should stream and stop once it observes
        // limit + 1 bytes.
        let body = vec![b'a'; 32];
        let (url, handle) = serve_once(body, None).await;

        let client = reqwest::Client::builder()
            .redirect(reqwest::redirect::Policy::none())
            .build()
            .expect("build client");
        let mut response = client.get(url).send().await.expect("send");

        let err = read_response_body_with_limit(&mut response, 10, "test")
            .await
            .expect_err("expected limit error");
        assert!(
            err.contains("limit 10") && err.contains("received 11"),
            "unexpected error: {err}"
        );

        handle.await.expect("server task");
    }
}
