/// State machine for desktop "oauth redirect" IPC.
///
/// Tauri does not guarantee that emitted events are queued before JS listeners are registered.
/// To avoid dropping OAuth redirect notifications on cold start, the Rust backend queues incoming
/// redirect URLs until the frontend signals readiness via an `oauth-redirect-ready` event.
///
/// The pending queue is intentionally bounded to avoid unbounded memory growth if the OS delivers
/// many redirects before the frontend installs its event listeners. When the cap is exceeded, we
/// drop the **oldest** URLs and keep the most recent ones so the latest user action wins.
const MAX_PENDING_URLS: usize = 64;

#[derive(Debug, Default)]
pub struct OauthRedirectState {
    ready: bool,
    pending_urls: Vec<String>,
}

impl OauthRedirectState {
    pub fn is_ready(&self) -> bool {
        self.ready
    }

    pub fn pending_len(&self) -> usize {
        self.pending_urls.len()
    }

    /// Queue a redirect if the frontend isn't ready yet.
    ///
    /// Returns `Some(urls)` if the caller should emit the event immediately (frontend is ready),
    /// or `None` if the request was queued.
    pub fn queue_or_emit(&mut self, urls: Vec<String>) -> Option<Vec<String>> {
        if self.ready {
            Some(urls)
        } else {
            self.pending_urls.extend(urls);
            if self.pending_urls.len() > MAX_PENDING_URLS {
                let overflow = self.pending_urls.len() - MAX_PENDING_URLS;
                // Drop oldest URLs first, keeping the most recent ones.
                self.pending_urls.drain(0..overflow);
            }
            None
        }
    }

    /// Mark the frontend as ready and return any queued URLs to flush.
    ///
    /// This flush happens at most once; subsequent calls return an empty Vec.
    pub fn mark_ready_and_drain(&mut self) -> Vec<String> {
        if self.ready {
            Vec::new()
        } else {
            self.ready = true;
            std::mem::take(&mut self.pending_urls)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn queues_until_ready_flushes_once_and_emits_immediately_after() {
        let mut state = OauthRedirectState::default();
        assert!(!state.is_ready());
        assert_eq!(state.pending_len(), 0);

        assert!(state.queue_or_emit(vec!["formula://a".into()]).is_none());
        assert_eq!(state.pending_len(), 1);

        assert!(state
            .queue_or_emit(vec!["formula://b".into(), "formula://c".into()])
            .is_none());
        assert_eq!(state.pending_len(), 3);

        let flushed = state.mark_ready_and_drain();
        assert!(state.is_ready());
        assert_eq!(
            flushed,
            vec!["formula://a", "formula://b", "formula://c"]
        );
        assert_eq!(state.pending_len(), 0);

        // The flush should happen exactly once.
        assert!(state.mark_ready_and_drain().is_empty());
        assert_eq!(state.pending_len(), 0);

        // Subsequent requests should be emitted immediately without growing the queue.
        let immediate = state.queue_or_emit(vec!["formula://d".into()]);
        assert_eq!(immediate, Some(vec!["formula://d".into()]));
        assert_eq!(state.pending_len(), 0);
    }

    #[test]
    fn caps_pending_urls_and_drops_oldest() {
        let mut state = OauthRedirectState::default();

        let mut urls = Vec::new();
        for idx in 0..(MAX_PENDING_URLS + 4) {
            urls.push(format!("formula://u{idx}"));
        }

        assert!(state.queue_or_emit(urls).is_none());
        assert_eq!(state.pending_len(), MAX_PENDING_URLS);

        let flushed = state.mark_ready_and_drain();
        let expected: Vec<String> = (4..(MAX_PENDING_URLS + 4))
            .map(|idx| format!("formula://u{idx}"))
            .collect();
        assert_eq!(flushed, expected);
    }

    #[test]
    fn tauri_main_wires_oauth_redirect_ready_handshake() {
        // `src/main.rs` is only compiled when the `desktop` feature is enabled (it depends on the
        // system WebView toolchain on Linux). Still, we want a lightweight regression test that
        // fails if someone removes the oauth-redirect readiness handshake wiring.
        //
        // Use a simple source-level check so this test runs in headless CI.
        let main_rs = include_str!("main.rs");

        // Ensure runtime oauth redirect requests route through the queue.
        let handle_oauth_start = main_rs
            .find("fn handle_oauth_redirect_request")
            .expect("desktop main.rs missing handle_oauth_redirect_request()");
        let handle_oauth_end = main_rs[handle_oauth_start..]
            .find("fn extract_open_file_paths")
            .map(|idx| handle_oauth_start + idx)
            .expect("desktop main.rs missing extract_open_file_paths() (used to bound oauth handler)");
        let handle_oauth_body = &main_rs[handle_oauth_start..handle_oauth_end];
        assert!(
            handle_oauth_body.contains(".queue_or_emit("),
            "handle_oauth_redirect_request must call OauthRedirectState::queue_or_emit"
        );

        // Ensure cold-start argv oauth redirect requests are queued too.
        let init_start = main_rs
            .find("let initial_oauth_urls")
            .expect("desktop main.rs missing initial_oauth_urls extraction");
        let init_window = main_rs
            .get(init_start..init_start.saturating_add(800))
            .unwrap_or(&main_rs[init_start..]);
        assert!(
            init_window.contains(".queue_or_emit("),
            "desktop main.rs must queue initial argv oauth redirect requests via OauthRedirectState::queue_or_emit"
        );

        // Ensure the ready signal is listened for, and that readiness is flipped + flushed from
        // within that listener.
        let ready_start = main_rs
            .find("listen(OAUTH_REDIRECT_READY_EVENT")
            .expect("desktop main.rs must listen for OAUTH_REDIRECT_READY_EVENT");
        let ready_after = &main_rs[ready_start..];
        let ready_end = ready_after
            .find("});")
            .map(|idx| idx + 3)
            .expect("failed to locate end of OAUTH_REDIRECT_READY_EVENT listener (expected `});`)");
        let ready_body = &ready_after[..ready_end];

        assert!(
            ready_body.contains(".mark_ready_and_drain("),
            "OAUTH_REDIRECT_READY_EVENT listener must call OauthRedirectState::mark_ready_and_drain"
        );
        let ready_calls_in_listener = ready_body.matches(".mark_ready_and_drain(").count();
        assert_eq!(
            ready_calls_in_listener, 1,
            "expected exactly one mark_ready_and_drain call inside OAUTH_REDIRECT_READY_EVENT listener, found {ready_calls_in_listener}"
        );
    }
}

