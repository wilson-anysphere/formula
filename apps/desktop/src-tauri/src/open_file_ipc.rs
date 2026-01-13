/// State machine for desktop "open file" IPC.
///
/// Tauri does not guarantee that emitted events are queued before JS listeners are registered.
/// To avoid dropping file-open requests on cold start, the Rust backend queues incoming
/// `open-file` requests until the frontend signals readiness via an `open-file-ready` event.
///
/// The pending queue is intentionally bounded to avoid unbounded memory growth if the OS delivers
/// many open-file events before the frontend installs its event listeners. When the cap is
/// exceeded, we drop the **oldest** paths and keep the most recent ones so the latest user action
/// wins.
const MAX_PENDING_PATHS: usize = 64;

#[derive(Debug, Default)]
pub struct OpenFileState {
    ready: bool,
    pending_paths: Vec<String>,
}

impl OpenFileState {
    pub fn is_ready(&self) -> bool {
        self.ready
    }

    pub fn pending_len(&self) -> usize {
        self.pending_paths.len()
    }

    /// Queue an open-file request if the frontend isn't ready yet.
    ///
    /// Returns `Some(paths)` if the caller should emit the event immediately (frontend is ready),
    /// or `None` if the request was queued.
    pub fn queue_or_emit(&mut self, paths: Vec<String>) -> Option<Vec<String>> {
        if self.ready {
            Some(paths)
        } else {
            self.pending_paths.extend(paths);
            if self.pending_paths.len() > MAX_PENDING_PATHS {
                let overflow = self.pending_paths.len() - MAX_PENDING_PATHS;
                // Drop oldest paths first, keeping the most recent ones.
                self.pending_paths.drain(0..overflow);
            }
            None
        }
    }

    /// Mark the frontend as ready and return any queued paths to flush.
    ///
    /// This flush happens at most once; subsequent calls return an empty Vec.
    pub fn mark_ready_and_drain(&mut self) -> Vec<String> {
        if self.ready {
            Vec::new()
        } else {
            self.ready = true;
            std::mem::take(&mut self.pending_paths)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn queues_until_ready_flushes_once_and_emits_immediately_after() {
        let mut state = OpenFileState::default();
        assert!(!state.is_ready());
        assert_eq!(state.pending_len(), 0);

        assert!(state.queue_or_emit(vec!["a.xlsx".into()]).is_none());
        assert_eq!(state.pending_len(), 1);

        assert!(state
            .queue_or_emit(vec!["b.csv".into(), "c.xlsm".into()])
            .is_none());
        assert_eq!(state.pending_len(), 3);

        let flushed = state.mark_ready_and_drain();
        assert!(state.is_ready());
        assert_eq!(flushed, vec!["a.xlsx", "b.csv", "c.xlsm"]);
        assert_eq!(state.pending_len(), 0);

        // The flush should happen exactly once.
        assert!(state.mark_ready_and_drain().is_empty());
        assert_eq!(state.pending_len(), 0);

        // Subsequent requests should be emitted immediately without growing the queue.
        let immediate = state.queue_or_emit(vec!["d.xlsx".into()]);
        assert_eq!(immediate, Some(vec!["d.xlsx".into()]));
        assert_eq!(state.pending_len(), 0);
    }

    #[test]
    fn tauri_main_wires_open_file_ready_handshake() {
        // `src/main.rs` is only compiled when the `desktop` feature is enabled (it depends on the
        // system WebView toolchain on Linux). Still, we want a lightweight regression test that
        // fails if someone removes the open-file readiness handshake wiring.
        //
        // Use a simple source-level check so this test runs in headless CI.
        let main_rs = include_str!("main.rs");

        // Ensure runtime open-file requests route through the queue.
        let handle_open_start = main_rs
            .find("fn handle_open_file_request")
            .expect("desktop main.rs missing handle_open_file_request()");
        let handle_open_end = main_rs[handle_open_start..]
            .find("fn handle_oauth_redirect_request")
            .map(|idx| handle_open_start + idx)
            .expect("desktop main.rs missing handle_oauth_redirect_request() (used to bound open-file handler)");
        let handle_open_body = &main_rs[handle_open_start..handle_open_end];
        assert!(
            handle_open_body.contains(".queue_or_emit("),
            "handle_open_file_request must call OpenFileState::queue_or_emit"
        );

        // Ensure cold-start argv open-file requests are queued too.
        let init_start = main_rs
            .find("let initial_paths")
            .expect("desktop main.rs missing initial_paths extraction");
        let init_window = main_rs
            .get(init_start..init_start.saturating_add(800))
            .unwrap_or(&main_rs[init_start..]);
        assert!(
            init_window.contains(".queue_or_emit("),
            "desktop main.rs must queue initial argv open-file requests via OpenFileState::queue_or_emit"
        );

        // Ensure the ready signal is listened for, and that readiness is flipped + flushed from
        // within that listener.
        let ready_start = main_rs
            .find("listen(OPEN_FILE_READY_EVENT")
            .expect("desktop main.rs must listen for OPEN_FILE_READY_EVENT");
        let ready_after = &main_rs[ready_start..];
        let ready_end = ready_after
            .find("});")
            .map(|idx| idx + 3)
            .expect("failed to locate end of OPEN_FILE_READY_EVENT listener (expected `});`)");
        let ready_body = &ready_after[..ready_end];

        assert!(
            ready_body.contains(".mark_ready_and_drain("),
            "OPEN_FILE_READY_EVENT listener must call OpenFileState::mark_ready_and_drain"
        );
        let ready_calls_in_listener = ready_body.matches(".mark_ready_and_drain(").count();
        assert_eq!(
            ready_calls_in_listener, 1,
            "expected exactly one mark_ready_and_drain call inside OPEN_FILE_READY_EVENT listener, found {ready_calls_in_listener}"
        );

        // Extra guardrail: the backend should only flip *open-file* readiness in response to the
        // frontend readiness signal. If `mark_ready_and_drain` starts getting called on the
        // `SharedOpenFileState` elsewhere (e.g. during startup), cold-start file-open events can
        // be emitted before the JS listener exists.
        let mut open_file_ready_calls = 0;
        for (idx, _) in main_rs.match_indices("state::<SharedOpenFileState>") {
            let window = main_rs.get(idx..idx.saturating_add(300)).unwrap_or(&main_rs[idx..]);
            open_file_ready_calls += window.matches(".mark_ready_and_drain(").count();
        }
        assert_eq!(
            open_file_ready_calls, 1,
            "expected exactly one mark_ready_and_drain call associated with SharedOpenFileState in desktop main.rs, found {open_file_ready_calls}"
        );
    }

    #[test]
    fn caps_pending_paths_and_drops_oldest() {
        let mut state = OpenFileState::default();

        let mut paths = Vec::new();
        for idx in 0..(MAX_PENDING_PATHS + 5) {
            paths.push(format!("p{idx}"));
        }

        assert!(state.queue_or_emit(paths).is_none());
        assert_eq!(state.pending_len(), MAX_PENDING_PATHS);

        let flushed = state.mark_ready_and_drain();
        let expected: Vec<String> = (5..(MAX_PENDING_PATHS + 5))
            .map(|idx| format!("p{idx}"))
            .collect();
        assert_eq!(flushed, expected);
    }
}
