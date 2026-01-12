/// State machine for desktop "open file" IPC.
///
/// Tauri does not guarantee that emitted events are queued before JS listeners are registered.
/// To avoid dropping file-open requests on cold start, the Rust backend queues incoming
/// `open-file` requests until the frontend signals readiness via an `open-file-ready` event.
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

        // Ensure the ready signal is listened for, and that the queue/flush helpers are used.
        assert!(
            main_rs.contains("listen(OPEN_FILE_READY_EVENT"),
            "desktop main.rs must listen for OPEN_FILE_READY_EVENT"
        );
        assert!(
            main_rs.contains("queue_or_emit("),
            "desktop main.rs must queue open-file requests via OpenFileState::queue_or_emit"
        );
        assert!(
            main_rs.contains("mark_ready_and_drain("),
            "desktop main.rs must flush pending open-file requests via OpenFileState::mark_ready_and_drain"
        );

        // Extra guardrail: the backend should only flip readiness in response to the frontend
        // readiness signal. If `mark_ready_and_drain` starts getting called elsewhere (e.g. during
        // startup), cold-start file-open events can be emitted before the JS listener exists.
        let ready_calls = main_rs.matches("mark_ready_and_drain(").count();
        assert_eq!(
            ready_calls, 1,
            "expected exactly one mark_ready_and_drain call in desktop main.rs, found {ready_calls}"
        );
    }
}
