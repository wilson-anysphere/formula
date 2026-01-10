use crate::storage::{CellChange, Result as StorageResult, Storage};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::Arc;
use std::time::Duration;
use tokio::sync::{mpsc, oneshot};
use tokio::task::JoinHandle;
use tokio::time::Instant;

#[derive(Debug, Clone)]
pub struct AutoSaveConfig {
    /// Debounce delay (default: 1 second).
    pub save_delay: Duration,
    /// Maximum delay between flushes (default: 5 seconds).
    pub max_delay: Duration,
}

impl Default for AutoSaveConfig {
    fn default() -> Self {
        Self {
            save_delay: Duration::from_secs(1),
            max_delay: Duration::from_secs(5),
        }
    }
}

enum Command {
    Record(CellChange),
    Flush(oneshot::Sender<StorageResult<()>>),
    Shutdown(oneshot::Sender<StorageResult<()>>),
}

/// Debounced, batched persistence for cell edits.
///
/// The autosave manager is intentionally storage-only (no UI concepts). Higher
/// layers can translate user operations into `CellChange`s.
pub struct AutoSaveManager {
    tx: mpsc::UnboundedSender<Command>,
    save_count: Arc<AtomicUsize>,
    handle: JoinHandle<()>,
}

impl AutoSaveManager {
    pub fn spawn(storage: Storage, config: AutoSaveConfig) -> Self {
        let (tx, mut rx) = mpsc::unbounded_channel::<Command>();
        let save_count = Arc::new(AtomicUsize::new(0));
        let save_count_task = save_count.clone();

        let handle = tokio::spawn(async move {
            let mut pending: Vec<CellChange> = Vec::new();
            // Track the last *successful* save so that a transient SQLite error
            // doesn't postpone future flush attempts indefinitely.
            let mut last_successful_save = Instant::now();
            let mut next_flush: Option<Instant> = None;

            loop {
                tokio::select! {
                    cmd = rx.recv() => {
                        match cmd {
                            Some(Command::Record(change)) => {
                                pending.push(change);

                                let now = Instant::now();
                                if now.duration_since(last_successful_save) >= config.max_delay {
                                    match flush_pending(&storage, &mut pending, &save_count_task) {
                                        Ok(()) => {
                                            last_successful_save = Instant::now();
                                            next_flush = None;
                                        }
                                        Err(_) => {
                                            // Keep pending changes and retry soon.
                                            next_flush = Some(Instant::now() + config.save_delay);
                                        }
                                    }
                                } else {
                                    next_flush = Some(now + config.save_delay);
                                }
                            }
                            Some(Command::Flush(reply)) => {
                                let result = flush_pending(&storage, &mut pending, &save_count_task);
                                if result.is_ok() {
                                    last_successful_save = Instant::now();
                                    next_flush = None;
                                } else {
                                    // Keep pending changes and schedule another attempt.
                                    next_flush = Some(Instant::now() + config.save_delay);
                                }
                                let _ = reply.send(result);
                            }
                            Some(Command::Shutdown(reply)) => {
                                let result = flush_pending(&storage, &mut pending, &save_count_task);
                                let _ = reply.send(result);
                                break;
                            }
                            None => {
                                let _ = flush_pending(&storage, &mut pending, &save_count_task);
                                break;
                            }
                        }
                    }
                    _ = async {
                        if let Some(deadline) = next_flush {
                            tokio::time::sleep_until(deadline).await;
                        }
                    }, if next_flush.is_some() => {
                        match flush_pending(&storage, &mut pending, &save_count_task) {
                            Ok(()) => {
                                last_successful_save = Instant::now();
                                next_flush = None;
                            }
                            Err(_) => {
                                // Keep pending changes and retry soon.
                                next_flush = Some(Instant::now() + config.save_delay);
                            }
                        }
                    }
                }
            }
        });

        Self {
            tx,
            save_count,
            handle,
        }
    }

    pub fn record_change(&self, change: CellChange) {
        // If the task has already exited we just drop the change; higher layers
        // should fall back to explicit persistence in that case.
        let _ = self.tx.send(Command::Record(change));
    }

    pub async fn flush(&self) -> StorageResult<()> {
        let (tx, rx) = oneshot::channel();
        let _ = self.tx.send(Command::Flush(tx));
        rx.await.unwrap_or(Ok(()))
    }

    pub async fn shutdown(self) -> StorageResult<()> {
        let (tx, rx) = oneshot::channel();
        let _ = self.tx.send(Command::Shutdown(tx));
        let result = rx.await.unwrap_or(Ok(()));
        let _ = self.handle.await;
        result
    }

    pub fn save_count(&self) -> usize {
        self.save_count.load(Ordering::Relaxed)
    }
}

fn flush_pending(
    storage: &Storage,
    pending: &mut Vec<CellChange>,
    save_count: &AtomicUsize,
) -> StorageResult<()> {
    if pending.is_empty() {
        return Ok(());
    }

    let changes = std::mem::take(pending);
    let result = storage.apply_cell_changes(&changes);
    if result.is_ok() {
        save_count.fetch_add(1, Ordering::Relaxed);
    } else {
        // If we failed, restore pending so we can retry on the next flush.
        *pending = changes;
    }
    result
}
