use crate::cache::MemoryManager;
use crate::storage::{CellChange, Result as StorageResult};
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
    Touch,
    Flush(oneshot::Sender<StorageResult<()>>),
    Shutdown(oneshot::Sender<StorageResult<()>>),
}

/// Debounced, batched persistence for cell edits.
///
/// The autosave manager drives `MemoryManager::flush_dirty_pages` on a debounce
/// timer. It does not persist each edit immediately, but guarantees that dirty
/// pages will be flushed (and persisted) within `max_delay`.
pub struct AutoSaveManager {
    memory: MemoryManager,
    tx: mpsc::UnboundedSender<Command>,
    save_count: Arc<AtomicUsize>,
    handle: JoinHandle<()>,
}

impl AutoSaveManager {
    pub fn spawn(memory: MemoryManager, config: AutoSaveConfig) -> Self {
        let (tx, mut rx) = mpsc::unbounded_channel::<Command>();
        let save_count = Arc::new(AtomicUsize::new(0));
        let save_count_task = save_count.clone();
        let memory_task = memory.clone();

        let handle = tokio::spawn(async move {
            // Track the last *successful* save so that a transient SQLite error
            // doesn't postpone future flush attempts indefinitely.
            let mut last_successful_save = Instant::now();
            let mut next_flush: Option<Instant> = None;

            loop {
                tokio::select! {
                    cmd = rx.recv() => {
                        match cmd {
                            Some(Command::Touch) => {
                                let now = Instant::now();
                                if now.duration_since(last_successful_save) >= config.max_delay {
                                    match flush_dirty_pages(&memory_task, &save_count_task) {
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
                                let result = flush_dirty_pages(&memory_task, &save_count_task);
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
                                let result = flush_dirty_pages(&memory_task, &save_count_task);
                                let _ = reply.send(result);
                                break;
                            }
                            None => {
                                let _ = flush_dirty_pages(&memory_task, &save_count_task);
                                break;
                            }
                        }
                    }
                    _ = async {
                        if let Some(deadline) = next_flush {
                            tokio::time::sleep_until(deadline).await;
                        }
                    }, if next_flush.is_some() => {
                        match flush_dirty_pages(&memory_task, &save_count_task) {
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
            memory,
            tx,
            save_count,
            handle,
        }
    }

    pub fn record_change(&self, change: CellChange) -> StorageResult<()> {
        self.memory.record_change(change)?;
        // If the task has already exited we just drop the wake-up signal; higher
        // layers can always fall back to explicit persistence.
        if self.tx.send(Command::Touch).is_err() {
            // The autosave task is gone; fall back to synchronous flushing so
            // callers still get persistence guarantees.
            flush_dirty_pages(&self.memory, &self.save_count)?;
        }
        Ok(())
    }

    pub fn notify_change(&self) {
        if self.tx.send(Command::Touch).is_err() {
            let _ = flush_dirty_pages(&self.memory, &self.save_count);
        }
    }

    pub async fn flush(&self) -> StorageResult<()> {
        let (tx, rx) = oneshot::channel();
        if self.tx.send(Command::Flush(tx)).is_err() {
            flush_dirty_pages(&self.memory, &self.save_count)?;
            return Ok(());
        }
        rx.await.unwrap_or_else(|_| flush_dirty_pages(&self.memory, &self.save_count))
    }

    pub async fn shutdown(self) -> StorageResult<()> {
        let (tx, rx) = oneshot::channel();
        let result = if self.tx.send(Command::Shutdown(tx)).is_err() {
            flush_dirty_pages(&self.memory, &self.save_count)
        } else {
            rx.await
                .unwrap_or_else(|_| flush_dirty_pages(&self.memory, &self.save_count))
        };
        let _ = self.handle.await;
        result
    }

    pub fn save_count(&self) -> usize {
        self.save_count.load(Ordering::Relaxed)
    }
}

fn flush_dirty_pages(memory: &MemoryManager, save_count: &AtomicUsize) -> StorageResult<()> {
    let outcome = memory.flush_dirty_pages()?;
    if outcome.persisted {
        save_count.fetch_add(1, Ordering::Relaxed);
    }
    Ok(())
}
