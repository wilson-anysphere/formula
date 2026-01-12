//! SQLite-backed storage for Formula workbooks.
//!
//! This crate is intentionally self-contained so it can be integrated into the
//! Tauri backend later. It follows the schema described in
//! `docs/04-data-model-storage.md` and exposes:
//! - SQLite schema creation/migration
//! - Workbook/sheet metadata operations
//! - Lazy cell range loading
//! - Batched transactional writes
//! - Autosave debouncing
//! - A viewport-driven, page-based LRU cache for memory management
//! - Optional encryption-at-rest for persisted workbooks (AES-256-GCM)
//!
//! For large sheets, prefer [`MemoryManager::load_viewport`] over full-range
//! scans. It pages in fixed-size tiles (configured via [`MemoryManagerConfig`])
//! and keeps the cache bounded with LRU eviction + dirty-page writeback.
//! For smoother scrolling, [`MemoryManager::load_viewport_with_margin`] can be
//! used to prefetch pages around the visible viewport.

mod autosave;
mod cache;
pub mod data_model;
pub mod encryption;
mod schema;
pub mod storage;
mod types;

pub use autosave::{AutoSaveConfig, AutoSaveManager};
pub use cache::{
    FlushOutcome, MemoryManager, MemoryManagerConfig, MemoryManagerMetrics, MemoryManagerStats,
    ViewportData,
};
pub use encryption::{EncryptionError, InMemoryKeyProvider, KeyProvider, KeyProviderError, KeyRing};
pub use storage::{Storage, StorageError};
pub use types::{
    CellData, CellSnapshot, CellValue, ImportModelWorkbookOptions, NamedRange, SheetMeta,
    SheetVisibility, Style, WorkbookMeta,
};

pub use storage::{CellChange, CellRange};

pub(crate) fn lock_unpoisoned<T>(mutex: &std::sync::Mutex<T>) -> std::sync::MutexGuard<'_, T> {
    mutex.lock().unwrap_or_else(|poisoned| poisoned.into_inner())
}
