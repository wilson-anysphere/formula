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

mod autosave;
mod cache;
pub mod encryption;
mod schema;
pub mod storage;
mod types;

pub use autosave::{AutoSaveConfig, AutoSaveManager};
pub use cache::{FlushOutcome, MemoryManager, MemoryManagerConfig, MemoryManagerStats, ViewportData};
pub use encryption::{EncryptionError, InMemoryKeyProvider, KeyProvider, KeyProviderError, KeyRing};
pub use storage::{Storage, StorageError};
pub use types::{
    CellData, CellSnapshot, CellValue, ImportModelWorkbookOptions, NamedRange, SheetMeta,
    SheetVisibility, Style, WorkbookMeta,
};

pub use storage::{CellChange, CellRange};
