use crate::storage::{CellChange, CellRange, Result as StorageResult, Storage};
use crate::types::{CellSnapshot, CellValue, SheetMeta};
use lru::LruCache;
use std::collections::HashMap;
use std::num::NonZeroUsize;
use std::sync::Mutex;
use uuid::Uuid;

#[derive(Debug, Clone)]
pub struct MemoryManagerConfig {
    /// Hard cap for the in-memory cache (default: 500MB).
    pub max_memory_bytes: usize,
    /// Max number of sheets to cache regardless of memory (default: 32).
    pub max_sheets: usize,
    /// Evict sheets when usage exceeds this fraction of the memory cap (default: 0.8).
    pub eviction_watermark: f64,
}

impl Default for MemoryManagerConfig {
    fn default() -> Self {
        Self {
            max_memory_bytes: 500 * 1024 * 1024,
            max_sheets: 32,
            eviction_watermark: 0.8,
        }
    }
}

#[derive(Debug, Clone)]
pub struct SheetData {
    pub meta: SheetMeta,
    pub cells: HashMap<(i64, i64), CellSnapshot>,
    pub is_dirty: bool,
    pending_changes: Vec<CellChange>,
}

impl SheetData {
    fn estimated_bytes(&self) -> usize {
        let mut bytes = 0usize;

        // Sheet metadata (names, etc).
        bytes += self.meta.name.len();
        if let Some(tab_color) = &self.meta.tab_color {
            bytes += tab_color.len();
        }
        if let Some(rel_id) = &self.meta.xlsx_rel_id {
            bytes += rel_id.len();
        }
        if let Some(metadata) = &self.meta.metadata {
            bytes += metadata.to_string().len();
        }

        // Each cell entry has HashMap overhead; we use a conservative constant.
        for snapshot in self.cells.values() {
            bytes += 64;
            match &snapshot.value {
                CellValue::String(s) => bytes += s.len(),
                CellValue::RichText(rt) => {
                    bytes += rt.text.len();
                    bytes += rt.runs.len() * 32;
                }
                CellValue::Error(err) => bytes += err.as_str().len(),
                CellValue::Array(arr) => {
                    // Rough estimate: nested cell values.
                    bytes += 64;
                    for row in &arr.data {
                        bytes += 16;
                        for v in row {
                            bytes += match v {
                                CellValue::String(s) => s.len(),
                                CellValue::RichText(rt) => rt.text.len() + rt.runs.len() * 32,
                                CellValue::Error(err) => err.as_str().len(),
                                CellValue::Array(_) => 64,
                                CellValue::Spill(_) => 16,
                                CellValue::Number(_) | CellValue::Boolean(_) | CellValue::Empty => {
                                    0
                                }
                            };
                        }
                    }
                }
                CellValue::Spill(_)
                | CellValue::Number(_)
                | CellValue::Boolean(_)
                | CellValue::Empty => {}
            }
            if let Some(formula) = &snapshot.formula {
                bytes += formula.len();
            }
        }

        bytes
    }
}

struct Inner {
    cache: LruCache<Uuid, SheetData>,
    bytes: usize,
}

/// In-memory sheet cache with LRU eviction.
///
/// This cache keeps *loaded* cell data in memory, but uses the underlying
/// `Storage` for lazy range loading. It is intentionally conservative and does
/// not attempt to be a full on-demand paging system yet.
pub struct MemoryManager {
    storage: Storage,
    config: MemoryManagerConfig,
    inner: Mutex<Inner>,
}

impl MemoryManager {
    pub fn new(storage: Storage, config: MemoryManagerConfig) -> Self {
        let cap = NonZeroUsize::new(config.max_sheets.max(1)).expect("max_sheets is non-zero");
        let inner = Inner {
            cache: LruCache::new(cap),
            bytes: 0,
        };
        Self {
            storage,
            config,
            inner: Mutex::new(inner),
        }
    }

    pub fn estimated_usage_bytes(&self) -> usize {
        self.inner
            .lock()
            .expect("memory manager mutex poisoned")
            .bytes
    }

    pub fn cached_sheet_count(&self) -> usize {
        self.inner
            .lock()
            .expect("memory manager mutex poisoned")
            .cache
            .len()
    }

    pub fn get_sheet(&self, sheet_id: Uuid) -> StorageResult<SheetMeta> {
        {
            let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
            if let Some(sheet) = inner.cache.get(&sheet_id) {
                return Ok(sheet.meta.clone());
            }
        }

        let meta = self.storage.get_sheet_meta(sheet_id)?;
        self.insert_sheet_if_missing(meta.clone())?;
        Ok(meta)
    }

    /// Load a visible cell range from SQLite and update the in-memory cache.
    pub fn load_visible_range(
        &self,
        sheet_id: Uuid,
        range: CellRange,
    ) -> StorageResult<Vec<((i64, i64), CellSnapshot)>> {
        self.ensure_sheet(sheet_id)?;

        let cells = self.storage.load_cells_in_range(sheet_id, range)?;

        let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
        if let Some(sheet) = inner.cache.get_mut(&sheet_id) {
            let before = sheet.estimated_bytes();
            for (coord, snapshot) in &cells {
                sheet.cells.insert(*coord, snapshot.clone());
            }
            let after = sheet.estimated_bytes();
            inner.bytes = inner.bytes.saturating_sub(before).saturating_add(after);
        }

        self.evict_if_needed(&mut inner)?;
        Ok(cells)
    }

    pub fn get_cached_cell(&self, sheet_id: Uuid, row: i64, col: i64) -> Option<CellSnapshot> {
        let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
        inner
            .cache
            .get(&sheet_id)
            .and_then(|sheet| sheet.cells.get(&(row, col)).cloned())
    }

    /// Record a change in-memory and mark the sheet as dirty.
    ///
    /// This does **not** automatically persist the change; callers can either
    /// flush explicitly or rely on eviction/close logic to persist pending
    /// changes.
    pub fn record_change(&self, change: CellChange) -> StorageResult<()> {
        self.ensure_sheet(change.sheet_id)?;

        let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
        if let Some(sheet) = inner.cache.get_mut(&change.sheet_id) {
            let before = sheet.estimated_bytes();

            // Update the in-memory view so reads reflect the user's edits.
            if change.data.is_truly_empty() {
                sheet.cells.remove(&(change.row, change.col));
            } else {
                sheet.cells.insert(
                    (change.row, change.col),
                    CellSnapshot {
                        value: change.data.value.clone(),
                        formula: change.data.formula.clone(),
                        style_id: None,
                    },
                );
            }

            sheet.pending_changes.push(change);
            sheet.is_dirty = true;

            let after = sheet.estimated_bytes();
            inner.bytes = inner.bytes.saturating_sub(before).saturating_add(after);
        }

        self.evict_if_needed(&mut inner)?;
        Ok(())
    }

    fn ensure_sheet(&self, sheet_id: Uuid) -> StorageResult<()> {
        {
            let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
            if inner.cache.contains(&sheet_id) {
                // Mark as recently used.
                inner.cache.get(&sheet_id);
                return Ok(());
            }
        }

        let meta = self.storage.get_sheet_meta(sheet_id)?;
        self.insert_sheet_if_missing(meta)?;
        Ok(())
    }

    fn insert_sheet_if_missing(&self, meta: SheetMeta) -> StorageResult<()> {
        let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
        if inner.cache.contains(&meta.id) {
            return Ok(());
        }

        let data = SheetData {
            meta: meta.clone(),
            cells: HashMap::new(),
            is_dirty: false,
            pending_changes: Vec::new(),
        };
        let bytes = data.estimated_bytes();

        if let Some(evicted) = inner.cache.put(meta.id, data) {
            inner.bytes = inner.bytes.saturating_sub(evicted.estimated_bytes());
        }
        inner.bytes = inner.bytes.saturating_add(bytes);

        self.evict_if_needed(&mut inner)?;
        Ok(())
    }

    fn evict_if_needed(&self, inner: &mut Inner) -> StorageResult<()> {
        let limit = (self.config.max_memory_bytes as f64 * self.config.eviction_watermark) as usize;
        while inner.bytes > limit && inner.cache.len() > 1 {
            if let Some((_sheet_id, sheet)) = inner.cache.pop_lru() {
                let sheet_bytes = sheet.estimated_bytes();

                if sheet.is_dirty && !sheet.pending_changes.is_empty() {
                    // Flush before evicting to avoid losing edits.
                    if let Err(err) = self.storage.apply_cell_changes(&sheet.pending_changes) {
                        // Failed to persist; keep the sheet in cache.
                        inner.cache.put(sheet.meta.id, sheet);
                        return Err(err);
                    }
                }

                inner.bytes = inner.bytes.saturating_sub(sheet_bytes);
            } else {
                break;
            }
        }
        Ok(())
    }
}
