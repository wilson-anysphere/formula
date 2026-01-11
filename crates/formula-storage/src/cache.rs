use crate::storage::{CellChange, CellRange, Result as StorageResult, Storage};
use crate::types::{CellData, CellSnapshot, CellValue, SheetMeta, Style};
use lru::LruCache;
use std::collections::{HashMap, HashSet};
use std::num::NonZeroUsize;
use std::sync::{Arc, Mutex};
use uuid::Uuid;

const PAGE_BASE_OVERHEAD_BYTES: usize = 128;
const HASHMAP_ENTRY_OVERHEAD_BYTES: usize = 64;
const PENDING_CHANGE_OVERHEAD_BYTES: usize = 64;

#[derive(Debug, Clone)]
pub struct MemoryManagerConfig {
    /// Hard cap for the in-memory cache (default: 500MB).
    pub max_memory_bytes: usize,
    /// Max number of pages to cache regardless of memory (default: 4096).
    pub max_pages: usize,
    /// Evict pages when usage exceeds this fraction of the memory cap (default: 0.8).
    pub eviction_watermark: f64,
    /// Rows per cached page/tile (default: 256).
    pub rows_per_page: usize,
    /// Columns per cached page/tile (default: 256).
    pub cols_per_page: usize,
}

impl Default for MemoryManagerConfig {
    fn default() -> Self {
        Self {
            max_memory_bytes: 500 * 1024 * 1024,
            max_pages: 4096,
            eviction_watermark: 0.8,
            rows_per_page: 256,
            cols_per_page: 256,
        }
    }
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub struct MemoryManagerStats {
    pub page_hits: u64,
    pub page_misses: u64,
    pub pages_loaded: u64,
    pub pages_evicted: u64,
    pub flush_transactions: u64,
    pub pages_flushed: u64,
    pub cell_changes_flushed: u64,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FlushOutcome {
    /// Number of `CellChange`s applied to the live SQLite database.
    pub changes_applied: usize,
    /// Whether the underlying storage was persisted durably (e.g. flushed to an
    /// encrypted container on disk).
    pub persisted: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct PageKey {
    sheet_id: Uuid,
    page_row: i64,
    page_col: i64,
}

#[derive(Debug, Clone)]
pub struct ViewportData {
    range: CellRange,
    cells: HashMap<(i64, i64), CellSnapshot>,
}

impl ViewportData {
    pub fn range(&self) -> CellRange {
        self.range
    }

    pub fn get(&self, row: i64, col: i64) -> Option<&CellSnapshot> {
        self.cells.get(&(row, col))
    }

    pub fn non_empty_cells(&self) -> impl Iterator<Item = (&(i64, i64), &CellSnapshot)> {
        self.cells.iter()
    }

    pub fn into_cells(self) -> HashMap<(i64, i64), CellSnapshot> {
        self.cells
    }
}

#[derive(Debug)]
struct SequencedCellChange {
    seq: u64,
    change: CellChange,
}

#[derive(Debug)]
struct PageData {
    cells: HashMap<(i64, i64), CellSnapshot>,
    pending_changes: Vec<SequencedCellChange>,
    bytes: usize,
}

impl PageData {
    fn new_loaded(cells: HashMap<(i64, i64), CellSnapshot>) -> Self {
        let mut bytes = PAGE_BASE_OVERHEAD_BYTES;
        for snapshot in cells.values() {
            bytes = bytes.saturating_add(estimate_cell_snapshot_bytes(snapshot));
        }
        Self {
            cells,
            pending_changes: Vec::new(),
            bytes,
        }
    }

    fn is_dirty(&self) -> bool {
        !self.pending_changes.is_empty()
    }
}

struct Inner {
    pages: LruCache<PageKey, PageData>,
    sheet_meta: HashMap<Uuid, SheetMeta>,
    bytes: usize,
    next_change_seq: u64,
    stats: MemoryManagerStats,
    needs_persist: bool,
    dirty_pages: HashSet<PageKey>,
}

/// In-memory page cache with LRU eviction.
///
/// Pages are fixed-size tiles keyed by `(sheet_id, page_row, page_col)`. The
/// cache is populated by `load_viewport`, which loads any missing pages from
/// SQLite and returns a `ViewportData` snapshot for the requested range.
///
/// Edits recorded through `record_change` update the cached page immediately and
/// mark it dirty. Dirty pages are flushed to SQLite on eviction and via
/// autosave.
#[derive(Clone)]
pub struct MemoryManager {
    storage: Storage,
    config: MemoryManagerConfig,
    inner: Arc<Mutex<Inner>>,
}

impl MemoryManager {
    pub fn new(storage: Storage, mut config: MemoryManagerConfig) -> Self {
        config.max_pages = config.max_pages.max(1);
        config.rows_per_page = config.rows_per_page.max(1);
        config.cols_per_page = config.cols_per_page.max(1);
        if !config.eviction_watermark.is_finite() {
            config.eviction_watermark = 0.8;
        }
        let cap = NonZeroUsize::new(config.max_pages).expect("max_pages is non-zero");
        let inner = Inner {
            pages: LruCache::new(cap),
            sheet_meta: HashMap::new(),
            bytes: 0,
            next_change_seq: 0,
            stats: MemoryManagerStats::default(),
            needs_persist: false,
            dirty_pages: HashSet::new(),
        };
        Self {
            storage,
            config,
            inner: Arc::new(Mutex::new(inner)),
        }
    }

    pub fn estimated_usage_bytes(&self) -> usize {
        self.inner
            .lock()
            .expect("memory manager mutex poisoned")
            .bytes
    }

    pub fn cached_page_count(&self) -> usize {
        self.inner
            .lock()
            .expect("memory manager mutex poisoned")
            .pages
            .len()
    }

    pub fn stats_snapshot(&self) -> MemoryManagerStats {
        self.inner
            .lock()
            .expect("memory manager mutex poisoned")
            .stats
    }

    pub fn get_sheet(&self, sheet_id: Uuid) -> StorageResult<SheetMeta> {
        {
            let inner = self.inner.lock().expect("memory manager mutex poisoned");
            if let Some(meta) = inner.sheet_meta.get(&sheet_id) {
                return Ok(meta.clone());
            }
        }

        let meta = self.storage.get_sheet_meta(sheet_id)?;
        let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
        inner.sheet_meta.insert(sheet_id, meta.clone());
        Ok(meta)
    }

    /// Load a cell viewport (inclusive range) and return the cached values for
    /// just the requested range.
    pub fn load_viewport(&self, sheet_id: Uuid, viewport: CellRange) -> StorageResult<ViewportData> {
        self.get_sheet(sheet_id)?;

        let page_keys = self.page_keys_for_range(sheet_id, viewport);

        let mut cells = HashMap::new();
        let mut missing = Vec::new();

        let add_cells_in_viewport = |key: PageKey,
                                     page_cells: &HashMap<(i64, i64), CellSnapshot>,
                                     cells: &mut HashMap<(i64, i64), CellSnapshot>| {
            let page_range = self.page_range(key);
            let row_start = viewport.row_start.max(page_range.row_start);
            let row_end = viewport.row_end.min(page_range.row_end);
            let col_start = viewport.col_start.max(page_range.col_start);
            let col_end = viewport.col_end.min(page_range.col_end);

            if row_start > row_end || col_start > col_end {
                return;
            }

            let area = (row_end - row_start + 1)
                .saturating_mul(col_end - col_start + 1)
                .max(0) as usize;

            // Heuristic: for dense pages and large intersections, iterating the
            // `HashMap` and filtering is cheaper than hashing every coord. For
            // small intersections relative to the number of stored cells,
            // probing each coord wins.
            if area.saturating_mul(4) <= page_cells.len() {
                for row in row_start..=row_end {
                    for col in col_start..=col_end {
                        if let Some(snapshot) = page_cells.get(&(row, col)) {
                            cells.insert((row, col), snapshot.clone());
                        }
                    }
                }
            } else {
                for (&(row, col), snapshot) in page_cells {
                    if row >= row_start && row <= row_end && col >= col_start && col <= col_end {
                        cells.insert((row, col), snapshot.clone());
                    }
                }
            }
        };

        {
            let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
            let mut hit_pages = 0u64;
            let mut missed_pages = 0u64;
            for key in &page_keys {
                if let Some(page) = inner.pages.get(key) {
                    hit_pages = hit_pages.saturating_add(1);
                    add_cells_in_viewport(*key, &page.cells, &mut cells);
                } else {
                    missed_pages = missed_pages.saturating_add(1);
                    missing.push(*key);
                }
            }
            inner.stats.page_hits = inner.stats.page_hits.saturating_add(hit_pages);
            inner.stats.page_misses = inner.stats.page_misses.saturating_add(missed_pages);
        }

        for key in missing {
            let range = self.page_range(key);
            let loaded = self.storage.load_cells_in_range(sheet_id, range)?;
            let mut page_cells = HashMap::new();
            for (coord, snapshot) in loaded {
                page_cells.insert(coord, snapshot);
            }
            let page = PageData::new_loaded(page_cells);

            let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
            inner.stats.pages_loaded = inner.stats.pages_loaded.saturating_add(1);
            self.insert_page_locked(&mut inner, key, page)?;

            if let Some(page) = inner.pages.get(&key) {
                add_cells_in_viewport(key, &page.cells, &mut cells);
            }
        }

        {
            let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
            // Keep memory bounded after new page inserts.
            self.evict_if_needed_locked(&mut inner)?;
        }

        Ok(ViewportData { range: viewport, cells })
    }

    /// Load a visible cell range and return the sparse list of cells.
    ///
    /// This is retained for backwards compatibility; prefer `load_viewport`.
    pub fn load_visible_range(
        &self,
        sheet_id: Uuid,
        range: CellRange,
    ) -> StorageResult<Vec<((i64, i64), CellSnapshot)>> {
        let viewport = self.load_viewport(sheet_id, range)?;
        let mut cells: Vec<_> = viewport.into_cells().into_iter().collect();
        cells.sort_by_key(|(coord, _)| *coord);
        Ok(cells)
    }

    pub fn get_cached_cell(&self, sheet_id: Uuid, row: i64, col: i64) -> Option<CellSnapshot> {
        let key = self.page_key_for_cell(sheet_id, row, col);
        let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
        inner
            .pages
            .get(&key)
            .and_then(|page| page.cells.get(&(row, col)).cloned())
    }

    /// Record a change in-memory and mark the owning page as dirty.
    ///
    /// Changes are persisted on eviction and via autosave (`flush_dirty_pages`).
    pub fn record_change(&self, change: CellChange) -> StorageResult<()> {
        self.get_sheet(change.sheet_id)?;

        let key = self.page_key_for_cell(change.sheet_id, change.row, change.col);
        let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
        if inner.pages.contains(&key) {
            inner.stats.page_hits = inner.stats.page_hits.saturating_add(1);
            self.apply_change_to_page_locked(&mut inner, key, change)?;
            self.evict_if_needed_locked(&mut inner)?;
            return Ok(());
        }

        inner.stats.page_misses = inner.stats.page_misses.saturating_add(1);
        drop(inner);

        let range = self.page_range(key);
        let loaded = self.storage.load_cells_in_range(change.sheet_id, range)?;
        let mut cells = HashMap::new();
        for (coord, snapshot) in loaded {
            cells.insert(coord, snapshot);
        }
        let page = PageData::new_loaded(cells);

        let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
        if !inner.pages.contains(&key) {
            inner.stats.pages_loaded = inner.stats.pages_loaded.saturating_add(1);
            self.insert_page_locked(&mut inner, key, page)?;
        }

        // If another thread inserted it while we were loading, we ignore the
        // locally loaded copy and apply the edit to the cached page.
        self.apply_change_to_page_locked(&mut inner, key, change)?;
        self.evict_if_needed_locked(&mut inner)?;
        Ok(())
    }

    fn apply_change_to_page_locked(
        &self,
        inner: &mut Inner,
        key: PageKey,
        change: CellChange,
    ) -> StorageResult<()> {
        if !inner.pages.contains(&key) {
            // Callers must ensure the page exists (either already cached or
            // inserted after loading from storage).
            let range = self.page_range(key);
            let loaded = self.storage.load_cells_in_range(change.sheet_id, range)?;
            let mut cells = HashMap::new();
            for (coord, snapshot) in loaded {
                cells.insert(coord, snapshot);
            }
            let page = PageData::new_loaded(cells);
            inner.stats.pages_loaded = inner.stats.pages_loaded.saturating_add(1);
            self.insert_page_locked(inner, key, page)?;
        }

        let seq = inner.next_change_seq;
        inner.next_change_seq = inner.next_change_seq.saturating_add(1);

        let page = inner
            .pages
            .get_mut(&key)
            .expect("page present after insert");

        let before_page_bytes = page.bytes;
        let was_clean = page.pending_changes.is_empty();
        // Update cached snapshot.
        if let Some(existing) = page.cells.get(&(change.row, change.col)) {
            page.bytes = page
                .bytes
                .saturating_sub(estimate_cell_snapshot_bytes(existing));
        }

        if change.data.is_truly_empty() {
            page.cells.remove(&(change.row, change.col));
        } else {
            page.cells.insert(
                (change.row, change.col),
                CellSnapshot {
                    value: change.data.value.clone(),
                    formula: change.data.formula.clone(),
                    style_id: None,
                },
            );
        }

        if let Some(updated) = page.cells.get(&(change.row, change.col)) {
            page.bytes = page
                .bytes
                .saturating_add(estimate_cell_snapshot_bytes(updated));
        }

        page.bytes = page.bytes.saturating_add(estimate_cell_change_bytes(&change));
        page.pending_changes.push(SequencedCellChange { seq, change });
        if was_clean {
            inner.dirty_pages.insert(key);
        }

        inner.bytes = inner
            .bytes
            .saturating_sub(before_page_bytes)
            .saturating_add(page.bytes);
        Ok(())
    }

    /// Flush all dirty pages to SQLite. Returns the number of flushed cell
    /// changes and whether the storage was durably persisted.
    pub fn flush_dirty_pages(&self) -> StorageResult<FlushOutcome> {
        let mut inner = self.inner.lock().expect("memory manager mutex poisoned");
        self.flush_pending_changes_upto_seq_locked(&mut inner, u64::MAX)
    }

    fn page_key_for_cell(&self, sheet_id: Uuid, row: i64, col: i64) -> PageKey {
        let rows = self.config.rows_per_page as i64;
        let cols = self.config.cols_per_page as i64;
        PageKey {
            sheet_id,
            page_row: row.div_euclid(rows),
            page_col: col.div_euclid(cols),
        }
    }

    fn page_range(&self, key: PageKey) -> CellRange {
        let rows = self.config.rows_per_page as i64;
        let cols = self.config.cols_per_page as i64;
        let row_start = key.page_row.saturating_mul(rows);
        let col_start = key.page_col.saturating_mul(cols);
        CellRange::new(
            row_start,
            row_start.saturating_add(rows.saturating_sub(1)),
            col_start,
            col_start.saturating_add(cols.saturating_sub(1)),
        )
    }

    fn page_keys_for_range(&self, sheet_id: Uuid, range: CellRange) -> Vec<PageKey> {
        let rows = self.config.rows_per_page as i64;
        let cols = self.config.cols_per_page as i64;

        let page_row_start = range.row_start.div_euclid(rows);
        let page_row_end = range.row_end.div_euclid(rows);
        let page_col_start = range.col_start.div_euclid(cols);
        let page_col_end = range.col_end.div_euclid(cols);

        let mut keys = Vec::new();
        for page_row in page_row_start..=page_row_end {
            for page_col in page_col_start..=page_col_end {
                keys.push(PageKey {
                    sheet_id,
                    page_row,
                    page_col,
                });
            }
        }
        keys
    }

    fn insert_page_locked(
        &self,
        inner: &mut Inner,
        key: PageKey,
        page: PageData,
    ) -> StorageResult<()> {
        // If another thread already inserted this page, do not merge the newly
        // loaded copy. We always load *full* pages, so the existing page already
        // contains a complete view plus any in-memory edits. Merging can also
        // resurrect stale data if the load raced with a flush.
        if inner.pages.contains(&key) {
            return Ok(());
        }

        // Respect the explicit page cap without relying on `LruCache::put`'s
        // implicit eviction (which does not expose the key).
        while inner.pages.len() >= self.config.max_pages {
            self.evict_one_page_locked(inner)?;
        }

        let page_bytes = page.bytes;
        if let Some(evicted) = inner.pages.put(key, page) {
            inner.bytes = inner.bytes.saturating_sub(evicted.bytes);
            inner.stats.pages_evicted = inner.stats.pages_evicted.saturating_add(1);
        }
        inner.bytes = inner.bytes.saturating_add(page_bytes);
        Ok(())
    }

    fn eviction_limit_bytes(&self) -> usize {
        let watermark = self.config.eviction_watermark.clamp(0.0, 1.0);
        (self.config.max_memory_bytes as f64 * watermark) as usize
    }

    fn evict_if_needed_locked(&self, inner: &mut Inner) -> StorageResult<()> {
        let limit = self.eviction_limit_bytes();
        while inner.bytes > limit && !inner.pages.is_empty() {
            if let Some((_lru_key, lru_page)) = inner.pages.peek_lru() {
                if lru_page.is_dirty() {
                    let max_seq = lru_page
                        .pending_changes
                        .iter()
                        .map(|c| c.seq)
                        .max()
                        .unwrap_or(0);
                    // Flush all changes up through the oldest page's latest change to
                    // preserve global ordering across pages.
                    self.flush_pending_changes_upto_seq_locked(inner, max_seq)?;
                }
            }

            if inner.bytes <= limit {
                break;
            }

            if let Some((key, page)) = inner.pages.pop_lru() {
                inner.dirty_pages.remove(&key);
                inner.stats.pages_evicted = inner.stats.pages_evicted.saturating_add(1);
                inner.bytes = inner.bytes.saturating_sub(page.bytes);
            }
        }
        Ok(())
    }

    fn evict_one_page_locked(&self, inner: &mut Inner) -> StorageResult<()> {
        let Some((_lru_key, lru_page)) = inner.pages.peek_lru() else {
            return Ok(());
        };

        if lru_page.is_dirty() {
            let max_seq = lru_page
                .pending_changes
                .iter()
                .map(|c| c.seq)
                .max()
                .unwrap_or(0);
            // Flush all changes up through the evicted page's latest change to
            // preserve global ordering across pages.
            self.flush_pending_changes_upto_seq_locked(inner, max_seq)?;
        }

        if let Some((key, page)) = inner.pages.pop_lru() {
            inner.dirty_pages.remove(&key);
            inner.stats.pages_evicted = inner.stats.pages_evicted.saturating_add(1);
            inner.bytes = inner.bytes.saturating_sub(page.bytes);
        }
        Ok(())
    }

    fn flush_pending_changes_upto_seq_locked(
        &self,
        inner: &mut Inner,
        upto_seq: u64,
    ) -> StorageResult<FlushOutcome> {
        let mut flushed: Vec<(u64, CellChange)> = Vec::new();
        let mut pages_flushed = 0u64;

        let dirty_keys: Vec<PageKey> = inner.dirty_pages.iter().copied().collect();
        for key in dirty_keys {
            let Some(page) = inner.pages.peek_mut(&key) else {
                inner.dirty_pages.remove(&key);
                continue;
            };

            if page.pending_changes.is_empty() {
                inner.dirty_pages.remove(&key);
                continue;
            }

            let mut keep = Vec::new();
            let mut flushed_any = false;
            let original = std::mem::take(&mut page.pending_changes);
            for sc in original {
                if sc.seq <= upto_seq {
                    page.bytes = page
                        .bytes
                        .saturating_sub(estimate_cell_change_bytes(&sc.change));
                    inner.bytes = inner
                        .bytes
                        .saturating_sub(estimate_cell_change_bytes(&sc.change));
                    flushed.push((sc.seq, sc.change));
                    flushed_any = true;
                } else {
                    keep.push(sc);
                }
            }
            page.pending_changes = keep;
            if page.pending_changes.is_empty() {
                inner.dirty_pages.remove(&key);
            }
            if flushed_any {
                pages_flushed += 1;
            }
        }

        if flushed.is_empty() {
            // Even without cell changes, encrypted storages may still require a
            // persist after a previous successful flush.
            if inner.needs_persist {
                self.storage.persist()?;
                inner.needs_persist = false;
                return Ok(FlushOutcome {
                    changes_applied: 0,
                    persisted: true,
                });
            }
            return Ok(FlushOutcome {
                changes_applied: 0,
                persisted: false,
            });
        }

        flushed.sort_by_key(|(seq, _)| *seq);
        let (seqs, mut changes): (Vec<u64>, Vec<CellChange>) = flushed
            .into_iter()
            .map(|(seq, change)| (seq, change))
            .unzip();

        let result = self.storage.apply_cell_changes(&changes);
        match result {
            Ok(()) => {
                inner.stats.flush_transactions = inner.stats.flush_transactions.saturating_add(1);
                inner.stats.pages_flushed = inner.stats.pages_flushed.saturating_add(pages_flushed);
                inner.stats.cell_changes_flushed = inner
                    .stats
                    .cell_changes_flushed
                    .saturating_add(changes.len() as u64);
                inner.needs_persist = true;

                self.storage.persist()?;
                inner.needs_persist = false;

                Ok(FlushOutcome {
                    changes_applied: changes.len(),
                    persisted: true,
                })
            }
            Err(err) => {
                // Restore pending changes to their pages (prepend, since all restored
                // changes have smaller seqs than the kept ones).
                let mut restore: HashMap<PageKey, Vec<SequencedCellChange>> = HashMap::new();
                for (seq, change) in seqs.into_iter().zip(changes.drain(..)) {
                    let key = self.page_key_for_cell(change.sheet_id, change.row, change.col);
                    restore
                        .entry(key)
                        .or_default()
                        .push(SequencedCellChange { seq, change });
                }

                for (key, page) in inner.pages.iter_mut() {
                    if let Some(mut restored) = restore.remove(key) {
                        for sc in &restored {
                            page.bytes = page
                                .bytes
                                .saturating_add(estimate_cell_change_bytes(&sc.change));
                            inner.bytes = inner
                                .bytes
                                .saturating_add(estimate_cell_change_bytes(&sc.change));
                        }
                        restored.extend(std::mem::take(&mut page.pending_changes));
                        page.pending_changes = restored;
                        inner.dirty_pages.insert(*key);
                    }
                }
                Err(err)
            }
        }
    }
}

fn estimate_cell_snapshot_bytes(snapshot: &CellSnapshot) -> usize {
    HASHMAP_ENTRY_OVERHEAD_BYTES
        .saturating_add(estimate_cell_value_bytes(&snapshot.value))
        .saturating_add(snapshot.formula.as_ref().map(|s| s.len()).unwrap_or(0))
        .saturating_add(snapshot.style_id.map(|_| 8).unwrap_or(0))
}

fn estimate_cell_change_bytes(change: &CellChange) -> usize {
    let mut bytes = PENDING_CHANGE_OVERHEAD_BYTES;
    bytes = bytes.saturating_add(estimate_cell_data_bytes(&change.data));
    if let Some(user_id) = &change.user_id {
        bytes = bytes.saturating_add(user_id.len());
    }
    bytes
}

fn estimate_cell_data_bytes(data: &CellData) -> usize {
    let mut bytes = estimate_cell_value_bytes(&data.value);
    if let Some(formula) = &data.formula {
        bytes = bytes.saturating_add(formula.len());
    }
    if let Some(style) = &data.style {
        bytes = bytes.saturating_add(estimate_style_bytes(style));
    }
    bytes
}

fn estimate_style_bytes(style: &Style) -> usize {
    let mut bytes = 0usize;
    if let Some(fmt) = &style.number_format {
        bytes = bytes.saturating_add(fmt.len());
    }
    if let Some(alignment) = &style.alignment {
        bytes = bytes.saturating_add(alignment.to_string().len());
    }
    if let Some(protection) = &style.protection {
        bytes = bytes.saturating_add(protection.to_string().len());
    }
    bytes
}

fn estimate_cell_value_bytes(value: &CellValue) -> usize {
    match value {
        CellValue::String(s) => s.len(),
        CellValue::RichText(rt) => rt.text.len().saturating_add(rt.runs.len().saturating_mul(32)),
        CellValue::Error(err) => err.as_str().len(),
        CellValue::Array(arr) => {
            let mut bytes = 64usize;
            for row in &arr.data {
                bytes = bytes.saturating_add(16);
                for v in row {
                    bytes = bytes.saturating_add(match v {
                        CellValue::String(s) => s.len(),
                        CellValue::RichText(rt) => rt.text.len() + rt.runs.len() * 32,
                        CellValue::Error(err) => err.as_str().len(),
                        CellValue::Array(_) => 64,
                        CellValue::Spill(_) => 16,
                        CellValue::Number(_) | CellValue::Boolean(_) | CellValue::Empty => 0,
                    });
                }
            }
            bytes
        }
        CellValue::Spill(_) | CellValue::Number(_) | CellValue::Boolean(_) | CellValue::Empty => 0,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn insert_page_skips_stale_reload_after_flush() {
        let storage = Storage::open_in_memory().expect("open storage");
        let workbook = storage
            .create_workbook("Book", None)
            .expect("create workbook");
        let sheet = storage
            .create_sheet(workbook.id, "Sheet", 0, None)
            .expect("create sheet");

        storage
            .apply_cell_changes(&[CellChange {
                sheet_id: sheet.id,
                row: 0,
                col: 0,
                data: CellData {
                    value: CellValue::Number(1.0),
                    formula: None,
                    style: None,
                },
                user_id: None,
            }])
            .expect("seed cell");

        let memory = MemoryManager::new(
            storage.clone(),
            MemoryManagerConfig {
                max_memory_bytes: 1024 * 1024,
                max_pages: 64,
                eviction_watermark: 1.0,
                rows_per_page: 64,
                cols_per_page: 64,
            },
        );

        // Load the page, then delete a cell so we have a pending tombstone.
        memory
            .load_viewport(sheet.id, CellRange::new(0, 0, 0, 0))
            .expect("load viewport");
        let stale_snapshot = memory
            .get_cached_cell(sheet.id, 0, 0)
            .expect("cached snapshot exists");
        memory
            .record_change(CellChange {
                sheet_id: sheet.id,
                row: 0,
                col: 0,
                data: CellData::empty(),
                user_id: None,
            })
            .expect("record deletion");

        // Flush the deletion, then simulate a concurrent/stale page load that
        // still contained the old value.
        memory.flush_dirty_pages().expect("flush");
        let key = memory.page_key_for_cell(sheet.id, 0, 0);
        let mut cells = HashMap::new();
        cells.insert((0, 0), stale_snapshot);
        let page = PageData::new_loaded(cells);

        {
            let mut inner = memory.inner.lock().expect("memory manager mutex poisoned");
            memory
                .insert_page_locked(&mut inner, key, page)
                .expect("insert");

            let page = inner.pages.get(&key).expect("page cached");
            assert!(
                !page.cells.contains_key(&(0, 0)),
                "stale reload should not resurrect deleted cells"
            );
            assert!(page.pending_changes.is_empty(), "page should be clean after flush");
        }
    }
}
