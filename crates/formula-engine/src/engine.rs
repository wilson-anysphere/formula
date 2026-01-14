use crate::bytecode;
use crate::calc_settings::{CalcSettings, CalculationMode};
use crate::date::ExcelDateSystem;
use crate::editing::rewrite::{
    rewrite_formula_for_copy_delta, rewrite_formula_for_range_map_with_resolver,
    rewrite_formula_for_sheet_delete_with_aliases,
    rewrite_formula_for_structural_edit_with_resolver, GridRange, RangeMapEdit, StructuralEdit,
};
use crate::editing::{
    CellChange, CellSnapshot, EditError, EditOp, EditResult, FormulaRewrite, MovedRange,
};
use crate::eval::{
    compile_canonical_expr, lower_ast, parse_a1, CellAddr, CompiledExpr, Expr, FormulaParseError,
    RangeRef, SheetReference, ValueResolver,
};
use crate::graph::{CellDeps, DependencyGraph as CalcGraph, Precedent, SheetRange};
use crate::iterative;
use crate::locale::{
    canonicalize_formula, canonicalize_formula_with_style, localize_formula,
    localize_formula_with_style, FormulaLocale, ValueLocaleConfig,
};
use crate::metadata::style_id_for_row_in_runs;
pub use crate::metadata::FormatRun;
use crate::pivot::{
    refresh_pivot, PivotRefreshContext, PivotRefreshError, PivotRefreshOutput, PivotSource,
    PivotTableDefinition, PivotTableId,
};
use crate::value::{Array, ErrorKind, Value};
use formula_format::{
    DateSystem as FmtDateSystem, FormatOptions as FmtFormatOptions, Value as FmtValue,
};
use formula_model::table::TableColumn;
use formula_model::{
    rewrite_table_names_in_formula, validate_table_name, CellId, CellRef, ColProperties,
    HorizontalAlignment, Range, RowProperties, Style, StyleTable, Table, TableError,
    EXCEL_MAX_COLS, EXCEL_MAX_ROWS,
};
#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
use rayon::{prelude::*, ThreadPool, ThreadPoolBuilder};
use std::cell::RefCell;
use std::cmp::{max, Ordering};
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet, VecDeque};
#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
use std::sync::OnceLock;
use std::sync::{Arc, Mutex};
use thiserror::Error;
use unicode_normalization::UnicodeNormalization;

mod bytecode_diagnostics;
mod pivot_refresh;
pub use bytecode_diagnostics::{
    BytecodeCompileReason, BytecodeCompileReportEntry, BytecodeCompileStats,
};

pub type SheetId = usize;

/// Host-provided workbook / system metadata surfaced via the Excel `INFO()` worksheet function.
///
/// The engine does not currently read live OS state at runtime (to keep evaluation deterministic
/// and portable). Hosts may populate these fields explicitly if they want `INFO()` to return
/// Excel-like values.
#[derive(Debug, Default, Clone, PartialEq)]
pub struct EngineInfo {
    /// `INFO("system")` override.
    ///
    /// When unset, `INFO("system")` defaults to `"pcdos"` for backward compatibility.
    pub system: Option<String>,
    /// `INFO("directory")`.
    pub directory: Option<String>,
    /// `INFO("osversion")`.
    pub osversion: Option<String>,
    /// `INFO("release")`.
    pub release: Option<String>,
    /// `INFO("version")`.
    pub version: Option<String>,
    /// `INFO("memavail")`.
    pub memavail: Option<f64>,
    /// `INFO("totmem")`.
    pub totmem: Option<f64>,
    /// Workbook-level fallback for `INFO("origin")`.
    pub origin: Option<String>,
    /// Per-sheet override for `INFO("origin")`, keyed by internal sheet id.
    pub origin_by_sheet: HashMap<SheetId, String>,
}

#[derive(Debug, Error)]
pub enum EngineError {
    #[error(transparent)]
    Address(#[from] crate::eval::AddressParseError),
    #[error(transparent)]
    Parse(#[from] FormulaParseError),
    #[error(transparent)]
    AstParse(#[from] crate::ParseError),
    #[error(transparent)]
    AstSerialize(#[from] crate::SerializeError),
    #[error("cannot delete last sheet")]
    CannotDeleteLastSheet,
    #[error(
        "range values dimensions mismatch: expected {expected_rows}x{expected_cols}, got {actual_rows}x{actual_cols}"
    )]
    RangeValuesDimensionMismatch {
        expected_rows: usize,
        expected_cols: usize,
        actual_rows: usize,
        actual_cols: usize,
    },
}

#[derive(Debug, Clone, PartialEq, Eq, Error)]
pub enum SheetLifecycleError {
    #[error("sheet not found")]
    SheetNotFound,
    #[error(transparent)]
    InvalidName(#[from] formula_model::SheetNameError),
    #[error("cannot delete last sheet")]
    CannotDeleteLastSheet,
    #[error("sheet index out of range")]
    IndexOutOfRange,
    #[error("engine error: {0}")]
    Internal(String),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RecalcMode {
    SingleThreaded,
    MultiThreaded,
}

#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
static RECALC_THREAD_POOL: OnceLock<Option<ThreadPool>> = OnceLock::new();

#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
fn recalc_thread_pool() -> Option<&'static ThreadPool> {
    RECALC_THREAD_POOL
        .get_or_init(build_recalc_thread_pool)
        .as_ref()
}

#[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
fn build_recalc_thread_pool() -> Option<ThreadPool> {
    // Rayon defaults to using `available_parallelism` threads for its global pool, which can be
    // excessive in test environments where the Rust test harness already spawns many threads
    // (`--test-threads`) and OS thread creation can fail with `EAGAIN`.
    //
    // Prefer a modest default and fall back to fewer threads if the pool cannot be created.
    let available = std::thread::available_parallelism()
        .map(|n| n.get())
        .unwrap_or(1);

    // Respect `RAYON_NUM_THREADS` if provided, but keep a conservative default otherwise.
    // This matches user expectations (and `scripts/cargo_agent.sh`) while still preventing
    // accidental "one pool per core" blowups on high-core CI hosts.
    let requested = std::env::var("RAYON_NUM_THREADS")
        .ok()
        .and_then(|s| s.parse::<usize>().ok())
        .filter(|n| *n > 0);

    let mut threads = match requested {
        Some(n) => n.min(available).max(1),
        None => available.min(8).max(1),
    };

    loop {
        match ThreadPoolBuilder::new().num_threads(threads).build() {
            Ok(pool) => return Some(pool),
            Err(_) if threads > 1 => {
                threads = threads / 2;
                continue;
            }
            Err(_) => return None,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct RecalcValueChange {
    pub sheet: String,
    pub addr: CellAddr,
    pub value: Value,
}

/// Scope for a defined name / named range.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NameScope<'a> {
    Workbook,
    Sheet(&'a str),
}

/// A defined name (named range) definition.
#[derive(Debug, Clone, PartialEq)]
pub enum NameDefinition {
    /// A constant scalar value (number/text/bool/error).
    Constant(Value),
    /// A reference/range definition (typically something like `Sheet1!$A$1:$B$3`).
    Reference(String),
    /// A formula definition stored as a canonical formula string (may evaluate to a scalar or reference).
    Formula(String),
}

#[derive(Debug, Clone)]
struct DefinedName {
    definition: NameDefinition,
    compiled: Option<CompiledExpr>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub(crate) struct CellKey {
    pub(crate) sheet: SheetId,
    pub(crate) addr: CellAddr,
}

#[derive(Debug, Clone)]
enum CompiledFormula {
    Ast(CompiledExpr),
    Bytecode(BytecodeFormula),
}

impl CompiledFormula {
    fn ast(&self) -> &CompiledExpr {
        match self {
            CompiledFormula::Ast(expr) => expr,
            CompiledFormula::Bytecode(bc) => &bc.ast,
        }
    }
}

#[derive(Debug, Clone)]
struct BytecodeFormula {
    ast: CompiledExpr,
    program: Arc<bytecode::Program>,
    /// Snapshot of the sheet-dimensions generation when this program was compiled.
    ///
    /// Whole-row/whole-column references (`A:A`, `1:1`) are expanded into explicit range endpoints
    /// during bytecode compilation. If any sheet's dimensions change afterwards (e.g. the sheet
    /// grows), the stored bytecode program can become stale. We track a global generation counter
    /// and fall back to AST evaluation when it no longer matches.
    sheet_dims_generation: u64,
}

#[derive(Debug, Clone)]
struct Cell {
    value: Value,
    /// Style id referencing [`Workbook::styles`].
    ///
    /// Excel preserves formatting when editing values/formulas, so engine APIs should avoid
    /// overwriting this field unless explicitly changing formatting.
    style_id: u32,
    /// Optional per-cell phonetic guide (furigana) metadata used by the `PHONETIC()` function.
    ///
    /// Lifecycle rules (Excel-like):
    /// - When a cell's *input* changes via `Engine::set_cell_value`, `Engine::set_cell_formula*`,
    ///   `Engine::set_range_values`, or copy/fill operations that overwrite cell contents, any
    ///   existing `phonetic` metadata is cleared (set to `None`) to avoid returning stale furigana
    ///   for new content.
    /// - When a cell is cleared via `Engine::clear_cell`, the cell record is removed entirely
    ///   (phonetic metadata is implicitly removed).
    /// - During recalculation, the engine may update cached `value` fields, but it must not mutate
    ///   `phonetic` metadata.
    phonetic: Option<String>,
    formula: Option<Arc<str>>,
    compiled: Option<CompiledFormula>,
    bytecode_compile_reason: Option<BytecodeCompileReason>,
    number_format: Option<String>,
    volatile: bool,
    thread_safe: bool,
    dynamic_deps: bool,
}

impl Default for Cell {
    fn default() -> Self {
        Self {
            value: Value::Blank,
            style_id: 0,
            phonetic: None,
            formula: None,
            compiled: None,
            bytecode_compile_reason: None,
            number_format: None,
            volatile: false,
            thread_safe: true,
            dynamic_deps: false,
        }
    }
}

#[derive(Debug, Clone)]
struct Sheet {
    cells: HashMap<CellAddr, Cell>,
    tables: Vec<Table>,
    names: HashMap<String, DefinedName>,
    /// Optional default style id for the entire worksheet.
    ///
    /// This is the lowest-precedence style layer in Excel's formatting chain:
    /// sheet < col < row < (range-run) < cell.
    default_style_id: Option<u32>,
    /// Sheet default column width in Excel "character" units.
    default_col_width: Option<f32>,
    /// Whether worksheet protection is enabled.
    ///
    /// Note: This is currently informational only; the engine does not enforce edit restrictions.
    sheet_protection_enabled: bool,
    /// Host-provided worksheet view metadata: the top-left visible cell in the current view.
    ///
    /// This is surfaced via `INFO("origin")`.
    origin: Option<CellAddr>,
    /// Reverse index of formula cells that depend on `INFO("origin")` for this sheet.
    ///
    /// This allows `set_sheet_origin` to efficiently mark only impacted cells dirty.
    origin_dependents: HashSet<CellAddr>,
    /// Logical row count for the worksheet grid.
    ///
    /// Defaults to Excel's row limit, but can grow beyond Excel to support very large sheets. The
    /// evaluator uses this for out-of-bounds `#REF!` semantics.
    row_count: u32,
    /// Logical column count for the worksheet grid.
    ///
    /// The engine currently enforces Excel's 16,384-column maximum.
    col_count: u32,
    /// Per-row formatting/visibility overrides.
    row_properties: BTreeMap<u32, RowProperties>,
    /// Per-column formatting/visibility overrides.
    col_properties: BTreeMap<u32, ColProperties>,
    /// Range-based formatting layer stored as per-column row interval runs.
    ///
    /// Runs are expected to be sorted by `start_row` and non-overlapping.
    format_runs_by_col: BTreeMap<u32, Vec<FormatRun>>,
}

impl Default for Sheet {
    fn default() -> Self {
        Self {
            cells: HashMap::new(),
            tables: Vec::new(),
            names: HashMap::new(),
            default_style_id: None,
            default_col_width: None,
            sheet_protection_enabled: false,
            origin: None,
            origin_dependents: HashSet::new(),
            // Default to Excel-compatible sheet bounds.
            row_count: EXCEL_MAX_ROWS,
            col_count: EXCEL_MAX_COLS,
            row_properties: BTreeMap::new(),
            col_properties: BTreeMap::new(),
            format_runs_by_col: BTreeMap::new(),
        }
    }
}

#[derive(Debug, Default, Clone)]
struct Workbook {
    sheets: Vec<Sheet>,
    /// Stable sheet keys for each sheet id.
    ///
    /// These are used for public Engine APIs (e.g. `set_cell_value`) and persistence. Indices are
    /// stable for the lifetime of a sheet; deleted sheets are represented by `None` so ids are
    /// never reused.
    sheet_keys: Vec<Option<String>>,
    /// Case-insensitive mapping (Excel semantics) from stable sheet key -> internal sheet id.
    sheet_key_to_id: HashMap<String, SheetId>,
    /// User-visible sheet tab names (display names) for each sheet id.
    ///
    /// These are used for functions that emit sheet names (e.g. `CELL("address")`) and for
    /// resolving user-visible sheet names in formulas/runtime (e.g. `INDIRECT`).
    sheet_display_names: Vec<Option<String>>,
    /// Case-insensitive mapping (Excel semantics) from sheet display name -> internal sheet id.
    sheet_display_name_to_id: HashMap<String, SheetId>,
    /// Current sheet tab order expressed as stable sheet ids.
    ///
    /// This is intentionally separate from `sheets`/`sheet_keys` so sheet ids remain stable when
    /// users reorder worksheet tabs.
    sheet_order: Vec<SheetId>,
    /// Cached mapping from stable sheet id to its current workbook tab-order index.
    ///
    /// The vector length always matches `sheet_keys.len()`. Deleted/missing sheet ids have a
    /// value of `usize::MAX`.
    sheet_tab_index_by_id: Vec<usize>,
    names: HashMap<String, DefinedName>,
    styles: StyleTable,
    workbook_directory: Option<String>,
    workbook_filename: Option<String>,
    pivots: HashMap<PivotTableId, PivotTableDefinition>,
    next_pivot_id: PivotTableId,
    /// Legacy text code page used for DBCS (`*B`) text functions.
    text_codepage: u16,
}

#[cfg(test)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum WorkbookRenameSheetError {
    SheetNotFound,
    DuplicateName,
}

impl Workbook {
    fn sheet_key(name: &str) -> String {
        // Excel compares sheet names case-insensitively across Unicode and applies compatibility
        // normalization (NFKC). We approximate this by normalizing with Unicode NFKC and then
        // applying Unicode uppercasing (locale-independent).
        //
        // This matches `formula_model::sheet_name_eq_case_insensitive`. Note: `casefold` alone is
        // not enough here; e.g. U+212A KELVIN SIGN (K) should match ASCII 'K' after NFKC
        // normalization.
        if name.is_ascii() {
            return name.to_ascii_uppercase();
        }
        name.nfkc().flat_map(|c| c.to_uppercase()).collect()
    }

    fn ensure_sheet(&mut self, sheet_key: &str) -> SheetId {
        // When adding sheets, treat the provided string as a stable key, but re-use existing sheets
        // when it matches either an existing stable key or display name. This preserves backwards
        // compatibility with call sites that address sheets by user-visible tab name.
        if let Some(id) = self.resolve_sheet_name(sheet_key) {
            return id;
        }
        let id = self.sheets.len();
        self.sheets.push(Sheet::default());
        self.sheet_keys.push(Some(sheet_key.to_string()));
        // Default display name to the stable key until overridden by the host.
        self.sheet_display_names.push(Some(sheet_key.to_string()));
        let key = Self::sheet_key(sheet_key);
        self.sheet_key_to_id.insert(key.clone(), id);
        // Ensure new sheets are resolvable by their initial display name (which defaults to the
        // key).
        self.sheet_display_name_to_id.insert(key, id);
        self.sheet_order.push(id);
        self.sheet_tab_index_by_id
            .push(self.sheet_order.len().saturating_sub(1));
        id
    }
    fn rebuild_sheet_tab_index_by_id(&mut self) {
        // Keep the cache aligned with the stable sheet-id space.
        self.sheet_tab_index_by_id
            .resize(self.sheet_keys.len(), usize::MAX);
        self.sheet_tab_index_by_id.fill(usize::MAX);
        for (idx, sheet_id) in self.sheet_order.iter().copied().enumerate() {
            if let Some(slot) = self.sheet_tab_index_by_id.get_mut(sheet_id) {
                *slot = idx;
            }
        }
    }

    fn tab_index_by_sheet_id(&self) -> &[usize] {
        &self.sheet_tab_index_by_id
    }

    fn sheet_id_by_key(&self, sheet_key: &str) -> Option<SheetId> {
        let key = Self::sheet_key(sheet_key);
        self.sheet_key_to_id.get(&key).copied()
    }

    fn sheet_id(&self, name: &str) -> Option<SheetId> {
        self.resolve_sheet_name(name)
    }

    fn resolve_sheet_name(&self, name: &str) -> Option<SheetId> {
        // Excel formulas reference the user-visible tab name. Resolve by display name
        // case-insensitively, but fall back to the stable key so existing call sites / persisted
        // workbooks that reference `sheet_key` keep working.
        let key = Self::sheet_key(name);
        self.sheet_display_name_to_id
            .get(&key)
            .copied()
            .or_else(|| self.sheet_key_to_id.get(&key).copied())
    }

    /// Update the user-visible display name for `sheet_id`.
    ///
    /// Returns `true` when the display name changed.
    fn set_sheet_display_name(&mut self, sheet_id: SheetId, display_name: &str) -> bool {
        if !self.sheet_exists(sheet_id) {
            return false;
        }
        let Some(current) = self
            .sheet_display_names
            .get(sheet_id)
            .and_then(|name| name.as_ref())
        else {
            return false;
        };
        if current == display_name {
            return false;
        }
        // Sheet display names are Excel worksheet tab names, so they must follow Excel's name
        // validation rules (max length, forbidden characters, etc.). If invalid, ignore the change
        // to preserve a consistent workbook name mapping (and avoid creating unparseable formulas).
        if formula_model::validate_sheet_name(display_name).is_err() {
            return false;
        }

        // Avoid creating ambiguous mappings when callers attempt to set a duplicate display name.
        let new_key = Self::sheet_key(display_name);
        if let Some(existing) = self.sheet_display_name_to_id.get(&new_key).copied() {
            if existing != sheet_id {
                return false;
            }
        }
        // Ensure display names do not shadow another sheet's stable key; stable keys must remain
        // resolvable for hosts that address worksheets by `sheet_key`.
        if let Some(existing) = self.sheet_key_to_id.get(&new_key).copied() {
            if existing != sheet_id {
                return false;
            }
        }

        let old_key = Self::sheet_key(current);
        if self.sheet_display_name_to_id.get(&old_key) == Some(&sheet_id) {
            self.sheet_display_name_to_id.remove(&old_key);
        }
        self.sheet_display_names[sheet_id] = Some(display_name.to_string());
        self.sheet_display_name_to_id.insert(new_key, sheet_id);
        true
    }

    #[cfg(test)]
    fn rename_sheet(
        &mut self,
        sheet_id: SheetId,
        new_name: &str,
    ) -> Result<(), WorkbookRenameSheetError> {
        if !self.sheet_exists(sheet_id) {
            return Err(WorkbookRenameSheetError::SheetNotFound);
        }

        let new_key = Self::sheet_key(new_name);
        // Enforce uniqueness across both stable keys and display names (Excel semantics).
        if let Some(existing) = self.resolve_sheet_name(new_name) {
            if existing != sheet_id {
                return Err(WorkbookRenameSheetError::DuplicateName);
            }
        }

        let old_key_name = self
            .sheet_keys
            .get(sheet_id)
            .and_then(|name| name.as_ref())
            .ok_or(WorkbookRenameSheetError::SheetNotFound)?
            .clone();
        let old_key = Self::sheet_key(&old_key_name);
        let old_display_name = self
            .sheet_display_names
            .get(sheet_id)
            .and_then(|name| name.as_ref())
            .ok_or(WorkbookRenameSheetError::SheetNotFound)?
            .clone();
        let old_display_key = Self::sheet_key(&old_display_name);

        self.sheet_keys[sheet_id] = Some(new_name.to_string());
        self.sheet_display_names[sheet_id] = Some(new_name.to_string());

        // Remove the old lookup key (if it still points at this sheet) and install the new one.
        // This keeps lookups consistent even when names are renormalized (e.g. `Å` -> `Å`).
        if self.sheet_key_to_id.get(&old_key) == Some(&sheet_id) {
            self.sheet_key_to_id.remove(&old_key);
        }
        if self.sheet_display_name_to_id.get(&old_display_key) == Some(&sheet_id) {
            self.sheet_display_name_to_id.remove(&old_display_key);
        }
        self.sheet_key_to_id.insert(new_key.clone(), sheet_id);
        self.sheet_display_name_to_id.insert(new_key, sheet_id);

        Ok(())
    }

    fn sheet_exists(&self, sheet: SheetId) -> bool {
        matches!(self.sheet_keys.get(sheet), Some(Some(_)))
    }

    fn sheet_name(&self, sheet: SheetId) -> Option<&str> {
        self.sheet_display_names.get(sheet)?.as_deref()
    }

    fn sheet_key_name(&self, sheet: SheetId) -> Option<&str> {
        self.sheet_keys.get(sheet)?.as_deref()
    }

    fn sheet_ids_in_order(&self) -> &[SheetId] {
        &self.sheet_order
    }

    #[cfg(test)]
    fn set_sheet_order(&mut self, new_order: Vec<SheetId>) {
        // Keep invariants explicit: sheet order is a permutation of the currently-live sheets.
        let existing: HashSet<SheetId> = self.sheet_order.iter().copied().collect();
        let mut seen: HashSet<SheetId> = HashSet::with_capacity(new_order.len());
        for &id in &new_order {
            assert!(
                self.sheet_exists(id),
                "sheet order contains missing sheet id {}",
                id
            );
            assert!(
                seen.insert(id),
                "sheet order contains duplicate sheet id {id}"
            );
        }
        assert_eq!(
            seen, existing,
            "sheet order must contain exactly the workbook's live sheets"
        );
        self.sheet_order = new_order;
        self.rebuild_sheet_tab_index_by_id();
    }

    fn sheet_order_index(&self, sheet: SheetId) -> Option<usize> {
        let idx = *self.sheet_tab_index_by_id.get(sheet)?;
        if idx == usize::MAX {
            return None;
        }
        // `sheet_tab_index_by_id` is a cache derived from `sheet_order`. In normal operation it is
        // updated whenever `sheet_order` changes, but internal tests (or future refactors) may
        // mutate `sheet_order` without updating the cache. Validate the cached index before
        // trusting it so we don't incorrectly treat a missing sheet as present.
        if self.sheet_order.get(idx).copied() == Some(sheet) {
            return Some(idx);
        }

        // Fallback: linear scan. This should be rare (cache is normally kept in sync).
        self.sheet_order.iter().position(|&id| id == sheet)
    }

    /// Reorder a worksheet id within the workbook's tab order.
    ///
    /// This only affects `sheet_order` (sheet ids are stable and do not change).
    fn reorder_sheet(&mut self, sheet: SheetId, new_index: usize) -> bool {
        if !self.sheet_exists(sheet) {
            return false;
        }
        if new_index >= self.sheet_order.len() {
            return false;
        }
        // Prefer searching `sheet_order` directly rather than relying on the cached tab index.
        //
        // `sheet_tab_index_by_id` is derived state; if invariants are broken (e.g. tests mutate
        // `sheet_order` directly), the cache may be stale. In that scenario we must reject the
        // reorder request without mutating the remaining order.
        let Some(current) = self.sheet_order.iter().position(|&id| id == sheet) else {
            // Keep the cache aligned with the current order so subsequent lookups behave
            // consistently even if the workbook is already in an inconsistent state.
            self.rebuild_sheet_tab_index_by_id();
            return false;
        };
        if current == new_index {
            // Even in a no-op reorder, keep the cache aligned with `sheet_order` in case it became
            // stale.
            self.rebuild_sheet_tab_index_by_id();
            return true;
        }
        let id = self.sheet_order.remove(current);
        // `new_index` is expressed in terms of the final tab order; inserting at that index after
        // removal produces the expected result (Vec::insert supports `index == len`).
        self.sheet_order.insert(new_index, id);
        self.rebuild_sheet_tab_index_by_id();
        true
    }
    /// Returns the sheet ids referenced by an Excel-style 3D sheet span (`Sheet1:Sheet3`).
    ///
    /// This respects the current workbook tab order. Reversed spans are allowed (e.g.
    /// `Sheet3:Sheet1`) and refer to the same set of sheets.
    fn sheet_span_ids(&self, start: SheetId, end: SheetId) -> Option<Vec<SheetId>> {
        let start_idx = self.sheet_order_index(start)?;
        let end_idx = self.sheet_order_index(end)?;
        let (start_idx, end_idx) = if start_idx <= end_idx {
            (start_idx, end_idx)
        } else {
            (end_idx, start_idx)
        };
        Some(self.sheet_order[start_idx..end_idx.saturating_add(1)].to_vec())
    }
    fn get_cell(&self, key: CellKey) -> Option<&Cell> {
        if !self.sheet_exists(key.sheet) {
            return None;
        }
        self.sheets.get(key.sheet)?.cells.get(&key.addr)
    }

    fn get_or_create_cell_mut(&mut self, key: CellKey) -> &mut Cell {
        assert!(
            self.sheet_exists(key.sheet),
            "attempted to access missing sheet id {}",
            key.sheet
        );
        self.sheets[key.sheet].cells.entry(key.addr).or_default()
    }

    fn set_tables(&mut self, sheet: SheetId, tables: Vec<Table>) {
        if !self.sheet_exists(sheet) {
            return;
        }
        if let Some(s) = self.sheets.get_mut(sheet) {
            s.tables = tables;
        }
    }

    fn grow_sheet_dimensions(&mut self, sheet: SheetId, addr: CellAddr) -> bool {
        if !self.sheet_exists(sheet) {
            return false;
        }
        let Some(s) = self.sheets.get_mut(sheet) else {
            return false;
        };
        // Dimensions are stored as counts (1-based), while addresses are 0-based.
        let new_row_count = s.row_count.max(addr.row.saturating_add(1));
        let new_col_count = s.col_count.max(addr.col.saturating_add(1));
        let changed = new_row_count != s.row_count || new_col_count != s.col_count;
        s.row_count = new_row_count;
        s.col_count = new_col_count;
        changed
    }
}

/// A node returned from auditing/introspection APIs (precedents/dependents).
///
/// This can represent either a single cell or a rectangular range reference without
/// expanding it into per-cell nodes (which is prohibitive for `A:A`, `1:1`, etc).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum PrecedentNode {
    Cell {
        sheet: SheetId,
        addr: CellAddr,
    },
    Range {
        sheet: SheetId,
        start: CellAddr,
        end: CellAddr,
    },
    /// Cell reference into an external workbook, e.g. `[Book.xlsx]Sheet1!A1`.
    ///
    /// `sheet` is the canonical external sheet key (`"[workbook]sheet"`, e.g. `"[Book.xlsx]Sheet1"`).
    ///
    /// For external 3D spans, auditing may also surface the unexpanded span key
    /// (`"[workbook]Sheet1:Sheet3"`) when expansion is not possible (e.g. no provider or missing
    /// `sheet_order`).
    ExternalCell {
        sheet: String,
        addr: CellAddr,
    },
    /// Range reference into an external workbook, e.g. `[Book.xlsx]Sheet1!A1:B3`.
    ///
    /// `sheet` is the canonical external sheet key (`"[workbook]sheet"`, e.g. `"[Book.xlsx]Sheet1"`).
    ///
    /// For external 3D spans, auditing may also surface the unexpanded span key
    /// (`"[workbook]Sheet1:Sheet3"`) when expansion is not possible (e.g. no provider or missing
    /// `sheet_order`).
    ExternalRange {
        sheet: String,
        start: CellAddr,
        end: CellAddr,
    },
    /// Dynamic array spill footprint for a spilled formula (origin -> footprint).
    ///
    /// Dynamic array evaluation is not implemented yet, but the auditing API reserves a
    /// node type so the UI can represent spill relationships without expanding every
    /// cell in the footprint.
    SpillRange {
        sheet: SheetId,
        origin: CellAddr,
        start: CellAddr,
        end: CellAddr,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DirtyReason {
    Cell(CellKey),
    ViaRange { from: CellKey, range: PrecedentNode },
}

#[derive(Debug, Clone)]
struct Spill {
    end: CellAddr,
    array: Array,
}

#[derive(Debug, Clone)]
struct BlockedSpill {
    blocker: CellKey,
}

#[derive(Debug, Default, Clone)]
struct SpillState {
    by_origin: HashMap<CellKey, Spill>,
    origin_by_cell: HashMap<CellKey, CellKey>,
    /// Spill origins currently evaluating to `#SPILL!` due to a blocked spill range.
    blocked_by_origin: HashMap<CellKey, BlockedSpill>,
    /// Reverse index: blocker cell -> spill origins that should be re-evaluated when the blocker changes.
    blocked_origins_by_cell: HashMap<CellKey, HashSet<CellKey>>,
}

pub struct Engine {
    workbook: Workbook,
    bytecode_cache: bytecode::BytecodeCache,
    bytecode_enabled: bool,
    /// Controls whether external workbook references (e.g. `[Book.xlsx]Sheet1!A1`) are treated as
    /// volatile roots.
    ///
    /// When enabled (the default), formulas containing any external reference are automatically
    /// re-evaluated on every recalculation pass. When disabled, external references are treated as
    /// non-volatile and will only refresh when explicitly invalidated via
    /// [`Engine::mark_external_sheet_dirty`] / [`Engine::mark_external_workbook_dirty`] or when the
    /// formula is otherwise marked dirty.
    external_refs_volatile: bool,
    /// Monotonic counter incremented whenever any sheet's configured dimensions change.
    ///
    /// See [`BytecodeFormula::sheet_dims_generation`].
    sheet_dims_generation: u64,
    external_value_provider: Option<Arc<dyn ExternalValueProvider>>,
    external_data_provider: Option<Arc<dyn ExternalDataProvider>>,
    pivot_registry: crate::pivot_registry::PivotRegistry,
    name_dependents: HashMap<String, HashSet<CellKey>>,
    cell_name_refs: HashMap<CellKey, HashSet<String>>,
    /// Reverse index: external sheet key -> local formula cells that reference it.
    external_sheet_dependents: HashMap<String, HashSet<CellKey>>,
    /// Forward index: local formula cell -> external sheet keys referenced by its formula.
    cell_external_sheet_refs: HashMap<CellKey, HashSet<String>>,
    /// Reverse index: external workbook identifier (e.g. `Book.xlsx`) -> local formula cells that
    /// reference any sheet within that workbook.
    external_workbook_dependents: HashMap<String, HashSet<CellKey>>,
    /// Forward index: local formula cell -> external workbook identifiers referenced by its formula.
    cell_external_workbook_refs: HashMap<CellKey, HashSet<String>>,
    /// Dynamic external precedents captured at runtime for dynamic-deps formulas.
    ///
    /// These precedents cannot be represented in the internal dependency graph yet, but are still
    /// surfaced via auditing APIs like [`Engine::precedents`].
    cell_dynamic_external_precedents: HashMap<CellKey, HashSet<crate::functions::Reference>>,
    /// Reverse index: external sheet key -> local formula cells that dereferenced it at runtime.
    dynamic_external_sheet_dependents: HashMap<String, HashSet<CellKey>>,
    /// Forward index: local formula cell -> external sheet keys dereferenced at runtime.
    cell_dynamic_external_sheet_refs: HashMap<CellKey, HashSet<String>>,
    /// Reverse index: external workbook identifier -> local formula cells that dereferenced it at runtime.
    dynamic_external_workbook_dependents: HashMap<String, HashSet<CellKey>>,
    /// Forward index: local formula cell -> external workbook identifiers dereferenced at runtime.
    cell_dynamic_external_workbook_refs: HashMap<CellKey, HashSet<String>>,
    /// Optimized dependency graph used for incremental recalculation ordering.
    calc_graph: CalcGraph,
    dirty: HashSet<CellKey>,
    dirty_reasons: HashMap<CellKey, DirtyReason>,
    calc_settings: CalcSettings,
    date_system: ExcelDateSystem,
    value_locale: ValueLocaleConfig,
    locale_config: crate::LocaleConfig,
    text_codepage: u16,
    circular_references: HashSet<CellKey>,
    spills: SpillState,
    next_recalc_id: u64,
    info: EngineInfo,
}

#[derive(Default)]
struct RecalcValueChangeCollector {
    before: HashMap<CellKey, Value>,
    after: HashMap<CellKey, Value>,
}

impl RecalcValueChangeCollector {
    fn record(&mut self, key: CellKey, before: Value, after: Value) {
        if before == after {
            return;
        }
        self.before.entry(key).or_insert(before);
        self.after.insert(key, after);
    }

    fn into_sorted_changes(self, workbook: &Workbook) -> Vec<RecalcValueChange> {
        let mut out = Vec::new();
        let tab_index_by_sheet = workbook.tab_index_by_sheet_id();
        let mut after: Vec<(CellKey, Value)> = self.after.into_iter().collect();
        after.sort_by(|(a_key, _), (b_key, _)| {
            sheet_tab_key(a_key.sheet, tab_index_by_sheet)
                .cmp(&sheet_tab_key(b_key.sheet, tab_index_by_sheet))
                .then_with(|| a_key.addr.row.cmp(&b_key.addr.row))
                .then_with(|| a_key.addr.col.cmp(&b_key.addr.col))
        });

        for (key, after) in after {
            let before = self
                .before
                .get(&key)
                .expect("recalc change must record before value");
            if *before == after {
                continue;
            }
            // Recalc change events are addressed using the stable sheet key (not the user-visible
            // display name) so hosts can map the change back onto their own sheet identifiers.
            let sheet = workbook
                .sheet_key_name(key.sheet)
                .unwrap_or_default()
                .to_string();
            out.push(RecalcValueChange {
                sheet,
                addr: key.addr,
                value: after,
            });
        }
        out
    }
}

impl Default for Engine {
    fn default() -> Self {
        Self::new()
    }
}

impl Engine {
    /// Create a new in-memory engine instance backed by an empty workbook.
    ///
    /// # Calculation mode
    ///
    /// The engine defaults to **manual** calculation mode to preserve historical behavior (and so
    /// formula evaluation happens when callers explicitly request it via `recalculate_*`).
    ///
    /// To opt into Excel-like automatic calculation semantics, set
    /// [`CalcSettings::calculation_mode`] via [`Engine::set_calc_settings`].
    pub fn new() -> Self {
        let mut workbook = Workbook::default();
        // Excel default (en-US) code page / ANSI.
        workbook.text_codepage = 1252;

        Self {
            workbook,
            bytecode_cache: bytecode::BytecodeCache::new(),
            bytecode_enabled: true,
            external_refs_volatile: true,
            sheet_dims_generation: 0,
            external_value_provider: None,
            external_data_provider: None,
            pivot_registry: crate::pivot_registry::PivotRegistry::default(),
            name_dependents: HashMap::new(),
            cell_name_refs: HashMap::new(),
            external_sheet_dependents: HashMap::new(),
            cell_external_sheet_refs: HashMap::new(),
            external_workbook_dependents: HashMap::new(),
            cell_external_workbook_refs: HashMap::new(),
            cell_dynamic_external_precedents: HashMap::new(),
            dynamic_external_sheet_dependents: HashMap::new(),
            cell_dynamic_external_sheet_refs: HashMap::new(),
            dynamic_external_workbook_dependents: HashMap::new(),
            cell_dynamic_external_workbook_refs: HashMap::new(),
            calc_graph: CalcGraph::new(),
            dirty: HashSet::new(),
            dirty_reasons: HashMap::new(),
            // Default to manual calculation to preserve historical engine behavior; callers can
            // opt into Excel-like automatic mode by setting `CalcSettings.calculation_mode`.
            calc_settings: CalcSettings {
                calculation_mode: CalculationMode::Manual,
                ..CalcSettings::default()
            },
            date_system: ExcelDateSystem::EXCEL_1900,
            value_locale: ValueLocaleConfig::default(),
            locale_config: crate::LocaleConfig::en_us(),
            text_codepage: 1252,
            circular_references: HashSet::new(),
            spills: SpillState::default(),
            next_recalc_id: 0,
            info: EngineInfo::default(),
        }
    }

    /// Replace the range-run formatting runs for a column.
    ///
    /// `runs` are expected to be sorted by `start_row` and non-overlapping, but this method will
    /// perform a best-effort normalization (sorting + dropping empty/default runs) to avoid
    /// corrupting engine state when callers pass malformed input.
    pub fn set_col_format_runs(
        &mut self,
        sheet: &str,
        col: u32,
        mut runs: Vec<FormatRun>,
    ) -> Result<(), EngineError> {
        if col >= EXCEL_MAX_COLS {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::ColumnOutOfRange,
            ));
        }

        // Normalize runs to preserve sparse semantics and prevent obvious corruption.
        runs.retain(|r| r.start_row < r.end_row_exclusive && r.style_id != 0);
        runs.sort_by_key(|r| r.start_row);

        let sheet_id = self.workbook.ensure_sheet(sheet);
        let mut sheet_dims_changed = false;
        if let Some(max_row) = runs
            .iter()
            .map(|r| r.end_row_exclusive.saturating_sub(1))
            .max()
        {
            if max_row >= i32::MAX as u32 {
                return Err(EngineError::Address(
                    crate::eval::AddressParseError::RowOutOfRange,
                ));
            }
            sheet_dims_changed = self
                .workbook
                .grow_sheet_dimensions(sheet_id, CellAddr { row: max_row, col });
        } else if self
            .workbook
            .grow_sheet_dimensions(sheet_id, CellAddr { row: 0, col })
        {
            sheet_dims_changed = true;
        }

        let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) else {
            return Ok(());
        };

        // Detect no-op updates to avoid unnecessary recalculation / dirtying.
        let before = sheet_state.format_runs_by_col.get(&col);
        let runs_changed = match (before, runs.is_empty()) {
            (None, true) => false,
            (Some(_), true) => true,
            (None, false) => true,
            (Some(existing), false) => existing != &runs,
        };

        if !runs_changed && !sheet_dims_changed {
            return Ok(());
        }

        if runs.is_empty() {
            sheet_state.format_runs_by_col.remove(&col);
        } else {
            sheet_state.format_runs_by_col.insert(col, runs);
        }

        if sheet_dims_changed {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
        }
        if sheet_dims_changed || (!self.calc_settings.full_precision && runs_changed) {
            // Formatting metadata can introduce new in-bounds coordinates. Sheet dimensions affect
            // out-of-bounds `#REF!` semantics, so conservatively refresh compiled results.
            //
            // In "precision as displayed" mode, formatting can affect stored numeric values at
            // formula boundaries, so retain conservative full-dirty behavior for any formatting
            // changes.
            self.mark_all_compiled_cells_dirty();
        }
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    pub fn calc_settings(&self) -> &CalcSettings {
        &self.calc_settings
    }

    /// Returns the host-provided metadata surfaced by the `INFO()` worksheet function.
    pub fn engine_info(&self) -> &EngineInfo {
        &self.info
    }

    /// Replace the host-provided metadata surfaced by the `INFO()` worksheet function.
    pub fn set_engine_info(&mut self, info: EngineInfo) {
        if self.info == info {
            return;
        }
        self.info = info;
        // `INFO()` is volatile, so any dependent formulas are included in the dependency graph's
        // volatile closure. We can therefore refresh results with a recalculation tick without
        // dirtying the entire workbook.
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Set the workbook-level default value for `INFO("origin")`.
    pub fn set_info_origin(&mut self, origin: Option<impl Into<String>>) {
        let origin = origin.map(Into::into);
        if self.info.origin == origin {
            return;
        }
        self.info.origin = origin;
        // `INFO("origin")` is volatile; trigger a recalculation tick in automatic modes.
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Set the per-sheet value for `INFO("origin")`.
    ///
    /// This overrides [`Engine::set_info_origin`] for the given sheet.
    pub fn set_info_origin_for_sheet(&mut self, sheet: &str, origin: Option<impl Into<String>>) {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let origin = origin.map(Into::into);
        let changed = match &origin {
            Some(v) => self.info.origin_by_sheet.get(&sheet_id) != Some(v),
            None => self.info.origin_by_sheet.contains_key(&sheet_id),
        };
        if !changed {
            return;
        }

        match origin {
            Some(v) => {
                self.info.origin_by_sheet.insert(sheet_id, v);
            }
            None => {
                self.info.origin_by_sheet.remove(&sheet_id);
            }
        }

        // `INFO("origin")` is volatile; trigger a recalculation tick in automatic modes.
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Ensure a sheet exists in the workbook.
    ///
    /// This is useful for workbook load flows where formulas may refer to other sheets
    /// that have not been populated yet; callers should create all sheets up-front
    /// before setting formulas to ensure cross-sheet references resolve correctly.
    pub fn ensure_sheet(&mut self, sheet: &str) {
        self.workbook.ensure_sheet(sheet);
    }

    /// Resolve a worksheet name to its stable [`SheetId`].
    ///
    /// Matching is case-insensitive and Unicode/NFKC-aware (Excel-like).
    ///
    /// Sheet ids remain stable across renames and tab reorders, but become invalid once a sheet is
    /// deleted.
    pub fn sheet_id(&self, name: &str) -> Option<SheetId> {
        self.workbook.sheet_id(name)
    }

    /// Resolve a stable [`SheetId`] back to its current worksheet name.
    pub fn sheet_name(&self, id: SheetId) -> Option<&str> {
        self.workbook.sheet_name(id)
    }

    /// Returns stable sheet ids in the current workbook tab order.
    pub fn sheet_ids_in_order(&self) -> Vec<SheetId> {
        self.workbook.sheet_ids_in_order().to_vec()
    }

    /// Returns worksheet stable keys (sheet identifiers) in the current workbook tab order.
    ///
    /// This differs from [`Engine::sheet_names_in_order`], which returns user-visible worksheet
    /// display names (tab names). Sheet keys remain stable across renames and are used by hosts
    /// (like the desktop DocumentController) to address worksheets consistently.
    pub fn sheet_keys_in_order(&self) -> Vec<String> {
        self.workbook
            .sheet_ids_in_order()
            .iter()
            .filter_map(|&id| self.workbook.sheet_key_name(id).map(|name| name.to_string()))
            .collect()
    }

    /// Returns worksheet display names in the current workbook tab order.
    pub fn sheet_names_in_order(&self) -> Vec<String> {
        self.workbook
            .sheet_ids_in_order()
            .iter()
            .filter_map(|&id| self.workbook.sheet_name(id).map(|name| name.to_string()))
            .collect()
    }

    /// Returns `(sheet_id, display_name)` pairs in the current workbook tab order.
    pub fn sheets_in_order(&self) -> Vec<(SheetId, String)> {
        self.workbook
            .sheet_ids_in_order()
            .iter()
            .filter_map(|&id| {
                self.workbook
                    .sheet_name(id)
                    .map(|name| (id, name.to_string()))
            })
            .collect()
    }

    /// Rename a worksheet by its stable [`SheetId`].
    ///
    /// This matches [`Engine::rename_sheet`] semantics:
    /// - Stored formula text is rewritten across the workbook (cells, tables, and defined names).
    /// - External workbook references (`[Book.xlsx]Sheet1!A1`) are **not** rewritten.
    ///
    /// If `id` is invalid or already deleted, this is a no-op and returns `Ok(())`.
    ///
    /// Renaming does not change the stable sheet id.
    pub fn rename_sheet_by_id(
        &mut self,
        id: SheetId,
        new_name: &str,
    ) -> Result<(), SheetLifecycleError> {
        fn expr_contains_formulatext(expr: &CompiledExpr) -> bool {
            let mut stack = vec![expr];
            while let Some(expr) = stack.pop() {
                match expr {
                    Expr::FunctionCall { name, args, .. } => {
                        if name.eq_ignore_ascii_case("FORMULATEXT") {
                            return true;
                        }
                        stack.extend(args.iter());
                    }
                    Expr::ArrayLiteral { values, .. } => stack.extend(values.iter()),
                    Expr::FieldAccess { base, .. } => stack.push(base),
                    Expr::Unary { expr, .. } | Expr::Postfix { expr, .. } => stack.push(expr),
                    Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
                        stack.push(left);
                        stack.push(right);
                    }
                    Expr::Call { callee, args } => {
                        stack.push(callee);
                        stack.extend(args.iter());
                    }
                    Expr::ImplicitIntersection(inner) | Expr::SpillRange(inner) => {
                        stack.push(inner);
                    }
                    Expr::Number(_)
                    | Expr::Text(_)
                    | Expr::Bool(_)
                    | Expr::Blank
                    | Expr::Error(_)
                    | Expr::NameRef(_)
                    | Expr::CellRef(_)
                    | Expr::RangeRef(_)
                    | Expr::StructuredRef(_) => {}
                }
            }
            false
        }

        if !self.workbook.sheet_exists(id) {
            // Invalid/deleted ids are treated as a no-op so hosts can safely apply stale operations.
            return Ok(());
        }

        formula_model::validate_sheet_name(new_name)?;
        // Enforce uniqueness across both stable keys and display names (Excel semantics).
        if let Some(existing) = self.workbook.resolve_sheet_name(new_name) {
            if existing != id {
                return Err(SheetLifecycleError::InvalidName(
                    formula_model::SheetNameError::DuplicateName,
                ));
            }
        }

        let old_name = self
            .workbook
            .sheet_name(id)
            .ok_or(SheetLifecycleError::SheetNotFound)?
            .to_string();
        let old_sheet_key = self.workbook.sheet_key_name(id).map(|s| s.to_string());
        if old_name == new_name {
            return Ok(());
        }
        let new_name_owned = new_name.to_string();

        // Rewrite stored formulas so future recompiles don't treat references as external.
        let mut any_formula_rewritten = false;
        let mut formulatext_cells: Vec<CellKey> = Vec::new();
        let sheet_ids: Vec<SheetId> = self.workbook.sheet_ids_in_order().to_vec();
        for sheet_id in sheet_ids {
            let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) else {
                continue;
            };

            for (addr, cell) in sheet.cells.iter_mut() {
                if let Some(compiled) = cell.compiled.as_ref() {
                    if expr_contains_formulatext(compiled.ast()) {
                        formulatext_cells.push(CellKey {
                            sheet: sheet_id,
                            addr: *addr,
                        });
                    }
                }

                let Some(formula) = cell.formula.as_mut() else {
                    continue;
                };
                let rewritten = formula_model::rewrite_sheet_names_in_formula(
                    formula.as_ref(),
                    &old_name,
                    new_name,
                );
                if rewritten != formula.as_ref() {
                    any_formula_rewritten = true;
                    *formula = rewritten.into();
                }
            }

            for table in &mut sheet.tables {
                for column in &mut table.columns {
                    if let Some(formula) = column.formula.as_mut() {
                        let rewritten = formula_model::rewrite_sheet_names_in_formula(
                            formula, &old_name, new_name,
                        );
                        *formula = rewritten;
                    }
                    if let Some(formula) = column.totals_formula.as_mut() {
                        let rewritten = formula_model::rewrite_sheet_names_in_formula(
                            formula, &old_name, new_name,
                        );
                        *formula = rewritten;
                    }
                }
            }

            for def in sheet.names.values_mut() {
                match &mut def.definition {
                    NameDefinition::Constant(_) => {}
                    NameDefinition::Reference(formula) | NameDefinition::Formula(formula) => {
                        let rewritten = formula_model::rewrite_sheet_names_in_formula(
                            formula, &old_name, new_name,
                        );
                        *formula = rewritten;
                    }
                }
            }
        }

        for def in self.workbook.names.values_mut() {
            match &mut def.definition {
                NameDefinition::Constant(_) => {}
                NameDefinition::Reference(formula) | NameDefinition::Formula(formula) => {
                    let rewritten =
                        formula_model::rewrite_sheet_names_in_formula(formula, &old_name, new_name);
                    *formula = rewritten;
                }
            }
        }

        // If this sheet's stable key matched its display name, keep them in sync when renaming so
        // the old sheet name is fully removed from name resolution (`Engine::sheet_id`) and does
        // not linger as an alternate stable-key alias.
        if old_sheet_key
            .as_ref()
            .is_some_and(|k| formula_model::sheet_name_eq_case_insensitive(k, &old_name))
        {
            let new_key = Workbook::sheet_key(new_name);
            if let Some(old_key) = old_sheet_key.as_ref().map(|k| Workbook::sheet_key(k)) {
                if self.workbook.sheet_key_to_id.get(&old_key) == Some(&id) {
                    self.workbook.sheet_key_to_id.remove(&old_key);
                }
            }
            self.workbook.sheet_keys[id] = Some(new_name.to_string());
            self.workbook.sheet_key_to_id.insert(new_key, id);
        }

        if !self.workbook.set_sheet_display_name(id, new_name) {
            return Err(SheetLifecycleError::Internal(format!(
                "failed to rename sheet id {id}"
            )));
        }

        // `rename_sheet_*` implements Excel-like rename semantics. When a host does not provide a
        // separate stable sheet key (i.e. the key matches the old display name), the key should be
        // updated as well so the old name no longer resolves via the fallback
        // `display_name -> key` lookup. This prevents operations like pivot refresh (which uses
        // `ensure_sheet`) from accidentally resurrecting the pre-rename name.
        let key_name = self
            .workbook
            .sheet_key_name(id)
            .unwrap_or(&old_name)
            .to_string();
        if formula_model::sheet_name_eq_case_insensitive(&key_name, &old_name) {
            let old_key = Workbook::sheet_key(&key_name);
            let new_key = Workbook::sheet_key(new_name);
            if let Some(slot) = self.workbook.sheet_keys.get_mut(id) {
                *slot = Some(new_name.to_string());
            }
            if self.workbook.sheet_key_to_id.get(&old_key) == Some(&id) {
                self.workbook.sheet_key_to_id.remove(&old_key);
            }
            self.workbook.sheet_key_to_id.insert(new_key, id);
        }

        // Keep stored pivot definitions aligned with sheet renames. Pivot definitions store sheet
        // names (not stable ids) and pivot refresh uses `set_cell_value`, which would otherwise
        // recreate the *old* sheet name via `ensure_sheet`.
        for pivot in self.workbook.pivots.values_mut() {
            if formula_model::sheet_name_eq_case_insensitive(&pivot.destination.sheet, &old_name) {
                pivot.destination.sheet = new_name_owned.clone();
            }
            if let crate::pivot::PivotSource::Range { sheet, .. } = &mut pivot.source {
                if formula_model::sheet_name_eq_case_insensitive(sheet, &old_name) {
                    *sheet = new_name_owned.clone();
                }
            }
        }

        // Renaming a worksheet rewrites stored formula text, but it should not change computed
        // numeric results. Avoid full-workbook dirtying/recalc; instead:
        // - Trigger a recalculation tick in automatic modes so volatile workbook information
        //   functions (e.g. `CELL`) can refresh their outputs.
        // - Mark `FORMULATEXT` formulas dirty when any stored formula text was rewritten so those
        //   cells can re-read updated formula strings.
        if any_formula_rewritten {
            for key in formulatext_cells {
                self.dirty.insert(key);
                self.dirty_reasons.remove(&key);
                self.calc_graph.mark_dirty(cell_id_from_key(key));
            }
            self.sync_dirty_from_calc_graph();
        }
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }

        Ok(())
    }

    /// Reorder a worksheet within the workbook's tab order.
    ///
    /// 3D sheet spans like `Sheet1:Sheet3!A1` are defined in terms of workbook tab order, so
    /// reordering sheets can change both formula semantics and dependency sets. The engine
    /// conservatively rebuilds the dependency graph (recompiling bytecode formulas) so any
    /// pre-expanded sheet spans are refreshed.
    pub fn reorder_sheet(&mut self, sheet: &str, new_index: usize) -> bool {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return false;
        };
        self.reorder_sheet_by_id(sheet_id, new_index).is_ok()
    }

    /// Reorder a worksheet within the workbook's tab order using its stable [`SheetId`].
    ///
    /// If `id` is invalid or already deleted, this is a no-op and returns `Ok(())`.
    pub fn reorder_sheet_by_id(
        &mut self,
        id: SheetId,
        new_index: usize,
    ) -> Result<(), SheetLifecycleError> {
        if !self.workbook.sheet_exists(id) {
            return Ok(());
        }
        if new_index >= self.workbook.sheet_order.len() {
            return Err(SheetLifecycleError::IndexOutOfRange);
        }
        let Some(original_index) = self.workbook.sheet_order_index(id) else {
            return Err(SheetLifecycleError::Internal(format!(
                "sheet id {id} missing from workbook sheet_order"
            )));
        };
        if original_index == new_index {
            return Ok(());
        }
        if !self.workbook.reorder_sheet(id, new_index) {
            return Err(SheetLifecycleError::Internal(format!(
                "failed to reorder sheet id {id}"
            )));
        }

        if let Err(e) = self
            .recompile_all_defined_names()
            .and_then(|_| self.rebuild_graph())
        {
            // Reordering should not introduce new parse errors (formulas are unchanged), but if
            // rebuilding fails for any reason, restore the previous order and best-effort rebuild
            // to keep the engine in a consistent state.
            let _ = self.workbook.reorder_sheet(id, original_index);
            let _ = self.recompile_all_defined_names();
            let _ = self.rebuild_graph();
            return Err(SheetLifecycleError::Internal(e.to_string()));
        }

        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }

        Ok(())
    }
    /// Insert (or reuse) a style in the workbook's style table, returning its stable id.
    pub fn intern_style(&mut self, style: Style) -> u32 {
        // Inserting a new style does not affect existing cell/row/col style ids, so it should not
        // invalidate formula results on its own. Callers that apply the returned id (e.g. via
        // `set_cell_style_id`) will trigger any necessary recalculation.
        self.workbook.styles.intern(style)
    }

    /// Set the style id for a cell.
    ///
    /// Note: Unlike [`Engine::clear_cell`], this does **not** clear a cell's value/formula. Excel
    /// preserves formatting when editing contents, so this API only touches the cell's formatting
    /// metadata.
    pub fn set_cell_style_id(
        &mut self,
        sheet: &str,
        addr: &str,
        style_id: u32,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let addr = parse_a1(addr)?;
        if addr.row >= i32::MAX as u32 {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::RowOutOfRange,
            ));
        }
        let sheet_dims_changed = self.workbook.grow_sheet_dimensions(sheet_id, addr);
        if sheet_dims_changed {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
            // Sheet dimensions affect out-of-bounds `#REF!` semantics for references. Formatting
            // edits can grow the sheet (creating new in-bounds coordinates), so conservatively mark
            // compiled formulas dirty so results refresh on the next recalculation.
            self.mark_all_compiled_cells_dirty();
        }

        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let existing_style_id = self
            .workbook
            .get_cell(key)
            .map(|cell| cell.style_id)
            .unwrap_or(0);
        if existing_style_id != style_id {
            let remove_cell = {
                let cell = self.workbook.get_or_create_cell_mut(key);
                cell.style_id = style_id;
                cell.value == Value::Blank
                    && cell.formula.is_none()
                    && cell.style_id == 0
                    && cell.phonetic.is_none()
                    && cell.number_format.is_none()
            };
            if remove_cell {
                if let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) {
                    sheet.cells.remove(&addr);
                }
            }
        } else if style_id == 0 {
            // Prune empty default-style cells to keep sheet storage sparse.
            let remove_cell = self.workbook.get_cell(key).is_some_and(|cell| {
                cell.value == Value::Blank
                    && cell.formula.is_none()
                    && cell.style_id == 0
                    && cell.phonetic.is_none()
                    && cell.number_format.is_none()
            });
            if remove_cell {
                if let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) {
                    sheet.cells.remove(&addr);
                }
            }
        }
        // We avoid full-workbook dirtying for style changes in default full-precision mode because
        // only volatile metadata functions (CELL/INFO) consult formatting state. In Excel's
        // "precision as displayed" mode, formatting can affect stored numeric values, so retain the
        // conservative full-dirty behavior.
        let style_changed = existing_style_id != style_id;
        if style_changed && !sheet_dims_changed && !self.calc_settings.full_precision {
            self.mark_all_compiled_cells_dirty();
        }
        if (sheet_dims_changed || style_changed)
            && self.calc_settings.calculation_mode != CalculationMode::Manual
        {
            self.recalculate();
        }
        Ok(())
    }

    /// Bulk-apply per-cell style ids.
    ///
    /// This is more efficient than calling [`Engine::set_cell_style_id`] repeatedly because it:
    /// - avoids per-cell A1 parsing,
    /// - grows sheet dimensions once, and
    /// - (when in automatic calculation mode) recalculates at most once.
    pub fn set_cell_style_ids(
        &mut self,
        sheet: &str,
        writes: &[(formula_model::CellRef, u32)],
    ) -> Result<(), EngineError> {
        if writes.is_empty() {
            return Ok(());
        }

        let sheet_id = self.workbook.ensure_sheet(sheet);

        // Validate coordinates and ensure the sheet is large enough for all style targets.
        let mut max_row = 0u32;
        let mut max_col = 0u32;
        for (cell, _) in writes {
            if cell.row >= i32::MAX as u32 {
                return Err(EngineError::Address(
                    crate::eval::AddressParseError::RowOutOfRange,
                ));
            }
            if cell.col >= EXCEL_MAX_COLS {
                return Err(EngineError::Address(
                    crate::eval::AddressParseError::ColumnOutOfRange,
                ));
            }
            max_row = max_row.max(cell.row);
            max_col = max_col.max(cell.col);
        }

        let sheet_dims_changed = self.workbook.grow_sheet_dimensions(
            sheet_id,
            CellAddr {
                row: max_row,
                col: max_col,
            },
        );
        if sheet_dims_changed {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
            self.mark_all_compiled_cells_dirty();
        }

        let mut style_changed = false;
        for (cell, style_id) in writes {
            let addr = CellAddr {
                row: cell.row,
                col: cell.col,
            };
            let key = CellKey {
                sheet: sheet_id,
                addr,
            };

            let existing_style_id = self
                .workbook
                .get_cell(key)
                .map(|cell| cell.style_id)
                .unwrap_or(0);
            if existing_style_id == *style_id {
                continue;
            }
            style_changed = true;

            let remove_cell = {
                let cell = self.workbook.get_or_create_cell_mut(key);
                cell.style_id = *style_id;
                cell.value == Value::Blank
                    && cell.formula.is_none()
                    && cell.style_id == 0
                    && cell.phonetic.is_none()
                    && cell.number_format.is_none()
            };
            if remove_cell {
                if let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) {
                    sheet_state.cells.remove(&addr);
                }
            }
        }

        // In "precision as displayed" mode, formatting can affect stored numeric values for formula
        // cells, so conservatively refresh compiled results once after applying all style edits.
        if style_changed && !sheet_dims_changed && !self.calc_settings.full_precision {
            self.mark_all_compiled_cells_dirty();
        }

        if (sheet_dims_changed || style_changed)
            && self.calc_settings.calculation_mode != CalculationMode::Manual
        {
            self.recalculate();
        }

        Ok(())
    }

    /// Set (or clear) the explicit width override for a column.
    pub fn set_col_width(&mut self, sheet: &str, col_0based: u32, width: Option<f32>) {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let sheet_dims_changed = self.workbook.grow_sheet_dimensions(
            sheet_id,
            CellAddr {
                row: 0,
                col: col_0based,
            },
        );
        if sheet_dims_changed {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
            // Sheet dimensions affect out-of-bounds `#REF!` semantics for references. When a sheet
            // grows, formulas that previously evaluated to `#REF!` may now become valid, so
            // conservatively mark all compiled formulas dirty.
            self.mark_all_compiled_cells_dirty();
        }

        let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) else {
            return;
        };

        let before = sheet.col_properties.get(&col_0based).and_then(|p| p.width);
        sheet
            .col_properties
            .entry(col_0based)
            .and_modify(|p| p.width = width)
            .or_insert_with(|| ColProperties {
                width,
                hidden: false,
                style_id: None,
            });

        // Prune default entries to keep the map sparse.
        if let Some(props) = sheet.col_properties.get(&col_0based) {
            if props.width.is_none() && !props.hidden && props.style_id.is_none() {
                sheet.col_properties.remove(&col_0based);
            }
        }

        let after = sheet.col_properties.get(&col_0based).and_then(|p| p.width);
        let props_changed = before != after;

        // Column width metadata affects `CELL("width")`, which is volatile. When only the
        // per-column width changes (without altering sheet dimensions), a recalculation tick is
        // sufficient to refresh dependents via the volatile closure (no full-workbook dirtying).
        if (sheet_dims_changed || props_changed)
            && self.calc_settings.calculation_mode != CalculationMode::Manual
        {
            self.recalculate();
        }
    }

    /// Set whether a column is user-hidden.
    pub fn set_col_hidden(&mut self, sheet: &str, col_0based: u32, hidden: bool) {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let sheet_dims_changed = self.workbook.grow_sheet_dimensions(
            sheet_id,
            CellAddr {
                row: 0,
                col: col_0based,
            },
        );
        if sheet_dims_changed {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
            // Sheet dimension growth can affect out-of-bounds semantics; see `set_col_width`.
            self.mark_all_compiled_cells_dirty();
        }

        let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) else {
            return;
        };
        let before = sheet
            .col_properties
            .get(&col_0based)
            .map(|p| p.hidden)
            .unwrap_or(false);
        sheet
            .col_properties
            .entry(col_0based)
            .and_modify(|p| p.hidden = hidden)
            .or_insert_with(|| ColProperties {
                width: None,
                hidden,
                style_id: None,
            });

        // Prune default entries to keep the map sparse.
        if let Some(props) = sheet.col_properties.get(&col_0based) {
            if props.width.is_none() && !props.hidden && props.style_id.is_none() {
                sheet.col_properties.remove(&col_0based);
            }
        }

        let after = sheet
            .col_properties
            .get(&col_0based)
            .map(|p| p.hidden)
            .unwrap_or(false);
        let props_changed = before != after;

        // Hidden state affects `CELL("width")` (hidden columns return 0). Like `set_col_width`,
        // avoid full-workbook dirtying when only the metadata changes (rely on volatile closure).
        if (sheet_dims_changed || props_changed)
            && self.calc_settings.calculation_mode != CalculationMode::Manual
        {
            self.recalculate();
        }
    }

    /// Replace the set of formatting runs for a column.
    ///
    /// Runs are interpreted as row ranges `[start_row, end_row_exclusive)`.
    /// The run style layer has precedence `sheet < col < row < range-run < cell`.
    pub fn set_format_runs_by_col(
        &mut self,
        sheet: &str,
        col_0based: u32,
        mut runs: Vec<FormatRun>,
    ) -> Result<(), EngineError> {
        if col_0based >= EXCEL_MAX_COLS {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::ColumnOutOfRange,
            ));
        }

        // Validate rows are within the engine's supported bounds.
        for run in &runs {
            if run.start_row >= run.end_row_exclusive {
                return Err(EngineError::Address(
                    crate::eval::AddressParseError::RowOutOfRange,
                ));
            }
            if run.start_row >= i32::MAX as u32 || run.end_row_exclusive > i32::MAX as u32 {
                return Err(EngineError::Address(
                    crate::eval::AddressParseError::RowOutOfRange,
                ));
            }
        }

        // Keep deterministic ordering.
        runs.sort_by_key(|r| (r.start_row, r.end_row_exclusive, r.style_id));

        let sheet_id = self.workbook.ensure_sheet(sheet);

        // Ensure the sheet dimensions include the formatted column and any run endpoints.
        if self.workbook.grow_sheet_dimensions(
            sheet_id,
            CellAddr {
                row: 0,
                col: col_0based,
            },
        ) {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
            self.mark_all_compiled_cells_dirty();
        }
        if let Some(max_row) = runs
            .iter()
            .map(|r| r.end_row_exclusive.saturating_sub(1))
            .max()
        {
            if self.workbook.grow_sheet_dimensions(
                sheet_id,
                CellAddr {
                    row: max_row,
                    col: col_0based,
                },
            ) {
                self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
                self.mark_all_compiled_cells_dirty();
            }
        }

        {
            let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) else {
                return Ok(());
            };

            // Store only non-default style runs to keep the representation sparse.
            runs.retain(|r| r.style_id != 0);
            if runs.is_empty() {
                sheet.format_runs_by_col.remove(&col_0based);
            } else {
                sheet.format_runs_by_col.insert(col_0based, runs);
            }
        }

        // Formatting changes can affect worksheet information functions that consult formatting
        // (e.g. `CELL("prefix")`). In full-precision mode, those functions are volatile so a recalc
        // tick is sufficient. In "precision as displayed" mode, retain conservative full-dirty
        // semantics.
        if !self.calc_settings.full_precision {
            self.mark_all_compiled_cells_dirty();
        }
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }

        Ok(())
    }
    /// Set workbook file metadata (directory + filename).
    ///
    /// This is used by worksheet/workbook information functions like `CELL("filename")` and
    /// `INFO("directory")`.
    ///
    /// Hosts may not have access to an OS-level directory path (e.g. web environments). In those
    /// cases, callers can supply just a filename.
    ///
    /// Passing `None` (or an empty string) clears the corresponding field.
    pub fn set_workbook_file_metadata(&mut self, directory: Option<&str>, filename: Option<&str>) {
        let directory = directory.map(|s| s.to_string()).filter(|s| !s.is_empty());
        let filename = filename.map(|s| s.to_string()).filter(|s| !s.is_empty());

        if self.workbook.workbook_directory == directory
            && self.workbook.workbook_filename == filename
        {
            return;
        }

        self.workbook.workbook_directory = directory;
        self.workbook.workbook_filename = filename;

        // Workbook metadata can affect worksheet information function outputs (e.g.
        // `CELL("filename")` / `INFO("directory")`). These functions are volatile, so a
        // recalculation tick is sufficient to refresh dependents (no full-workbook dirtying).
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Update the user-visible display (tab) name for a sheet, without changing its stable key.
    ///
    /// The engine uses stable sheet keys for persistence and public APIs (e.g. `set_cell_value`),
    /// but Excel formulas and worksheet-info functions emit/resolve the display name.
    ///
    /// This is a metadata-only update: it does not rewrite stored formulas. Display names are only
    /// observed by volatile worksheet functions (e.g. `CELL("address")`) and runtime-parsed
    /// references (e.g. `INDIRECT(...)`), so a recalculation tick is sufficient to refresh
    /// dependents in automatic calculation modes (no full-workbook dirtying needed).
    ///
    /// Invalid Excel sheet names (e.g. containing `:` / `[]`, leading/trailing apostrophe, or names
    /// longer than 31 UTF-16 code units) are ignored and the workbook is left unchanged.
    pub fn set_sheet_display_name(&mut self, sheet_key: &str, display_name: &str) {
        let Some(sheet_id) = self.workbook.sheet_id_by_key(sheet_key) else {
            return;
        };
        let Some(old_display_name) = self.workbook.sheet_name(sheet_id).map(|s| s.to_string())
        else {
            return;
        };
        if !self.workbook.set_sheet_display_name(sheet_id, display_name) {
            return;
        }

        // Pivot definitions store sheet names as raw strings. If we leave stale references behind,
        // refreshing a pivot can silently resurrect the old display name via `ensure_sheet`.
        for pivot in self.workbook.pivots.values_mut() {
            if formula_model::sheet_name_eq_case_insensitive(
                &pivot.destination.sheet,
                &old_display_name,
            ) {
                pivot.destination.sheet = display_name.to_string();
            }
            if let PivotSource::Range { sheet, .. } = &mut pivot.source {
                if formula_model::sheet_name_eq_case_insensitive(sheet, &old_display_name) {
                    *sheet = display_name.to_string();
                }
            }
        }

        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Delete a worksheet from the workbook and rewrite any remaining formulas/defined names that
    /// referenced it.
    ///
    /// This matches Excel's sheet delete semantics:
    /// - Local references to the deleted sheet become `#REF!`.
    /// - 3D references (`Sheet1:Sheet3!A1`) shift their boundaries inward when the deleted sheet was
    ///   a boundary.
    /// - External workbook references (`[Book.xlsx]Sheet1!A1`) are **not** rewritten.
    pub fn delete_sheet(&mut self, sheet: &str) -> Result<(), EngineError> {
        let Some(deleted_sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(());
        };
        if self.workbook.sheet_order.len() <= 1 {
            return Err(EngineError::CannotDeleteLastSheet);
        }

        let deleted_sheet_key = self
            .workbook
            .sheet_key_name(deleted_sheet_id)
            .unwrap_or(sheet)
            .to_string();
        let deleted_sheet_display_name = self
            .workbook
            .sheet_name(deleted_sheet_id)
            .unwrap_or(sheet)
            .to_string();

        let deleted_sheet_key_norm = Workbook::sheet_key(&deleted_sheet_key);
        let deleted_sheet_display_key_norm = Workbook::sheet_key(&deleted_sheet_display_name);

        // Pivot definitions store sheet names as strings. If we leave stale references behind,
        // refreshing a pivot can silently resurrect the deleted sheet via `ensure_sheet`.
        //
        // Also drop table-backed pivots that referenced tables from the deleted sheet.
        let deleted_table_ids: HashSet<u32> = self
            .workbook
            .sheets
            .get(deleted_sheet_id)
            .map(|sheet_state| sheet_state.tables.iter().map(|t| t.id).collect())
            .unwrap_or_default();
        self.workbook.pivots.retain(|_, def| {
            let dest_key = Workbook::sheet_key(&def.destination.sheet);
            let destination_matches =
                dest_key == deleted_sheet_key_norm || dest_key == deleted_sheet_display_key_norm;
            let source_matches = match &def.source {
                PivotSource::Range { sheet, .. } => {
                    let key = Workbook::sheet_key(sheet);
                    key == deleted_sheet_key_norm || key == deleted_sheet_display_key_norm
                }
                PivotSource::Table { table_id } => deleted_table_ids.contains(table_id),
            };
            !(destination_matches || source_matches)
        });

        // Drop any registered pivot metadata associated with the deleted sheet.
        self.pivot_registry.prune_sheet(deleted_sheet_id);

        // Drop any per-sheet engine metadata keyed by the deleted sheet id.
        self.info.origin_by_sheet.remove(&deleted_sheet_id);
        // Keep the pre-delete sheet tab order so 3D span boundary shift logic can resolve adjacent
        // sheets (Excel shifts a deleted 3D boundary one sheet inward).
        //
        // Note: Excel formulas are defined in terms of the user-visible sheet display names, but
        // the engine also supports referencing sheets by their stable key. We capture *both* so we
        // can invalidate references regardless of which name form appears in the stored formula
        // text.
        let sheet_order_keys: Vec<String> = self
            .workbook
            .sheet_ids_in_order()
            .iter()
            .filter_map(|&id| {
                self.workbook
                    .sheet_key_name(id)
                    .map(|name| name.to_string())
            })
            .collect();
        let sheet_order_display_names: Vec<String> = self
            .workbook
            .sheet_ids_in_order()
            .iter()
            .filter_map(|&id| self.workbook.sheet_name(id).map(|name| name.to_string()))
            .collect();
        // Formulas can reference sheets by either their user-visible display name (Excel semantics)
        // or by their stable sheet key (backward compatibility / host-provided identifiers). When
        // deleting a sheet, rewrite both forms so references cannot resurrect if a sheet with the
        // same name/key is later re-created.
        let rewrite_deleted_sheet_formula =
            |formula: &str, origin: crate::CellAddr| -> Option<String> {
                let (rewritten, changed) = rewrite_formula_for_sheet_delete_with_aliases(
                    formula,
                    origin,
                    &deleted_sheet_key,
                    &deleted_sheet_display_name,
                    &sheet_order_keys,
                    &sheet_order_display_names,
                );
                changed.then_some(rewritten)
            };

        // Mark the sheet as deleted while keeping its id stable.
        self.workbook
            .sheet_key_to_id
            .remove(&deleted_sheet_key_norm);
        self.workbook
            .sheet_display_name_to_id
            .remove(&deleted_sheet_display_key_norm);
        if let Some(name) = self.workbook.sheet_keys.get_mut(deleted_sheet_id) {
            *name = None;
        }
        if let Some(name) = self.workbook.sheet_display_names.get_mut(deleted_sheet_id) {
            *name = None;
        }
        self.workbook
            .sheet_order
            .retain(|&id| id != deleted_sheet_id);
        self.workbook.rebuild_sheet_tab_index_by_id();

        // Drop any sheet-scoped state for the deleted worksheet.
        if let Some(sheet_state) = self.workbook.sheets.get_mut(deleted_sheet_id) {
            sheet_state.cells.clear();
            sheet_state.tables.clear();
            sheet_state.names.clear();
            sheet_state.default_style_id = None;
            sheet_state.row_properties.clear();
            sheet_state.col_properties.clear();
            sheet_state.format_runs_by_col.clear();
        }

        // Rewrite formulas stored in remaining sheets.
        let remaining_sheet_ids = self.workbook.sheet_ids_in_order().to_vec();
        for sheet_id in &remaining_sheet_ids {
            let Some(sheet) = self.workbook.sheets.get_mut(*sheet_id) else {
                continue;
            };
            for (addr, cell) in sheet.cells.iter_mut() {
                let Some(formula) = cell.formula.as_deref() else {
                    continue;
                };
                let origin = crate::CellAddr::new(addr.row, addr.col);
                if let Some(rewritten) = rewrite_deleted_sheet_formula(formula, origin) {
                    cell.formula = Some(rewritten.into());
                }
            }

            // Rewrite table calculated column + totals formulas.
            for table in &mut sheet.tables {
                for column in &mut table.columns {
                    if let Some(formula) = column.formula.as_mut() {
                        if let Some(rewritten) =
                            rewrite_deleted_sheet_formula(formula, crate::CellAddr::new(0, 0))
                        {
                            *formula = rewritten;
                        }
                    }
                    if let Some(formula) = column.totals_formula.as_mut() {
                        if let Some(rewritten) =
                            rewrite_deleted_sheet_formula(formula, crate::CellAddr::new(0, 0))
                        {
                            *formula = rewritten;
                        }
                    }
                }
            }
        }

        // Rewrite workbook-scoped names.
        for def in self.workbook.names.values_mut() {
            let formula = match &mut def.definition {
                NameDefinition::Constant(_) => continue,
                NameDefinition::Reference(formula) | NameDefinition::Formula(formula) => formula,
            };
            let origin = crate::CellAddr::new(0, 0);
            if let Some(rewritten) = rewrite_deleted_sheet_formula(formula, origin) {
                *formula = rewritten;
            }
        }

        // Rewrite sheet-scoped names in remaining sheets.
        for sheet_id in &remaining_sheet_ids {
            let Some(sheet) = self.workbook.sheets.get_mut(*sheet_id) else {
                continue;
            };
            for def in sheet.names.values_mut() {
                let formula = match &mut def.definition {
                    NameDefinition::Constant(_) => continue,
                    NameDefinition::Reference(formula) | NameDefinition::Formula(formula) => {
                        formula
                    }
                };
                let origin = crate::CellAddr::new(0, 0);
                if let Some(rewritten) = rewrite_deleted_sheet_formula(formula, origin) {
                    *formula = rewritten;
                }
            }
        }

        // Sheet deletion changes tab order and can invalidate references; rebuild compiled
        // representations to ensure 3D spans and dependencies refresh correctly.
        self.recompile_all_defined_names()?;
        self.rebuild_graph()?;

        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }
    /// Returns an immutable view of the tables defined on `sheet`.
    ///
    /// Tables are needed to resolve structured references like `Table1[Col]` and `[@Col]`.
    pub fn sheet_tables(&self, sheet: &str) -> Option<&[Table]> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        Some(self.workbook.sheets.get(sheet_id)?.tables.as_slice())
    }

    /// Rename a worksheet and rewrite formulas that reference it (Excel-like).
    ///
    /// Returns `false` if `old_name` does not exist or `new_name` is invalid/conflicts with another
    /// sheet.
    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> bool {
        let Some(sheet_id) = self.workbook.sheet_id(old_name) else {
            return false;
        };
        self.rename_sheet_by_id(sheet_id, new_name).is_ok()
    }

    /// Delete a worksheet by its stable [`SheetId`].
    ///
    /// This matches [`Engine::delete_sheet`] semantics (including formula rewrites and dependency
    /// graph rebuilds).
    ///
    /// If `id` is invalid or already deleted, this is a no-op and returns `Ok(())`.
    pub fn delete_sheet_by_id(&mut self, id: SheetId) -> Result<(), SheetLifecycleError> {
        if !self.workbook.sheet_exists(id) {
            return Ok(());
        }

        let name = self
            .workbook
            .sheet_key_name(id)
            .or_else(|| self.workbook.sheet_name(id))
            .expect("sheet exists")
            .to_string();
        self.delete_sheet(&name).map_err(|e| match e {
            EngineError::CannotDeleteLastSheet => SheetLifecycleError::CannotDeleteLastSheet,
            other => SheetLifecycleError::Internal(other.to_string()),
        })
    }
    /// Store a pivot table definition in the engine and return its allocated id.
    ///
    /// The engine will automatically rewrite pivot source/destination references on structural
    /// edits (insert/delete/move ranges).
    pub fn add_pivot_table(&mut self, mut def: PivotTableDefinition) -> PivotTableId {
        let id = self.workbook.next_pivot_id;
        self.workbook.next_pivot_id = self.workbook.next_pivot_id.wrapping_add(1);
        def.id = id;
        self.workbook.pivots.insert(id, def);
        id
    }

    pub fn pivot_table(&self, pivot_id: PivotTableId) -> Option<&PivotTableDefinition> {
        self.workbook.pivots.get(&pivot_id)
    }

    pub fn pivot_table_mut(&mut self, pivot_id: PivotTableId) -> Option<&mut PivotTableDefinition> {
        self.workbook.pivots.get_mut(&pivot_id)
    }

    /// Refresh a pivot table by id, writing the computed output into its destination worksheet.
    pub fn refresh_pivot_table(
        &mut self,
        pivot_id: PivotTableId,
    ) -> Result<PivotRefreshOutput, PivotRefreshError> {
        // `refresh_pivot` uses per-cell write/clear calls. In automatic calculation mode, those
        // would trigger recalculation repeatedly, which is prohibitively expensive for large pivot
        // outputs. Temporarily force manual mode so we can recalculate at most once at the end.
        let previous = self.calc_settings.clone();
        let forced_manual = previous.calculation_mode != CalculationMode::Manual;
        if forced_manual {
            let mut manual = previous.clone();
            manual.calculation_mode = CalculationMode::Manual;
            self.set_calc_settings(manual);
        }

        let result = self.refresh_pivot_table_internal(pivot_id);

        if forced_manual {
            self.set_calc_settings(previous.clone());
            // Best-effort: the refresh may have partially written output before returning an error.
            // Recalculate once so downstream formulas see a coherent state.
            if previous.calculation_mode != CalculationMode::Manual {
                self.recalculate();
            }
        }

        result
    }

    /// Refresh all pivot tables in deterministic workbook order, writing their outputs into the
    /// destination worksheets.
    pub fn refresh_all_pivots(
        &mut self,
    ) -> Result<Vec<(PivotTableId, PivotRefreshOutput)>, PivotRefreshError> {
        // Determine refresh order before mutating the pivot map.
        let tab_index = self.workbook.tab_index_by_sheet_id();
        let mut ids: Vec<PivotTableId> = self.workbook.pivots.keys().copied().collect();
        ids.sort_by_key(|id| {
            let Some(pivot) = self.workbook.pivots.get(id) else {
                return (usize::MAX, u32::MAX, u32::MAX, *id);
            };
            let sheet_id = self
                .workbook
                .sheet_id(&pivot.destination.sheet)
                .unwrap_or(usize::MAX);
            let tab = tab_index.get(sheet_id).copied().unwrap_or(usize::MAX);
            (
                tab,
                pivot.destination.cell.row,
                pivot.destination.cell.col,
                *id,
            )
        });

        let previous = self.calc_settings.clone();
        let forced_manual = previous.calculation_mode != CalculationMode::Manual;
        if forced_manual {
            let mut manual = previous.clone();
            manual.calculation_mode = CalculationMode::Manual;
            self.set_calc_settings(manual);
        }

        let mut outputs = Vec::with_capacity(ids.len());
        for id in ids {
            let out = self.refresh_pivot_table_internal(id)?;
            outputs.push((id, out));
        }

        if forced_manual {
            self.set_calc_settings(previous.clone());
            if previous.calculation_mode != CalculationMode::Manual {
                self.recalculate();
            }
        }

        Ok(outputs)
    }

    fn refresh_pivot_table_internal(
        &mut self,
        pivot_id: PivotTableId,
    ) -> Result<PivotRefreshOutput, PivotRefreshError> {
        let mut def = self
            .workbook
            .pivots
            .remove(&pivot_id)
            .ok_or(PivotRefreshError::UnknownPivot(pivot_id))?;
        let result = refresh_pivot(self, &mut def);
        self.workbook.pivots.insert(pivot_id, def);
        result
    }

    /// Returns the configured worksheet dimensions for `sheet` (row/column count).
    ///
    /// When unset, sheets default to Excel-compatible dimensions
    /// (`EXCEL_MAX_ROWS` x `EXCEL_MAX_COLS`).
    pub fn sheet_dimensions(&self, sheet: &str) -> Option<(u32, u32)> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let sheet = self.workbook.sheets.get(sheet_id)?;
        Some((sheet.row_count, sheet.col_count))
    }

    /// Set (or clear) the sheet's default column width in Excel "character" units.
    ///
    /// This is surfaced to worksheet information functions like `CELL("width")` and corresponds to
    /// the worksheet's OOXML `<sheetFormatPr defaultColWidth="...">` attribute.
    pub fn set_sheet_default_col_width(&mut self, sheet: &str, width: Option<f32>) {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let Some(s) = self.workbook.sheets.get_mut(sheet_id) else {
            return;
        };

        if s.default_col_width == width {
            return;
        }
        s.default_col_width = width;

        // Default column width metadata can affect `CELL("width")` outputs (and any downstream
        // calculations that depend on those values). `CELL()` is volatile, so a recalculation tick
        // is sufficient to refresh dependents via the volatile closure (no full-workbook dirtying).
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Enable/disable worksheet protection.
    ///
    /// This is currently informational only; it does not enforce edits. Worksheet information
    /// functions like `CELL("protect")` must report the cell's locked formatting state regardless
    /// of whether sheet protection is enabled (Excel behavior).
    pub fn set_sheet_protection_enabled(&mut self, sheet: &str, enabled: bool) {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) else {
            return;
        };
        if sheet.sheet_protection_enabled == enabled {
            return;
        }
        sheet.sheet_protection_enabled = enabled;
        // Worksheet protection is currently informational only and does not affect formula
        // evaluation, so no recalculation is required.
    }

    /// Returns whether worksheet protection is enabled for `sheet`.
    pub fn sheet_protection_enabled(&self, sheet: &str) -> Option<bool> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        self.workbook
            .sheets
            .get(sheet_id)
            .map(|sheet| sheet.sheet_protection_enabled)
    }

    /// Configure the logical worksheet grid size for `sheet`.
    ///
    /// This affects whole-row/whole-column references like `1:1` and `A:A`, which are resolved
    /// against the sheet's configured dimensions.
    ///
    /// Notes:
    /// - `col_count` is limited to Excel's 16,384-column maximum.
    /// - `row_count` is limited to `i32::MAX` for now because several internal evaluation paths
    ///   (notably the bytecode engine) use 32-bit coordinates.
    pub fn set_sheet_dimensions(
        &mut self,
        sheet: &str,
        row_count: u32,
        col_count: u32,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);

        if row_count == 0 {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::RowOutOfRange,
            ));
        }
        if col_count == 0 {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::ColumnOutOfRange,
            ));
        }
        if col_count > EXCEL_MAX_COLS {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::ColumnOutOfRange,
            ));
        }
        if row_count > i32::MAX as u32 {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::RowOutOfRange,
            ));
        }

        // Prevent shrinking below existing populated/spill cells. This avoids having "stored" cells
        // that can no longer be addressed within the sheet's configured bounds.
        if let Some(sheet_state) = self.workbook.sheets.get(sheet_id) {
            for addr in sheet_state.cells.keys() {
                if addr.row >= row_count {
                    return Err(EngineError::Address(
                        crate::eval::AddressParseError::RowOutOfRange,
                    ));
                }
                if addr.col >= col_count {
                    return Err(EngineError::Address(
                        crate::eval::AddressParseError::ColumnOutOfRange,
                    ));
                }
            }
            for (origin, spill) in &self.spills.by_origin {
                if origin.sheet != sheet_id {
                    continue;
                }
                if spill.end.row >= row_count {
                    return Err(EngineError::Address(
                        crate::eval::AddressParseError::RowOutOfRange,
                    ));
                }
                if spill.end.col >= col_count {
                    return Err(EngineError::Address(
                        crate::eval::AddressParseError::ColumnOutOfRange,
                    ));
                }
            }
        }

        let mut changed = false;
        if let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) {
            changed = sheet_state.row_count != row_count || sheet_state.col_count != col_count;
            sheet_state.row_count = row_count;
            sheet_state.col_count = col_count;
        }
        if changed {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
        }
        if !changed {
            return Ok(());
        }

        // Changing sheet dimensions affects:
        // - Out-of-bounds `#REF!` semantics for references
        // - Whole-row/whole-column references (`A:A`, `1:1`) inside bytecode programs (range
        //   endpoints are expanded during compilation)
        //
        // Mark all compiled formula cells dirty so results refresh on the next recalculation, and
        // bump `sheet_dims_generation` so stale bytecode programs fall back to AST evaluation.
        self.mark_all_compiled_cells_dirty();

        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }

        Ok(())
    }

    /// Set the host-provided top-left visible cell ("origin") for `sheet`.
    ///
    /// Excel's `INFO("origin")` is tied to the active window's view state (scroll position +
    /// frozen panes). The core engine is deterministic and does not query UI state directly, so
    /// hosts should provide the current origin explicitly.
    ///
    /// When the origin changes, any formulas that depend on `INFO("origin")` are marked dirty so
    /// the next recalculation observes the updated view state.
    pub fn set_sheet_origin(
        &mut self,
        sheet: &str,
        origin: Option<&str>,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);

        let origin = origin.map(str::trim).filter(|s| !s.is_empty());
        let origin = match origin {
            Some(addr) => Some(parse_a1(addr)?),
            None => None,
        };

        // Reject out-of-bounds origin coordinates. This keeps `INFO("origin")` deterministic and
        // avoids fabricating invalid absolute A1 strings.
        if let Some(addr) = origin {
            if let Some(sheet_state) = self.workbook.sheets.get(sheet_id) {
                if addr.row >= sheet_state.row_count {
                    return Err(EngineError::Address(
                        crate::eval::AddressParseError::RowOutOfRange,
                    ));
                }
                if addr.col >= sheet_state.col_count {
                    return Err(EngineError::Address(
                        crate::eval::AddressParseError::ColumnOutOfRange,
                    ));
                }
            }
        }

        let dependents: Vec<CellKey> = {
            let sheet_state = self
                .workbook
                .sheets
                .get_mut(sheet_id)
                .expect("sheet just ensured must exist");
            if sheet_state.origin == origin {
                return Ok(());
            }
            sheet_state.origin = origin;
            sheet_state
                .origin_dependents
                .iter()
                .copied()
                .map(|addr| CellKey {
                    sheet: sheet_id,
                    addr,
                })
                .collect()
        };

        for key in dependents {
            let cell_id = cell_id_from_key(key);
            self.mark_dirty_including_self_with_reasons(key);
            self.calc_graph.mark_dirty(cell_id);
        }
        self.sync_dirty_from_calc_graph();

        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }

        Ok(())
    }

    /// Return the workbook style table (interned style objects).
    pub fn style_table(&self) -> &StyleTable {
        &self.workbook.styles
    }

    /// Replace the workbook style table.
    ///
    /// This is primarily intended for workbook load flows (XLSX import, persistence hydrate) so
    /// style ids referenced by cells/rows/cols resolve consistently during formula evaluation (e.g.
    /// `CELL("protect")`).
    pub fn set_style_table(&mut self, styles: StyleTable) {
        self.workbook.styles = styles;

        // Formatting metadata affects worksheet information functions like `CELL("format")`, but
        // those functions are volatile so a recalculation tick is sufficient in the default
        // full-precision mode.
        //
        // In "precision as displayed" mode, formatting can affect stored numeric values at formula
        // boundaries, so retain the conservative full-dirty behavior.
        if !self.calc_settings.full_precision {
            self.mark_all_compiled_cells_dirty();
        }
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Set the worksheet default style id (layered formatting base) for `sheet`.
    ///
    /// `None` (or `Some(0)`) resets the default style to `0`.
    pub fn set_sheet_default_style_id(&mut self, sheet: &str, style_id: Option<u32>) {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let style_id = style_id.filter(|id| *id != 0);

        let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) else {
            return;
        };
        if sheet_state.default_style_id == style_id {
            return;
        }
        sheet_state.default_style_id = style_id;

        if !self.calc_settings.full_precision {
            self.mark_all_compiled_cells_dirty();
        }
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Set a row formatting style layer for `sheet` at 0-based `row0`.
    ///
    /// `None` (or `Some(0)`) clears the row style.
    pub fn set_row_style_id(&mut self, sheet: &str, row0: u32, style_id: Option<u32>) {
        // Keep row indices within the engine's supported coordinate space.
        if row0 >= i32::MAX as u32 {
            return;
        }

        let sheet_id = self.workbook.ensure_sheet(sheet);
        let style_id = style_id.filter(|id| *id != 0);
        let dims_changed = style_id
            .map(|_| {
                self.workbook
                    .grow_sheet_dimensions(sheet_id, CellAddr { row: row0, col: 0 })
            })
            .unwrap_or(false);

        if dims_changed {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
        }

        let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) else {
            return;
        };

        let prev = sheet_state
            .row_properties
            .get(&row0)
            .and_then(|p| p.style_id)
            .filter(|id| *id != 0);
        if prev == style_id && !dims_changed {
            return;
        }

        if prev != style_id {
            match style_id {
                Some(id) => {
                    sheet_state.row_properties.entry(row0).or_default().style_id = Some(id);
                }
                None => {
                    if let Some(props) = sheet_state.row_properties.get_mut(&row0) {
                        props.style_id = None;
                    }
                }
            }
        }

        // Keep the map sparse when no other per-row properties are set.
        if let Some(props) = sheet_state.row_properties.get(&row0) {
            if *props == RowProperties::default() {
                sheet_state.row_properties.remove(&row0);
            }
        }
        if dims_changed || !self.calc_settings.full_precision {
            self.mark_all_compiled_cells_dirty();
        }
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Set a column formatting style layer for `sheet` at 0-based `col0`.
    ///
    /// `None` (or `Some(0)`) clears the column style.
    pub fn set_col_style_id(&mut self, sheet: &str, col0: u32, style_id: Option<u32>) {
        // The engine enforces Excel's fixed 16,384-column grid.
        if col0 >= EXCEL_MAX_COLS {
            return;
        }

        let sheet_id = self.workbook.ensure_sheet(sheet);
        let style_id = style_id.filter(|id| *id != 0);
        let dims_changed = style_id
            .map(|_| {
                self.workbook
                    .grow_sheet_dimensions(sheet_id, CellAddr { row: 0, col: col0 })
            })
            .unwrap_or(false);

        if dims_changed {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
        }

        let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) else {
            return;
        };

        let prev = sheet_state
            .col_properties
            .get(&col0)
            .and_then(|p| p.style_id)
            .filter(|id| *id != 0);
        if prev == style_id && !dims_changed {
            return;
        }

        if prev != style_id {
            match style_id {
                Some(id) => {
                    sheet_state.col_properties.entry(col0).or_default().style_id = Some(id);
                }
                None => {
                    if let Some(props) = sheet_state.col_properties.get_mut(&col0) {
                        props.style_id = None;
                    }
                }
            }
        }

        // Keep the map sparse when no other per-column properties are set.
        if let Some(props) = sheet_state.col_properties.get(&col0) {
            if *props == ColProperties::default() {
                sheet_state.col_properties.remove(&col0);
            }
        }

        if dims_changed || !self.calc_settings.full_precision {
            self.mark_all_compiled_cells_dirty();
        }
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    pub fn set_calc_settings(&mut self, settings: CalcSettings) {
        self.calc_settings = settings;
    }

    fn fmt_date_system(&self) -> FmtDateSystem {
        match self.date_system {
            ExcelDateSystem::Excel1900 { .. } => FmtDateSystem::Excel1900,
            ExcelDateSystem::Excel1904 => FmtDateSystem::Excel1904,
        }
    }

    fn fmt_options(&self) -> FmtFormatOptions {
        FmtFormatOptions {
            locale: self.value_locale.separators,
            date_system: self.fmt_date_system(),
        }
    }

    // Style-layer resolution helper retained for upcoming formatting APIs. This is currently not
    // referenced by the engine core, but keeping it here makes the intended precedence rules
    // explicit and avoids duplicating the logic across future call sites.
    #[allow(dead_code)]
    fn effective_style_id_at(&self, key: CellKey) -> u32 {
        // If this cell is part of a spilled array, use the spill origin's formatting (Excel
        // displays spilled outputs using the origin cell's formatting).
        if let Some(origin) = self.spill_origin_key(key) {
            if origin != key {
                return self.effective_style_id_at(origin);
            }
        }

        // Prefer an explicit cell style. Treat style_id 0 (default style) as "inherit" so
        // row/column/sheet default styles can still apply.
        let cell_style_id = self
            .workbook
            .get_cell(key)
            .map(|cell| cell.style_id)
            .unwrap_or(0);
        if cell_style_id != 0 {
            return cell_style_id;
        }

        let Some(sheet_state) = self.workbook.sheets.get(key.sheet) else {
            return 0;
        };

        if let Some(style_id) = sheet_state
            .row_properties
            .get(&key.addr.row)
            .and_then(|props| props.style_id)
            .filter(|id| *id != 0)
        {
            return style_id;
        }

        if let Some(style_id) = sheet_state
            .col_properties
            .get(&key.addr.col)
            .and_then(|props| props.style_id)
            .filter(|id| *id != 0)
        {
            return style_id;
        }

        if let Some(style_id) = sheet_state.default_style_id.filter(|id| *id != 0) {
            return style_id;
        }

        0
    }

    /// Resolve the effective number format pattern for rounding in "precision as displayed" mode.
    ///
    /// Resolution order:
    /// 1. Explicit [`Cell::number_format`] override (if set).
    /// 2. Spill-origin formatting semantics (spilled outputs observe the origin cell's formatting).
    /// 3. Effective number format resolved from layered style ids:
    ///    sheet < col < row < range-run < cell (per-property).
    fn number_format_pattern_for_rounding(&self, key: CellKey) -> Option<&str> {
        // Explicit per-cell override always wins.
        if let Some(fmt) = self
            .workbook
            .get_cell(key)
            .and_then(|cell| cell.number_format.as_deref())
        {
            return Some(fmt);
        }

        // If this cell is a spilled output, use the spill origin's formatting.
        if let Some(origin) = self.spill_origin_key(key) {
            if origin != key {
                return self.number_format_pattern_for_rounding(origin);
            }
        }

        let Some(sheet_state) = self.workbook.sheets.get(key.sheet) else {
            return None;
        };

        let cell_style_id = self
            .workbook
            .get_cell(key)
            .map(|cell| cell.style_id)
            .unwrap_or(0);
        let row_style_id = sheet_state
            .row_properties
            .get(&key.addr.row)
            .and_then(|props| props.style_id)
            .unwrap_or(0);
        let col_style_id = sheet_state
            .col_properties
            .get(&key.addr.col)
            .and_then(|props| props.style_id)
            .unwrap_or(0);
        let run_style_id = style_id_for_row_in_runs(
            sheet_state
                .format_runs_by_col
                .get(&key.addr.col)
                .map(|runs| runs.as_slice()),
            key.addr.row,
        );
        let sheet_style_id = sheet_state.default_style_id.unwrap_or(0);

        // Style precedence matches DocumentController layering:
        // sheet < col < row < range-run < cell
        //
        // When a style does not specify a number format (`number_format=None`), it is treated as
        // "inherit" so lower-precedence layers can contribute the number format.
        for style_id in [
            cell_style_id,
            run_style_id,
            row_style_id,
            col_style_id,
            sheet_style_id,
        ] {
            if style_id == 0 {
                continue;
            }
            if let Some(fmt) = self
                .workbook
                .styles
                .get(style_id)
                .and_then(|style| style.number_format.as_deref())
            {
                return Some(fmt);
            }
        }

        None
    }

    fn round_number_as_displayed(&self, number: f64, format_pattern: Option<&str>) -> f64 {
        if self.calc_settings.full_precision {
            return number;
        }

        // Excel's "precision as displayed" mode ("Set precision as displayed") rounds numeric
        // values at cell boundaries based on the cell's number format.
        //
        // We implement this by:
        // 1) Formatting the number using `formula-format` (Excel-compatible formatting),
        // 2) Parsing the formatted text back into a number using the engine's numeric coercion
        //    logic (locale-aware, percent-aware).
        //
        // If the formatted string cannot be parsed back into a number (e.g. date/time formats or
        // patterns with non-numeric literal text), we fall back to storing the full-precision value.
        let options = self.fmt_options();
        let fmt_value = FmtValue::Number(number);
        let formatted = formula_format::format_value(fmt_value, format_pattern, &options);
        match crate::coercion::number::parse_number_strict(
            &formatted.text,
            options.locale.decimal_sep,
            Some(options.locale.thousands_sep),
        ) {
            Ok(parsed) => parsed,
            Err(_) => number,
        }
    }

    pub fn locale_config(&self) -> &crate::LocaleConfig {
        &self.locale_config
    }

    pub fn set_locale_config(&mut self, locale: crate::LocaleConfig) {
        if self.locale_config == locale {
            return;
        }
        self.locale_config = locale;
        self.mark_all_compiled_cells_dirty();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Configure an [`ExternalValueProvider`].
    ///
    /// This provider is used both for:
    /// - **External workbook references** like `=[Book.xlsx]Sheet1!A1`.
    /// - **Out-of-band values** for the current workbook when a cell is not present in the engine's
    ///   internal grid storage (useful for streaming/virtualized sheets).
    ///
    /// See [`ExternalValueProvider`] for the canonical external sheet-key formats and
    /// [`ExternalValueProvider::sheet_order`] semantics used for expanding external 3D sheet spans.
    pub fn set_external_value_provider(
        &mut self,
        provider: Option<Arc<dyn ExternalValueProvider>>,
    ) {
        self.external_value_provider = provider;
    }

    pub fn set_external_data_provider(&mut self, provider: Option<Arc<dyn ExternalDataProvider>>) {
        self.external_data_provider = provider;
    }

    /// Returns whether external workbook references are treated as volatile roots.
    ///
    /// See [`Engine::set_external_refs_volatile`] for details.
    pub fn external_refs_volatile(&self) -> bool {
        self.external_refs_volatile
    }

    /// Configure whether external workbook references (e.g. `[Book.xlsx]Sheet1!A1`) are treated as
    /// volatile.
    ///
    /// - When enabled (the default), formulas containing external references are re-evaluated on
    ///   every recalculation pass (Excel-compatible behavior).
    /// - When disabled, external references are considered non-volatile and will only refresh after
    ///   explicit invalidation via [`Engine::mark_external_sheet_dirty`] /
    ///   [`Engine::mark_external_workbook_dirty`] or when the formula is otherwise marked dirty.
    pub fn set_external_refs_volatile(&mut self, volatile: bool) {
        if self.external_refs_volatile == volatile {
            return;
        }
        self.external_refs_volatile = volatile;

        // Refresh per-cell volatile flags and dependency-graph volatile roots. This is a rare,
        // engine-wide configuration knob, so scanning formula cells is acceptable.
        let tables_by_sheet: Vec<Vec<Table>> = self
            .workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if self.workbook.sheet_exists(sheet_id) {
                    s.tables.clone()
                } else {
                    Vec::new()
                }
            })
            .collect();

        let mut updates: Vec<(CellKey, bool)> = Vec::new();
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            if !self.workbook.sheet_exists(sheet_id) {
                continue;
            }
            for (addr, cell) in &sheet.cells {
                let Some(compiled) = cell.compiled.as_ref() else {
                    continue;
                };
                let key = CellKey {
                    sheet: sheet_id,
                    addr: *addr,
                };
                let (_, is_volatile, _, _, _) = analyze_expr_flags(
                    compiled.ast(),
                    key,
                    &tables_by_sheet,
                    &self.workbook,
                    self.external_refs_volatile,
                );
                if is_volatile != cell.volatile {
                    updates.push((key, is_volatile));
                }
            }
        }

        for (key, is_volatile) in updates {
            if let Some(cell) = self
                .workbook
                .sheets
                .get_mut(key.sheet)
                .and_then(|s| s.cells.get_mut(&key.addr))
            {
                cell.volatile = is_volatile;
            }
            self.calc_graph
                .set_cell_volatile(cell_id_from_key(key), is_volatile);
        }

        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Explicitly mark all formulas that reference the given external sheet key as dirty.
    ///
    /// `sheet_key` must be in canonical external sheet form, e.g. `"[Book.xlsx]Sheet1"`.
    ///
    /// External-workbook 3D span references like `"[Book.xlsx]Sheet1:Sheet3"` are expanded to their
    /// component per-sheet keys for invalidation when sheet order is available; when it is not, the
    /// raw span key is tracked as a single dependency and can be invalidated directly.
    pub fn mark_external_sheet_dirty(&mut self, sheet_key: &str) {
        let sheet_key = sheet_key.trim();
        let mut cells: HashSet<CellKey> = HashSet::new();
        if let Some(static_cells) = self.external_sheet_dependents.get(sheet_key) {
            cells.extend(static_cells.iter().copied());
        }
        if let Some(dynamic_cells) = self.dynamic_external_sheet_dependents.get(sheet_key) {
            cells.extend(dynamic_cells.iter().copied());
        }
        if cells.is_empty() {
            return;
        }

        for key in cells {
            self.dirty_reasons.remove(&key);
            self.calc_graph.mark_dirty(cell_id_from_key(key));
        }
        self.sync_dirty_from_calc_graph();

        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Explicitly mark all formulas that reference any sheet within `workbook` as dirty.
    ///
    /// `workbook` should match the workbook component of an external sheet key, e.g. `"Book.xlsx"`
    /// for the sheet key `"[Book.xlsx]Sheet1"`.
    pub fn mark_external_workbook_dirty(&mut self, workbook: &str) {
        let mut workbook = workbook.trim();

        // Normalize a few common caller-provided forms:
        // - the workbook component itself: `Book.xlsx`
        // - a bracket-wrapped workbook id: `[Book.xlsx]` (strip wrapper brackets)
        //
        // Workbook names can contain literal `[` characters and escape literal `]` characters as
        // `]]`. Avoid naïvely stripping leading/trailing brackets when the brackets are part of the
        // workbook id itself (e.g. a filename like `[Book]` yields a workbook id of `"[Book]]"`).
        if let Some(inner) = Self::strip_wrapping_workbook_brackets(workbook) {
            workbook = inner;
        }

        let mut cells: HashSet<CellKey> = HashSet::new();
        if let Some(static_cells) = self.external_workbook_dependents.get(workbook) {
            cells.extend(static_cells.iter().copied());
        }
        if let Some(dynamic_cells) = self.dynamic_external_workbook_dependents.get(workbook) {
            cells.extend(dynamic_cells.iter().copied());
        }
        if cells.is_empty() {
            return;
        }

        for key in cells {
            self.dirty_reasons.remove(&key);
            self.calc_graph.mark_dirty(cell_id_from_key(key));
        }
        self.sync_dirty_from_calc_graph();

        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Strip leading/trailing `[...]` from a workbook id string when the brackets represent an
    /// Excel-style workbook prefix.
    ///
    /// This uses Excel escaping rules where literal `]` characters are escaped by doubling them
    /// (`]]`). Workbook ids may contain literal `[` characters, which do **not** introduce nesting.
    ///
    /// Returns `None` when `workbook` does not appear to be wrapped.
    fn strip_wrapping_workbook_brackets(workbook: &str) -> Option<&str> {
        let bytes = workbook.as_bytes();
        if bytes.first() != Some(&b'[') {
            return None;
        }

        // Find the closing `]` for the opening bracket using Excel workbook escaping rules.
        let mut i = 1usize;
        while i < bytes.len() {
            if bytes[i] == b']' {
                if bytes.get(i + 1) == Some(&b']') {
                    i += 2;
                    continue;
                }

                // Only treat this as a wrapper when it spans the entire string.
                if i + 1 == bytes.len() {
                    return Some(&workbook[1..i]);
                }
                return None;
            }

            // Advance by UTF-8 char boundaries so we don't accidentally interpret `[` / `]` bytes
            // inside multi-byte sequences as actual bracket characters.
            let ch = workbook[i..].chars().next()?;
            i += ch.len_utf8();
        }

        None
    }

    pub fn bytecode_program_count(&self) -> usize {
        self.bytecode_cache.program_count()
    }

    /// Returns aggregate bytecode compilation coverage statistics for the current workbook.
    ///
    /// This is an introspection-only API; it does not affect evaluation behavior.
    pub fn bytecode_compile_stats(&self) -> BytecodeCompileStats {
        let mut stats = BytecodeCompileStats::default();

        for sheet in &self.workbook.sheets {
            for cell in sheet.cells.values() {
                if cell.formula.is_none() {
                    continue;
                }

                stats.total_formula_cells += 1;
                match cell.compiled.as_ref() {
                    Some(CompiledFormula::Bytecode(_)) => {
                        stats.compiled += 1;
                    }
                    Some(CompiledFormula::Ast(_)) => {
                        stats.fallback += 1;
                        let reason = cell
                            .bytecode_compile_reason
                            .clone()
                            .unwrap_or(BytecodeCompileReason::IneligibleExpr);
                        *stats.fallback_reasons.entry(reason).or_insert(0) += 1;
                    }
                    None => {
                        // Formula cells should always have a compiled representation, but avoid
                        // panicking in this introspection API.
                        stats.fallback += 1;
                        *stats
                            .fallback_reasons
                            .entry(BytecodeCompileReason::IneligibleExpr)
                            .or_insert(0) += 1;
                    }
                }
            }
        }

        stats
    }

    /// Returns a per-cell list of formulas that were not compiled to bytecode.
    ///
    /// The results are deterministically ordered by `(tab_index, row, col)` (where `tab_index`
    /// follows the workbook's sheet tab order) and truncated to `limit`.
    pub fn bytecode_compile_report(&self, limit: usize) -> Vec<BytecodeCompileReportEntry> {
        if limit == 0 {
            return Vec::new();
        }

        let mut entries: Vec<(usize, SheetId, CellAddr, BytecodeCompileReason)> = Vec::new();

        for (tab_index, &sheet_id) in self.workbook.sheet_ids_in_order().iter().enumerate() {
            let Some(sheet) = self.workbook.sheets.get(sheet_id) else {
                continue;
            };
            for (addr, cell) in &sheet.cells {
                if cell.formula.is_none() {
                    continue;
                }
                if matches!(cell.compiled.as_ref(), Some(CompiledFormula::Bytecode(_))) {
                    continue;
                }
                let reason = cell
                    .bytecode_compile_reason
                    .clone()
                    .unwrap_or(BytecodeCompileReason::IneligibleExpr);
                entries.push((tab_index, sheet_id, *addr, reason));
            }
        }

        entries.sort_by_key(|(tab_index, _, addr, _)| (*tab_index, addr.row, addr.col));

        let mut out = Vec::new();
        for (_, sheet_id, addr, reason) in entries.into_iter().take(limit) {
            let sheet = self
                .workbook
                .sheet_name(sheet_id)
                .unwrap_or_default()
                .to_string();
            out.push(BytecodeCompileReportEntry {
                sheet,
                addr,
                reason,
            });
        }
        out
    }

    pub fn set_bytecode_enabled(&mut self, enabled: bool) {
        if self.bytecode_enabled == enabled {
            return;
        }
        self.bytecode_enabled = enabled;

        // Rebuild compiled formula variants to match the new bytecode setting. This ensures tests
        // (and callers) can force AST vs bytecode evaluation deterministically.
        let mut updates: Vec<(CellKey, CompiledFormula, Option<BytecodeCompileReason>)> =
            Vec::new();
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            if !self.workbook.sheet_exists(sheet_id) {
                continue;
            }
            for (addr, cell) in &sheet.cells {
                let Some(formula) = cell.formula.as_deref() else {
                    continue;
                };
                let Some(compiled) = cell.compiled.as_ref() else {
                    continue;
                };

                let key = CellKey {
                    sheet: sheet_id,
                    addr: *addr,
                };
                let ast = compiled.ast().clone();

                if !enabled {
                    updates.push((
                        key,
                        CompiledFormula::Ast(ast),
                        Some(BytecodeCompileReason::Disabled),
                    ));
                    continue;
                }

                let origin = crate::CellAddr::new(addr.row, addr.col);
                let parsed = match crate::parse_formula(
                    formula,
                    crate::ParseOptions {
                        locale: crate::LocaleConfig::en_us(),
                        reference_style: crate::ReferenceStyle::A1,
                        normalize_relative_to: Some(origin),
                    },
                ) {
                    Ok(parsed) => parsed,
                    Err(_) => {
                        updates.push((
                            key,
                            CompiledFormula::Ast(ast),
                            Some(BytecodeCompileReason::IneligibleExpr),
                        ));
                        continue;
                    }
                };

                let (compiled_formula, bytecode_compile_reason) = match self.try_compile_bytecode(
                    &parsed.expr,
                    key,
                    cell.thread_safe,
                    cell.dynamic_deps,
                ) {
                    Ok(program) => (
                        CompiledFormula::Bytecode(BytecodeFormula {
                            ast,
                            program,
                            sheet_dims_generation: self.sheet_dims_generation,
                        }),
                        None,
                    ),
                    Err(reason) => (CompiledFormula::Ast(ast), Some(reason)),
                };
                updates.push((key, compiled_formula, bytecode_compile_reason));
            }
        }

        for (key, compiled, bytecode_compile_reason) in updates {
            if let Some(cell) = self
                .workbook
                .sheets
                .get_mut(key.sheet)
                .and_then(|sheet| sheet.cells.get_mut(&key.addr))
            {
                cell.compiled = Some(compiled);
                cell.bytecode_compile_reason = bytecode_compile_reason;
            }
        }

        self.mark_all_compiled_cells_dirty();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    pub fn set_date_system(&mut self, system: ExcelDateSystem) {
        if self.date_system == system {
            return;
        }
        self.date_system = system;
        self.mark_all_compiled_cells_dirty();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    pub fn date_system(&self) -> ExcelDateSystem {
        self.date_system
    }

    pub fn set_value_locale(&mut self, value_locale: ValueLocaleConfig) {
        if self.value_locale == value_locale {
            return;
        }
        self.value_locale = value_locale;
        self.mark_all_compiled_cells_dirty();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    pub fn set_value_locale_id(&mut self, locale_id: &str) -> bool {
        let Some(config) = ValueLocaleConfig::for_locale_id(locale_id) else {
            return false;
        };
        self.set_value_locale(config);
        true
    }

    pub fn value_locale(&self) -> ValueLocaleConfig {
        self.value_locale
    }

    /// Workbook text codepage (Windows code page number).
    ///
    /// This is used for legacy DBCS behaviors like `ASC` / `DBCS`.
    pub fn text_codepage(&self) -> u16 {
        self.text_codepage
    }

    pub fn set_text_codepage(&mut self, text_codepage: u16) {
        if self.text_codepage == text_codepage {
            return;
        }
        self.text_codepage = text_codepage;
        self.workbook.text_codepage = text_codepage;
        self.mark_all_compiled_cells_dirty();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    fn mark_all_compiled_cells_dirty(&mut self) {
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            if !self.workbook.sheet_exists(sheet_id) {
                continue;
            }
            for (addr, cell) in &sheet.cells {
                if cell.compiled.is_some() {
                    let key = CellKey {
                        sheet: sheet_id,
                        addr: *addr,
                    };
                    self.dirty.insert(key);
                    self.dirty_reasons.remove(&key);
                    self.calc_graph.mark_dirty(cell_id_from_key(key));
                }
            }
        }

        self.sync_dirty_from_calc_graph();
    }

    pub fn has_dirty_cells(&self) -> bool {
        !self.dirty.is_empty()
    }

    pub fn circular_reference_count(&self) -> usize {
        self.circular_references.len()
    }

    /// Set the number format pattern for a cell (e.g. `"0.00"`, `"0%"`).
    ///
    /// When `None` (or an empty/whitespace string) is provided, the cell behaves like Excel's
    /// `"General"` format.
    pub fn set_cell_number_format(
        &mut self,
        sheet: &str,
        addr: &str,
        format_pattern: Option<String>,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let addr = parse_a1(addr)?;

        let format_pattern =
            format_pattern.and_then(|s| if s.trim().is_empty() { None } else { Some(s) });

        // Keep the same safety bounds as `set_cell_value`/formula compilation paths.
        if addr.row >= i32::MAX as u32 {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::RowOutOfRange,
            ));
        }

        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let existing = self
            .workbook
            .get_cell(key)
            .and_then(|cell| cell.number_format.as_deref());
        if existing == format_pattern.as_deref() {
            return Ok(());
        }

        let sheet_dims_changed =
            format_pattern.is_some() && self.workbook.grow_sheet_dimensions(sheet_id, addr);
        if sheet_dims_changed {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
            self.mark_all_compiled_cells_dirty();
        }

        match format_pattern {
            Some(pattern) => {
                let cell = self.workbook.get_or_create_cell_mut(key);
                cell.number_format = Some(pattern);
            }
            None => {
                if let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) {
                    if let Some(cell) = sheet_state.cells.get_mut(&addr) {
                        cell.number_format = None;

                        // Preserve sparse semantics when clearing the explicit override:
                        // blank + no formula + default style => remove the cell entry entirely.
                        let remove_cell = cell.value == Value::Blank
                            && cell.formula.is_none()
                            && cell.style_id == 0
                            && cell.phonetic.is_none()
                            && cell.number_format.is_none();
                        if remove_cell {
                            sheet_state.cells.remove(&addr);
                        }
                    }
                }
            }
        }

        // In "precision as displayed" mode, cell number format changes can affect stored numeric
        // values at formula boundaries, even when the formula itself is unchanged. Mark the cell
        // dirty so it is re-evaluated on the next recalculation tick.
        if !self.calc_settings.full_precision {
            let has_formula = self
                .workbook
                .get_cell(key)
                .is_some_and(|cell| cell.formula.is_some());
            if has_formula {
                let cell_id = cell_id_from_key(key);
                self.mark_dirty_including_self_with_reasons(key);
                self.calc_graph.mark_dirty(cell_id);
                self.sync_dirty_from_calc_graph();
            }
        }

        // Number format changes affect volatile metadata functions (e.g. `CELL("format")`).
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    /// Get a cell's number format pattern.
    ///
    /// Returns `Ok(None)` when the cell has no explicit number format (Excel `"General"`) or does
    /// not exist.
    pub fn cell_number_format(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Option<String>, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(None);
        };
        let addr = parse_a1(addr)?;
        if let Some(sheet_state) = self.workbook.sheets.get(sheet_id) {
            if addr.row >= sheet_state.row_count || addr.col >= sheet_state.col_count {
                return Ok(None);
            }
        }
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        Ok(self
            .workbook
            .get_cell(key)
            .and_then(|cell| cell.number_format.clone()))
    }

    pub fn set_cell_value(
        &mut self,
        sheet: &str,
        addr: &str,
        value: impl Into<Value>,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let addr = parse_a1(addr)?;
        // The engine supports rows beyond Excel's default 1,048,576 limit, but some internal
        // evaluation paths (notably the bytecode engine and reference rewriting) use 32-bit
        // coordinates. Keep sheet growth bounded to `i32::MAX` rows so all row/offset conversions
        // remain sound.
        if addr.row >= i32::MAX as u32 {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::RowOutOfRange,
            ));
        }
        if self.workbook.grow_sheet_dimensions(sheet_id, addr) {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
            // Sheet dimensions affect out-of-bounds `#REF!` semantics for references. If the sheet
            // grows, formulas that previously evaluated to `#REF!` may now become valid (and vice
            // versa for any future shrinking API), so conservatively mark all compiled formulas
            // dirty to ensure results refresh on the next recalculation.
            self.mark_all_compiled_cells_dirty();
        }
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let cell_id = cell_id_from_key(key);

        let format_pattern = self.number_format_pattern_for_rounding(key);
        let value: Value = value.into();
        let value = match value {
            Value::Number(n) => Value::Number(self.round_number_as_displayed(n, format_pattern)),
            other => other,
        };

        self.clear_spill_for_cell(key);
        self.clear_blocked_spill_for_origin(key);

        // Replace any existing formula and dependencies.
        self.calc_graph.remove_cell(cell_id);
        self.clear_cell_name_refs(key);
        self.clear_cell_external_refs(key);
        self.clear_cell_dynamic_external_precedents(key);
        self.dirty.remove(&key);
        self.dirty_reasons.remove(&key);

        let remove_cell = {
            let cell = self.workbook.get_or_create_cell_mut(key);
            cell.value = value;
            cell.phonetic = None;
            cell.formula = None;
            cell.compiled = None;
            cell.bytecode_compile_reason = None;
            cell.volatile = false;
            cell.thread_safe = true;
            cell.dynamic_deps = false;

            // Preserve sparse semantics when clearing contents:
            // - blank + no formula + default style => remove the cell entry entirely
            // - otherwise keep the entry (e.g. style-only cell)
            cell.value == Value::Blank
                && cell.formula.is_none()
                && cell.style_id == 0
                && cell.phonetic.is_none()
                && cell.number_format.is_none()
        };
        if let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) {
            // Clearing a cell value removes any formula metadata, so drop it from the
            // `INFO("origin")` dependency index as well.
            sheet.origin_dependents.remove(&addr);
            if remove_cell {
                sheet.cells.remove(&addr);
            }
        }

        // Mark downstream dependents dirty.
        self.mark_dirty_dependents_with_reasons(key);
        self.calc_graph.mark_dirty(cell_id);
        self.mark_dirty_blocked_spill_origins_for_cell(key);
        self.sync_dirty_from_calc_graph();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    pub fn get_cell_style_id(&self, sheet: &str, addr: &str) -> Result<Option<u32>, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(None);
        };
        let addr = parse_a1(addr)?;
        Ok(self
            .workbook
            .sheets
            .get(sheet_id)
            .and_then(|sheet| sheet.cells.get(&addr))
            .map(|cell| cell.style_id))
    }
    /// Sets the phonetic guide string (furigana) metadata for a single cell.
    ///
    /// This metadata is consumed by Excel's `PHONETIC(...)` worksheet function. When unset,
    /// `PHONETIC(reference)` should fall back to the referenced cell's displayed text.
    pub fn set_cell_phonetic(
        &mut self,
        sheet: &str,
        addr: &str,
        phonetic: Option<String>,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let addr = parse_a1(addr)?;
        // Keep coordinates bounded so internal 32-bit conversions remain sound.
        if addr.row >= i32::MAX as u32 {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::RowOutOfRange,
            ));
        }
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let cell_id = cell_id_from_key(key);

        let existing = self
            .workbook
            .get_cell(key)
            .and_then(|cell| cell.phonetic.as_deref());
        if existing == phonetic.as_deref() {
            return Ok(());
        }

        if phonetic.is_some() && self.workbook.grow_sheet_dimensions(sheet_id, addr) {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
            self.mark_all_compiled_cells_dirty();
        }

        match phonetic {
            None => {
                let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) else {
                    return Ok(());
                };
                let remove_cell = {
                    let Some(cell) = sheet.cells.get_mut(&addr) else {
                        return Ok(());
                    };
                    cell.phonetic = None;
                    cell.value == Value::Blank
                        && cell.formula.is_none()
                        && cell.style_id == 0
                        && cell.phonetic.is_none()
                        && cell.number_format.is_none()
                };
                if remove_cell {
                    sheet.cells.remove(&addr);
                }
            }
            Some(phonetic) => {
                let cell = self.workbook.get_or_create_cell_mut(key);
                cell.phonetic = Some(phonetic);
            }
        }

        self.mark_dirty_dependents_with_reasons(key);
        self.calc_graph.mark_dirty(cell_id);
        self.sync_dirty_from_calc_graph();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    /// Set a rectangular range of literal values.
    ///
    /// This is a bulk variant of [`Engine::set_cell_value`]. It applies all values in the range
    /// while deferring recalculation until the end (at most once).
    ///
    /// - `values` must be a matrix with dimensions matching `range.height()` x `range.width()`.
    /// - When `recalc` is `true` and the workbook is in an automatic calculation mode, the engine
    ///   recalculates once after the entire range has been applied.
    pub fn set_range_values(
        &mut self,
        sheet: &str,
        range: Range,
        values: &[Vec<Value>],
        recalc: bool,
    ) -> Result<(), EngineError> {
        let expected_rows = range.height() as usize;
        let expected_cols = range.width() as usize;

        if values.len() != expected_rows {
            let actual_cols = values.get(0).map(|row| row.len()).unwrap_or(0);
            return Err(EngineError::RangeValuesDimensionMismatch {
                expected_rows,
                expected_cols,
                actual_rows: values.len(),
                actual_cols,
            });
        }
        for row in values {
            if row.len() != expected_cols {
                return Err(EngineError::RangeValuesDimensionMismatch {
                    expected_rows,
                    expected_cols,
                    actual_rows: values.len(),
                    actual_cols: row.len(),
                });
            }
        }

        let sheet_id = self.workbook.ensure_sheet(sheet);

        // Enforce Excel's fixed 16,384-column limit and the engine's i32 row bound.
        if range.end.row >= i32::MAX as u32 {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::RowOutOfRange,
            ));
        }
        if range.end.col >= EXCEL_MAX_COLS {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::ColumnOutOfRange,
            ));
        }

        if self
            .workbook
            .grow_sheet_dimensions(sheet_id, cell_addr_from_cell_ref(range.end))
        {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
            // Sheet dimensions affect out-of-bounds `#REF!` semantics for references. If the sheet
            // grows, formulas that previously evaluated to `#REF!` may now become valid, so
            // conservatively mark all compiled formulas dirty to ensure results refresh on the next
            // recalculation.
            self.mark_all_compiled_cells_dirty();
        }

        let start_row = range.start.row;
        let start_col = range.start.col;

        for (r_off, row_values) in values.iter().enumerate() {
            let row = start_row + r_off as u32;
            for (c_off, raw_value) in row_values.iter().enumerate() {
                let col = start_col + c_off as u32;
                let addr = CellAddr { row, col };
                let key = CellKey {
                    sheet: sheet_id,
                    addr,
                };

                // Match `set_cell_value` semantics, including "precision as displayed" rounding for
                // numeric literals.
                let value = if self.calc_settings.full_precision {
                    raw_value.clone()
                } else {
                    match raw_value {
                        Value::Number(n) => {
                            let format_pattern = self.number_format_pattern_for_rounding(key);
                            Value::Number(self.round_number_as_displayed(*n, format_pattern))
                        }
                        other => other.clone(),
                    }
                };

                // Treat `Value::Blank` as clearing cell contents. Preserve formatting for style-only
                // cells, but keep storage sparse by pruning empty default cells.
                let is_blank = value == Value::Blank;

                let in_spill = self.spill_origin_key(key).is_some();
                let needs_update = if in_spill {
                    true
                } else if let Some(cell) = self.workbook.get_cell(key) {
                    // Mirror `set_cell_value`: even if the scalar value is unchanged, setting a
                    // value should clear any formula/bytecode state and reset per-cell evaluation
                    // metadata.
                    //
                    // Also keep sparse semantics: clearing a default-style blank cell should remove
                    // its storage entry.
                    let should_prune_default_blank_cell = is_blank
                        && cell.value == Value::Blank
                        && cell.formula.is_none()
                        && cell.style_id == 0
                        && cell.phonetic.is_none()
                        && cell.number_format.is_none();

                    cell.formula.is_some()
                        || cell.compiled.is_some()
                        || cell.bytecode_compile_reason.is_some()
                        || cell.volatile
                        || !cell.thread_safe
                        || cell.dynamic_deps
                        || cell.phonetic.is_some()
                        || cell.value != value
                        || should_prune_default_blank_cell
                } else {
                    !is_blank
                };

                if !needs_update {
                    continue;
                }

                let cell_id = cell_id_from_key(key);

                self.clear_spill_for_cell(key);
                self.clear_blocked_spill_for_origin(key);

                // Replace any existing formula and dependencies.
                self.calc_graph.remove_cell(cell_id);
                self.clear_cell_name_refs(key);
                self.clear_cell_external_refs(key);
                self.clear_cell_dynamic_external_precedents(key);
                self.dirty.remove(&key);
                self.dirty_reasons.remove(&key);

                if is_blank {
                    let remove_cell = {
                        let cell = self.workbook.get_or_create_cell_mut(key);
                        cell.value = Value::Blank;
                        cell.phonetic = None;
                        cell.formula = None;
                        cell.compiled = None;
                        cell.bytecode_compile_reason = None;
                        cell.volatile = false;
                        cell.thread_safe = true;
                        cell.dynamic_deps = false;

                        // Preserve sparse semantics when clearing contents:
                        // - blank + no formula + default style => remove the cell entry entirely
                        // - otherwise keep the entry (e.g. style-only cell)
                        cell.value == Value::Blank
                            && cell.formula.is_none()
                            && cell.style_id == 0
                            && cell.phonetic.is_none()
                            && cell.number_format.is_none()
                    };
                    if let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) {
                        sheet_state.origin_dependents.remove(&addr);
                        if remove_cell {
                            sheet_state.cells.remove(&addr);
                        }
                    }
                } else {
                    {
                        let cell = self.workbook.get_or_create_cell_mut(key);
                        cell.value = value;
                        cell.phonetic = None;
                        cell.formula = None;
                        cell.compiled = None;
                        cell.bytecode_compile_reason = None;
                        cell.volatile = false;
                        cell.thread_safe = true;
                        cell.dynamic_deps = false;
                    }
                    if let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) {
                        sheet_state.origin_dependents.remove(&addr);
                    }
                };

                // Mark downstream dependents dirty.
                self.mark_dirty_dependents_with_reasons(key);
                self.calc_graph.mark_dirty(cell_id);
                self.mark_dirty_blocked_spill_origins_for_cell(key);
            }
        }

        self.sync_dirty_from_calc_graph();
        if recalc && self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    /// Clears a rectangular range of cells, removing them from the workbook's sparse storage.
    ///
    /// This is a bulk variant of [`Engine::clear_cell`]. It clears all cells in the range while
    /// deferring recalculation until the end (at most once).
    pub fn clear_range(
        &mut self,
        sheet: &str,
        range: Range,
        recalc: bool,
    ) -> Result<(), EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(());
        };

        if let Some(sheet_state) = self.workbook.sheets.get(sheet_id) {
            if range.end.row >= sheet_state.row_count {
                return Err(EngineError::Address(
                    crate::eval::AddressParseError::RowOutOfRange,
                ));
            }
            if range.end.col >= sheet_state.col_count {
                return Err(EngineError::Address(
                    crate::eval::AddressParseError::ColumnOutOfRange,
                ));
            }
        }

        let start_row = range.start.row;
        let end_row = range.end.row;
        let start_col = range.start.col;
        let end_col = range.end.col;

        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let addr = CellAddr { row, col };
                let key = CellKey {
                    sheet: sheet_id,
                    addr,
                };

                let in_spill = self.spill_origin_key(key).is_some();
                let has_cell = self
                    .workbook
                    .sheets
                    .get(sheet_id)
                    .and_then(|s| s.cells.get(&addr))
                    .is_some();
                if !in_spill && !has_cell {
                    continue;
                }

                let cell_id = cell_id_from_key(key);

                self.clear_spill_for_cell(key);
                self.clear_blocked_spill_for_origin(key);

                self.calc_graph.remove_cell(cell_id);
                self.clear_cell_name_refs(key);
                self.clear_cell_external_refs(key);
                self.clear_cell_dynamic_external_precedents(key);
                self.dirty.remove(&key);
                self.dirty_reasons.remove(&key);

                if let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) {
                    sheet_state.cells.remove(&addr);
                    sheet_state.origin_dependents.remove(&addr);
                }

                self.mark_dirty_dependents_with_reasons(key);
                self.calc_graph.mark_dirty(cell_id);
                self.mark_dirty_blocked_spill_origins_for_cell(key);
            }
        }

        self.sync_dirty_from_calc_graph();
        if recalc && self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }
    /// Clears a cell's stored value/formula *and formatting* so it behaves as if it does not exist.
    ///
    /// This is distinct from setting a cell to [`Value::Blank`], which behaves like Excel "clear
    /// contents" and preserves a cell's formatting (style id) when present.
    pub fn clear_cell(&mut self, sheet: &str, addr: &str) -> Result<(), EngineError> {
        let addr = parse_a1(addr)?;
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(());
        };
        if let Some(sheet_state) = self.workbook.sheets.get(sheet_id) {
            if addr.row >= sheet_state.row_count {
                return Err(EngineError::Address(
                    crate::eval::AddressParseError::RowOutOfRange,
                ));
            }
            if addr.col >= sheet_state.col_count {
                return Err(EngineError::Address(
                    crate::eval::AddressParseError::ColumnOutOfRange,
                ));
            }
        }
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let cell_id = cell_id_from_key(key);

        self.clear_spill_for_cell(key);
        self.clear_blocked_spill_for_origin(key);

        // Remove any existing formula and dependencies.
        self.calc_graph.remove_cell(cell_id);
        self.clear_cell_name_refs(key);
        self.clear_cell_external_refs(key);
        self.clear_cell_dynamic_external_precedents(key);
        self.dirty.remove(&key);
        self.dirty_reasons.remove(&key);

        if let Some(sheet) = self.workbook.sheets.get_mut(sheet_id) {
            sheet.cells.remove(&addr);
            sheet.origin_dependents.remove(&addr);
        }

        // Mark downstream dependents dirty.
        self.mark_dirty_dependents_with_reasons(key);
        self.calc_graph.mark_dirty(cell_id);
        self.mark_dirty_blocked_spill_origins_for_cell(key);
        self.sync_dirty_from_calc_graph();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    /// Register a pivot table's metadata for use by `GETPIVOTDATA`.
    ///
    /// The pivot must already have been rendered/applied into `destination` on `sheet`.
    /// `GETPIVOTDATA` will resolve any reference within `destination` back to this pivot.
    pub fn register_pivot_table(
        &mut self,
        sheet: &str,
        destination: Range,
        pivot: crate::pivot::PivotTable,
    ) -> Result<(), crate::pivot_registry::PivotRegistryError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let destination = crate::pivot_registry::PivotDestination {
            start: CellAddr {
                row: destination.start.row,
                col: destination.start.col,
            },
            end: CellAddr {
                row: destination.end.row,
                col: destination.end.col,
            },
        };
        let entry = crate::pivot_registry::PivotRegistryEntry::new(sheet_id, destination, pivot)?;
        self.pivot_registry.register(entry);
        Ok(())
    }

    /// Clears all registered pivots.
    pub fn clear_pivot_registry(&mut self) {
        self.pivot_registry.clear();
    }

    /// Returns the currently registered pivot table metadata entries.
    ///
    /// These entries are used by `GETPIVOTDATA` to resolve references within rendered pivot output
    /// grids back to their pivot caches/configuration.
    pub fn pivot_registry_entries(&self) -> &[crate::pivot_registry::PivotRegistryEntry] {
        self.pivot_registry.entries()
    }

    /// Replace the set of tables for a given worksheet.
    ///
    /// Tables are needed to resolve structured references like `Table1[Col]` and `[@Col]`.
    pub fn set_sheet_tables(&mut self, sheet: &str, tables: Vec<Table>) {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        self.workbook.set_tables(sheet_id, tables);

        let tables_by_sheet: Vec<Vec<Table>> = self
            .workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if self.workbook.sheet_exists(sheet_id) {
                    s.tables.clone()
                } else {
                    Vec::new()
                }
            })
            .collect();

        // Structured reference resolution can change which cells a formula depends on, so refresh
        // dependencies for all formulas.
        //
        // Table changes can also change how bytecode compilation lowers structured references
        // (e.g. `Table1[Col]` expands/shrinks as the table grows), so rebuild the compiled
        // variant (AST vs bytecode) for all formula cells.
        let mut formulas: Vec<(CellKey, String, CompiledExpr)> = Vec::new();
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            if !self.workbook.sheet_exists(sheet_id) {
                continue;
            }
            for (addr, cell) in &sheet.cells {
                let Some(formula) = cell.formula.as_deref() else {
                    continue;
                };
                let Some(compiled) = cell.compiled.as_ref() else {
                    continue;
                };
                formulas.push((
                    CellKey {
                        sheet: sheet_id,
                        addr: *addr,
                    },
                    formula.to_string(),
                    compiled.ast().clone(),
                ));
            }
        }

        for (key, formula, ast) in formulas {
            let cell_id = cell_id_from_key(key);
            let (names, volatile, thread_safe, dynamic_deps, origin_deps) = analyze_expr_flags(
                &ast,
                key,
                &tables_by_sheet,
                &self.workbook,
                self.external_refs_volatile,
            );
            self.set_cell_name_refs(key, names);
            let (external_sheets, external_workbooks) = analyze_external_dependencies(
                &ast,
                key,
                &self.workbook,
                self.external_value_provider.as_deref(),
            );
            self.set_cell_external_refs(key, external_sheets, external_workbooks);

            let calc_precedents =
                analyze_calc_precedents(&ast, key, &tables_by_sheet, &self.workbook, &self.spills);
            let mut calc_vec: Vec<Precedent> = calc_precedents.into_iter().collect();
            calc_vec.sort_by_key(|p| match p {
                Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col, 0u32, 0u32),
                Precedent::Range(r) => (
                    1u8,
                    r.sheet_id,
                    r.range.start.row,
                    r.range.start.col,
                    r.range.end.row,
                    r.range.end.col,
                ),
            });
            let deps = CellDeps::new(calc_vec).volatile(volatile);
            self.calc_graph.update_cell_dependencies(cell_id, deps);

            let (compiled_formula, bytecode_compile_reason) = if !self.bytecode_enabled {
                (
                    CompiledFormula::Ast(ast.clone()),
                    Some(BytecodeCompileReason::Disabled),
                )
            } else {
                let origin = crate::CellAddr::new(key.addr.row, key.addr.col);
                let parsed = crate::parse_formula(
                    &formula,
                    crate::ParseOptions {
                        locale: crate::LocaleConfig::en_us(),
                        reference_style: crate::ReferenceStyle::A1,
                        normalize_relative_to: Some(origin),
                    },
                );
                match parsed {
                    Ok(parsed) => match self.try_compile_bytecode(
                        &parsed.expr,
                        key,
                        thread_safe,
                        dynamic_deps,
                    ) {
                        Ok(program) => (
                            CompiledFormula::Bytecode(BytecodeFormula {
                                ast: ast.clone(),
                                program,
                                sheet_dims_generation: self.sheet_dims_generation,
                            }),
                            None,
                        ),
                        Err(reason) => (CompiledFormula::Ast(ast.clone()), Some(reason)),
                    },
                    Err(_) => (
                        CompiledFormula::Ast(ast.clone()),
                        Some(BytecodeCompileReason::IneligibleExpr),
                    ),
                }
            };

            {
                let cell = self.workbook.get_or_create_cell_mut(key);
                cell.compiled = Some(compiled_formula);
                cell.bytecode_compile_reason = bytecode_compile_reason;
                cell.volatile = volatile;
                cell.thread_safe = thread_safe;
                cell.dynamic_deps = dynamic_deps;
            }

            if let Some(sheet_state) = self.workbook.sheets.get_mut(key.sheet) {
                if origin_deps {
                    sheet_state.origin_dependents.insert(key.addr);
                } else {
                    sheet_state.origin_dependents.remove(&key.addr);
                }
            }

            self.dirty.insert(key);
            self.dirty_reasons.remove(&key);
            self.calc_graph.mark_dirty(cell_id);
        }

        // Ensure the engine-level dirty set matches the dependency graph (including any spill
        // output nodes that were marked dirty as dependents).
        self.sync_dirty_from_calc_graph();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
    }

    /// Rename an Excel table (ListObject) and rewrite any impacted formulas.
    ///
    /// This emulates Excel's "Rename Table" behavior:
    /// - The new name is validated using [`formula_model::validate_table_name`].
    /// - Table names are workbook-scoped and must be unique (case-insensitive) across both
    ///   `Table.name` and `Table.display_name`.
    /// - Formulas referencing either the table `name` or `display_name` are rewritten to preserve
    ///   semantics.
    ///
    /// Returns a list of rewritten cell formulas so callers can update UI state.
    pub fn rename_table(
        &mut self,
        old_name: &str,
        new_name: &str,
    ) -> Result<Vec<FormulaRewrite>, TableError> {
        let new_name = new_name.trim();
        validate_table_name(new_name)?;

        let (sheet_idx, table_idx) = self
            .workbook
            .sheets
            .iter()
            .enumerate()
            .find_map(|(si, sheet)| {
                sheet
                    .tables
                    .iter()
                    .position(|t| {
                        t.name.eq_ignore_ascii_case(old_name)
                            || t.display_name.eq_ignore_ascii_case(old_name)
                    })
                    .map(|ti| (si, ti))
            })
            .ok_or(TableError::TableNotFound)?;

        let actual_old_name = self.workbook.sheets[sheet_idx].tables[table_idx]
            .name
            .clone();
        let actual_old_display_name = self.workbook.sheets[sheet_idx].tables[table_idx]
            .display_name
            .clone();

        // Enforce workbook-wide uniqueness (case-insensitive) across both `name` and `display_name`.
        for (si, sheet) in self.workbook.sheets.iter().enumerate() {
            for (ti, table) in sheet.tables.iter().enumerate() {
                if si == sheet_idx && ti == table_idx {
                    continue;
                }
                if table.name.eq_ignore_ascii_case(new_name)
                    || table.display_name.eq_ignore_ascii_case(new_name)
                {
                    return Err(TableError::DuplicateName);
                }
            }
        }

        // Build rename pairs. Excel rewrites references to either `name` or `display_name`.
        let mut renames = vec![(actual_old_name.clone(), new_name.to_string())];
        if !actual_old_display_name.eq_ignore_ascii_case(&actual_old_name) {
            renames.push((actual_old_display_name, new_name.to_string()));
        }

        let sheet_names = sheet_names_by_id(&self.workbook);
        let mut rewrites: Vec<FormulaRewrite> = Vec::new();

        // 1) Rewrite worksheet cell formulas (and update compiled IR to match).
        for (sheet_id, sheet) in self.workbook.sheets.iter_mut().enumerate() {
            let sheet_name = sheet_names
                .get(&sheet_id)
                .cloned()
                .unwrap_or_else(|| sheet_id.to_string());

            for (addr, cell) in sheet.cells.iter_mut() {
                let Some(formula) = cell.formula.as_deref() else {
                    continue;
                };
                let rewritten = rewrite_table_names_in_formula(formula, &renames);
                if rewritten == formula {
                    continue;
                }
                let before = formula.to_string();
                let after = rewritten.clone();
                cell.formula = Some(rewritten.into());

                if let Some(compiled) = cell.compiled.as_mut() {
                    rewrite_table_names_in_compiled_formula(compiled, &renames);
                }

                rewrites.push(FormulaRewrite {
                    sheet: sheet_name.clone(),
                    cell: CellRef::new(addr.row, addr.col),
                    before,
                    after,
                });
            }

            // 2) Rewrite formulas stored in table metadata (best-effort).
            for table in &mut sheet.tables {
                for column in &mut table.columns {
                    if let Some(formula) = column.formula.as_mut() {
                        *formula = rewrite_table_names_in_formula(formula, &renames);
                    }
                    if let Some(formula) = column.totals_formula.as_mut() {
                        *formula = rewrite_table_names_in_formula(formula, &renames);
                    }
                }
            }
        }

        // 3) Rewrite defined-name formulas (workbook + sheet scoped).
        rewrite_table_names_in_defined_names(&mut self.workbook.names, &renames);
        for sheet in &mut self.workbook.sheets {
            rewrite_table_names_in_defined_names(&mut sheet.names, &renames);
        }

        // 4) Rename the actual table.
        let renamed = &mut self.workbook.sheets[sheet_idx].tables[table_idx];
        renamed.name = new_name.to_string();
        renamed.display_name = new_name.to_string();

        // Mark formulas dirty so recalculation picks up the updated table metadata / formula text.
        self.mark_all_compiled_cells_dirty();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }

        Ok(rewrites)
    }

    /// Returns the current set of tables for `sheet`.
    ///
    /// This is primarily intended for inspection/testing (e.g. verifying structured reference
    /// metadata). Callers should treat table definitions as immutable workbook metadata and use
    /// [`Engine::set_sheet_tables`] to replace them.
    pub fn get_sheet_tables(&self, sheet: &str) -> Option<&[Table]> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let sheet = self.workbook.sheets.get(sheet_id)?;
        Some(sheet.tables.as_slice())
    }

    #[allow(dead_code)]
    fn recompile_all_formula_cells(&mut self) -> Result<(), EngineError> {
        let tables_by_sheet: Vec<Vec<Table>> = self
            .workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if self.workbook.sheet_exists(sheet_id) {
                    s.tables.clone()
                } else {
                    Vec::new()
                }
            })
            .collect();

        // Collect formula cells up-front to avoid borrow conflicts while recompiling.
        let mut formulas: Vec<(CellKey, String)> = Vec::new();
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            if !self.workbook.sheet_exists(sheet_id) {
                continue;
            }
            for (addr, cell) in &sheet.cells {
                if let Some(formula) = cell.formula.as_deref() {
                    formulas.push((
                        CellKey {
                            sheet: sheet_id,
                            addr: *addr,
                        },
                        formula.to_string(),
                    ));
                }
            }
        }

        for (key, formula) in formulas {
            let origin = crate::CellAddr::new(key.addr.row, key.addr.col);
            let parsed = crate::parse_formula(
                &formula,
                crate::ParseOptions {
                    locale: crate::LocaleConfig::en_us(),
                    reference_style: crate::ReferenceStyle::A1,
                    normalize_relative_to: Some(origin),
                },
            )?;

            let mut resolve_sheet = |name: &str| self.workbook.sheet_id(name);
            let mut sheet_dims = |sheet_id: usize| {
                self.workbook
                    .sheets
                    .get(sheet_id)
                    .map(|s| (s.row_count, s.col_count))
                    .unwrap_or((EXCEL_MAX_ROWS, EXCEL_MAX_COLS))
            };
            let compiled_ast = compile_canonical_expr(
                &parsed.expr,
                key.sheet,
                key.addr,
                &mut resolve_sheet,
                &mut sheet_dims,
            );

            let (names, volatile, thread_safe, dynamic_deps, origin_deps) = analyze_expr_flags(
                &compiled_ast,
                key,
                &tables_by_sheet,
                &self.workbook,
                self.external_refs_volatile,
            );
            self.set_cell_name_refs(key, names);
            let (external_sheets, external_workbooks) = analyze_external_dependencies(
                &compiled_ast,
                key,
                &self.workbook,
                self.external_value_provider.as_deref(),
            );
            self.set_cell_external_refs(key, external_sheets, external_workbooks);

            let calc_precedents = analyze_calc_precedents(
                &compiled_ast,
                key,
                &tables_by_sheet,
                &self.workbook,
                &self.spills,
            );
            let mut calc_vec: Vec<Precedent> = calc_precedents.into_iter().collect();
            calc_vec.sort_by_key(|p| match p {
                Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col, 0u32, 0u32),
                Precedent::Range(r) => (
                    1u8,
                    r.sheet_id,
                    r.range.start.row,
                    r.range.start.col,
                    r.range.end.row,
                    r.range.end.col,
                ),
            });
            let cell_id = cell_id_from_key(key);
            let deps = CellDeps::new(calc_vec).volatile(volatile);
            self.calc_graph.update_cell_dependencies(cell_id, deps);

            let (compiled_formula, bytecode_compile_reason) =
                match self.try_compile_bytecode(&parsed.expr, key, thread_safe, dynamic_deps) {
                    Ok(program) => (
                        CompiledFormula::Bytecode(BytecodeFormula {
                            ast: compiled_ast.clone(),
                            program,
                            sheet_dims_generation: self.sheet_dims_generation,
                        }),
                        None,
                    ),
                    Err(reason) => (CompiledFormula::Ast(compiled_ast), Some(reason)),
                };

            if let Some(cell) = self
                .workbook
                .sheets
                .get_mut(key.sheet)
                .and_then(|s| s.cells.get_mut(&key.addr))
            {
                cell.compiled = Some(compiled_formula);
                cell.bytecode_compile_reason = bytecode_compile_reason;
                cell.volatile = volatile;
                cell.thread_safe = thread_safe;
                cell.dynamic_deps = dynamic_deps;
            }

            if let Some(sheet_state) = self.workbook.sheets.get_mut(key.sheet) {
                if origin_deps {
                    sheet_state.origin_dependents.insert(key.addr);
                } else {
                    sheet_state.origin_dependents.remove(&key.addr);
                }
            }

            // Mark the formula (and its transitive dependents) dirty so recalculation picks up any
            // semantic changes (e.g. whole-column range expansion).
            self.dirty.insert(key);
            self.dirty_reasons.remove(&key);
            self.calc_graph.mark_dirty(cell_id);
        }

        self.sync_dirty_from_calc_graph();

        Ok(())
    }

    fn recompile_all_defined_names(&mut self) -> Result<(), EngineError> {
        let mut updates: Vec<(Option<SheetId>, String, Option<CompiledExpr>)> = Vec::new();

        for (name, def) in &self.workbook.names {
            let compiled = match &def.definition {
                NameDefinition::Constant(_) => None,
                NameDefinition::Reference(formula) | NameDefinition::Formula(formula) => {
                    let ast = crate::parse_formula(
                        formula,
                        crate::ParseOptions {
                            locale: crate::LocaleConfig::en_us(),
                            reference_style: crate::ReferenceStyle::A1,
                            normalize_relative_to: None,
                        },
                    )?;
                    let parsed = lower_ast(&ast, None);
                    Some(self.compile_name_expr(&parsed))
                }
            };
            updates.push((None, name.clone(), compiled));
        }

        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            for (name, def) in &sheet.names {
                let compiled = match &def.definition {
                    NameDefinition::Constant(_) => None,
                    NameDefinition::Reference(formula) | NameDefinition::Formula(formula) => {
                        let ast = crate::parse_formula(
                            formula,
                            crate::ParseOptions {
                                locale: crate::LocaleConfig::en_us(),
                                reference_style: crate::ReferenceStyle::A1,
                                normalize_relative_to: None,
                            },
                        )?;
                        let parsed = lower_ast(&ast, None);
                        Some(self.compile_name_expr(&parsed))
                    }
                };
                updates.push((Some(sheet_id), name.clone(), compiled));
            }
        }

        for (scope_sheet, name, compiled) in updates {
            match scope_sheet {
                None => {
                    if let Some(def) = self.workbook.names.get_mut(&name) {
                        def.compiled = compiled;
                    }
                }
                Some(sheet_id) => {
                    if let Some(def) = self
                        .workbook
                        .sheets
                        .get_mut(sheet_id)
                        .and_then(|s| s.names.get_mut(&name))
                    {
                        def.compiled = compiled;
                    }
                }
            }
        }

        Ok(())
    }

    pub fn define_name(
        &mut self,
        name: &str,
        scope: NameScope<'_>,
        definition: NameDefinition,
    ) -> Result<(), EngineError> {
        let name_key = normalize_defined_name(name);
        if name_key.is_empty() {
            return Ok(());
        }

        let compiled = match &definition {
            NameDefinition::Constant(_) => None,
            NameDefinition::Reference(formula) | NameDefinition::Formula(formula) => {
                let ast = crate::parse_formula(
                    formula,
                    crate::ParseOptions {
                        locale: crate::LocaleConfig::en_us(),
                        reference_style: crate::ReferenceStyle::A1,
                        normalize_relative_to: None,
                    },
                )?;
                let parsed = lower_ast(&ast, None);
                Some(self.compile_name_expr(&parsed))
            }
        };

        let entry = DefinedName {
            definition,
            compiled,
        };

        match scope {
            NameScope::Workbook => {
                self.workbook.names.insert(name_key.clone(), entry);
            }
            NameScope::Sheet(sheet_name) => {
                let sheet_id = self.workbook.ensure_sheet(sheet_name);
                self.workbook.sheets[sheet_id]
                    .names
                    .insert(name_key.clone(), entry);
            }
        }

        self.refresh_cells_after_name_change(&name_key);
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    pub fn remove_name(&mut self, name: &str, scope: NameScope<'_>) -> Option<NameDefinition> {
        let name_key = normalize_defined_name(name);
        if name_key.is_empty() {
            return None;
        }

        let removed = match scope {
            NameScope::Workbook => self.workbook.names.remove(&name_key),
            NameScope::Sheet(sheet_name) => {
                let sheet_id = self.workbook.sheet_id(sheet_name)?;
                self.workbook
                    .sheets
                    .get_mut(sheet_id)?
                    .names
                    .remove(&name_key)
            }
        };

        let removed_def = removed.map(|n| n.definition);
        if removed_def.is_some() {
            self.refresh_cells_after_name_change(&name_key);
            if self.calc_settings.calculation_mode != CalculationMode::Manual {
                self.recalculate();
            }
        }
        removed_def
    }

    pub fn get_name(&self, name: &str, scope: NameScope<'_>) -> Option<&NameDefinition> {
        let name_key = normalize_defined_name(name);
        if name_key.is_empty() {
            return None;
        }

        match scope {
            NameScope::Workbook => self.workbook.names.get(&name_key).map(|n| &n.definition),
            NameScope::Sheet(sheet_name) => {
                let sheet_id = self.workbook.sheet_id(sheet_name)?;
                self.workbook
                    .sheets
                    .get(sheet_id)?
                    .names
                    .get(&name_key)
                    .map(|n| &n.definition)
            }
        }
    }

    pub fn set_cell_formula(
        &mut self,
        sheet: &str,
        addr: &str,
        formula: &str,
    ) -> Result<(), EngineError> {
        let sheet_id = self.workbook.ensure_sheet(sheet);
        let addr = parse_a1(addr)?;
        if addr.row >= i32::MAX as u32 {
            return Err(EngineError::Address(
                crate::eval::AddressParseError::RowOutOfRange,
            ));
        }
        if self.workbook.grow_sheet_dimensions(sheet_id, addr) {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
            self.mark_all_compiled_cells_dirty();
        }
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let cell_id = cell_id_from_key(key);
        self.clear_spill_for_cell(key);
        self.clear_blocked_spill_for_origin(key);
        self.clear_cell_dynamic_external_precedents(key);

        let origin = crate::CellAddr::new(addr.row, addr.col);
        let parsed = crate::parse_formula(
            formula,
            crate::ParseOptions {
                locale: crate::LocaleConfig::en_us(),
                reference_style: crate::ReferenceStyle::A1,
                normalize_relative_to: Some(origin),
            },
        )?;
        let mut resolve_sheet = |name: &str| self.workbook.sheet_id(name);
        let mut sheet_dims = |sheet_id: usize| {
            self.workbook
                .sheets
                .get(sheet_id)
                .map(|s| (s.row_count, s.col_count))
                .unwrap_or((EXCEL_MAX_ROWS, EXCEL_MAX_COLS))
        };
        let compiled = compile_canonical_expr(
            &parsed.expr,
            sheet_id,
            addr,
            &mut resolve_sheet,
            &mut sheet_dims,
        );
        let tables_by_sheet: Vec<Vec<Table>> = self
            .workbook
            .sheets
            .iter()
            .map(|s| s.tables.clone())
            .collect();
        let (names, volatile, thread_safe, dynamic_deps, origin_deps) = analyze_expr_flags(
            &compiled,
            key,
            &tables_by_sheet,
            &self.workbook,
            self.external_refs_volatile,
        );
        self.set_cell_name_refs(key, names);
        let (external_sheets, external_workbooks) = analyze_external_dependencies(
            &compiled,
            key,
            &self.workbook,
            self.external_value_provider.as_deref(),
        );
        self.set_cell_external_refs(key, external_sheets, external_workbooks);

        // Optimized precedents for calculation ordering (range nodes are not expanded).
        let calc_precedents = analyze_calc_precedents(
            &compiled,
            key,
            &tables_by_sheet,
            &self.workbook,
            &self.spills,
        );
        let mut calc_vec: Vec<Precedent> = calc_precedents.into_iter().collect();
        calc_vec.sort_by_key(|p| match p {
            Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col, 0u32, 0u32),
            Precedent::Range(r) => (
                1u8,
                r.sheet_id,
                r.range.start.row,
                r.range.start.col,
                r.range.end.row,
                r.range.end.col,
            ),
        });
        let deps = CellDeps::new(calc_vec).volatile(volatile);
        self.calc_graph.update_cell_dependencies(cell_id, deps);

        let (compiled_formula, bytecode_compile_reason) =
            match self.try_compile_bytecode(&parsed.expr, key, thread_safe, dynamic_deps) {
                Ok(program) => (
                    CompiledFormula::Bytecode(BytecodeFormula {
                        ast: compiled.clone(),
                        program,
                        sheet_dims_generation: self.sheet_dims_generation,
                    }),
                    None,
                ),
                Err(reason) => (CompiledFormula::Ast(compiled), Some(reason)),
            };

        let cell = self.workbook.get_or_create_cell_mut(key);
        cell.phonetic = None;
        cell.formula = Some(Arc::from(formula));
        cell.compiled = Some(compiled_formula);
        cell.bytecode_compile_reason = bytecode_compile_reason;
        cell.volatile = volatile;
        cell.thread_safe = thread_safe;
        cell.dynamic_deps = dynamic_deps;

        if let Some(sheet_state) = self.workbook.sheets.get_mut(sheet_id) {
            if origin_deps {
                sheet_state.origin_dependents.insert(addr);
            } else {
                sheet_state.origin_dependents.remove(&addr);
            }
        }

        // Recalculate this cell and anything depending on it.
        self.mark_dirty_including_self_with_reasons(key);
        self.calc_graph.mark_dirty(cell_id);
        self.mark_dirty_blocked_spill_origins_for_cell(key);
        self.sync_dirty_from_calc_graph();
        if self.calc_settings.calculation_mode != CalculationMode::Manual {
            self.recalculate();
        }
        Ok(())
    }

    /// Set a cell formula that was entered using a different reference style (e.g. R1C1).
    ///
    /// The engine persists formulas in canonical A1 form, so this parses the input formula with
    /// the provided [`crate::ParseOptions`] and then serializes it back to A1 based on the cell
    /// location (`addr`).
    pub fn set_cell_formula_with_options(
        &mut self,
        sheet: &str,
        addr: &str,
        formula: &str,
        mut opts: crate::ParseOptions,
    ) -> Result<(), EngineError> {
        let origin_eval = parse_a1(addr)?;
        let origin = crate::CellAddr::new(origin_eval.row, origin_eval.col);

        // Normalize any relative A1 coordinates against the destination cell so the AST is
        // origin-relative regardless of the input reference style.
        opts.normalize_relative_to = Some(origin);

        let ast = crate::parse_formula(formula, opts)?;

        let canonical = ast.to_string(crate::SerializeOptions {
            locale: crate::LocaleConfig::en_us(),
            reference_style: crate::ReferenceStyle::A1,
            // Preserve `_xlfn.` prefixes for round-trip safety when callers include them.
            include_xlfn_prefix: true,
            origin: Some(origin),
            omit_equals: false,
        })?;

        self.set_cell_formula(sheet, addr, &canonical)
    }

    /// Convenience wrapper around [`Engine::set_cell_formula_with_options`] for R1C1 formulas.
    pub fn set_cell_formula_r1c1(
        &mut self,
        sheet: &str,
        addr: &str,
        formula_r1c1: &str,
    ) -> Result<(), EngineError> {
        self.set_cell_formula_with_options(
            sheet,
            addr,
            formula_r1c1,
            crate::ParseOptions {
                locale: crate::LocaleConfig::en_us(),
                reference_style: crate::ReferenceStyle::R1C1,
                normalize_relative_to: None,
            },
        )
    }

    /// Set a cell formula that was entered in a locale-specific display format.
    ///
    /// This converts the incoming formula to canonical form before parsing and
    /// persistence. Canonical form uses English function names and `,`/`.` for
    /// separators, which matches XLSX expectations and keeps storage stable across
    /// UI locale changes.
    pub fn set_cell_formula_localized(
        &mut self,
        sheet: &str,
        addr: &str,
        localized_formula: &str,
        locale: &FormulaLocale,
    ) -> Result<(), EngineError> {
        let canonical = canonicalize_formula(localized_formula, locale)?;
        self.set_cell_formula(sheet, addr, &canonical)
    }

    /// Set a cell formula that was entered in a locale-specific display format using R1C1
    /// reference style.
    ///
    /// This converts the incoming formula to canonical form before translating R1C1 references into
    /// the persisted A1 representation.
    pub fn set_cell_formula_localized_r1c1(
        &mut self,
        sheet: &str,
        addr: &str,
        localized_formula_r1c1: &str,
        locale: &FormulaLocale,
    ) -> Result<(), EngineError> {
        let canonical = canonicalize_formula_with_style(
            localized_formula_r1c1,
            locale,
            crate::ReferenceStyle::R1C1,
        )?;
        self.set_cell_formula_r1c1(sheet, addr, &canonical)
    }

    pub fn get_cell_value(&self, sheet: &str, addr: &str) -> Value {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Value::Blank;
        };
        let Ok(addr) = parse_a1(addr) else {
            return Value::Error(ErrorKind::Ref);
        };
        if let Some(sheet) = self.workbook.sheets.get(sheet_id) {
            if addr.row >= sheet.row_count || addr.col >= sheet.col_count {
                return Value::Error(ErrorKind::Ref);
            }
        }
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        if let Some(v) = self.spilled_cell_value(key) {
            return v;
        }
        if let Some(cell) = self.workbook.get_cell(key) {
            // When using an external value provider (e.g. columnar sheet backing), the engine can
            // store "style-only" cell records (blank + no formula + non-default style) for
            // formatting overlays. These should not clobber provider-backed values.
            //
            // Treat blank non-formula cells as *not* providing a value override so lookups can
            // fall through to the external provider.
            if cell.formula.is_some() || cell.value != Value::Blank {
                return cell.value.clone();
            }
        }

        if let Some(provider) = &self.external_value_provider {
            // Use the workbook's canonical stable sheet key to keep provider lookups stable even
            // when callers pass a different casing or a user-visible display name.
            if let Some(sheet_name) = self.workbook.sheet_key_name(sheet_id) {
                if let Some(v) = provider.get(sheet_name, addr) {
                    return v;
                }
            }
        }

        Value::Blank
    }

    /// Bulk-read a rectangular worksheet range.
    ///
    /// Values are returned in **row-major** order (`values[row][col]`) and include explicit
    /// [`Value::Blank`] entries for any unset cells within the requested rectangle.
    ///
    /// This is intended for performance-sensitive callers (e.g. pivot cache building) that would
    /// otherwise issue per-cell [`Engine::get_cell_value`] calls, which require repeated sheet-name
    /// lookups and A1 parsing.
    pub fn get_range_values(
        &self,
        sheet: &str,
        range: Range,
    ) -> Result<Vec<Vec<Value>>, EngineError> {
        let width = range.width() as usize;
        let height = range.height() as usize;

        let mut out: Vec<Vec<Value>> = Vec::with_capacity(height);
        for _ in 0..height {
            out.push(vec![Value::Blank; width]);
        }

        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(out);
        };
        let Some(sheet_state) = self.workbook.sheets.get(sheet_id) else {
            return Ok(out);
        };

        let row_count = sheet_state.row_count;
        let col_count = sheet_state.col_count;
        let cells = &sheet_state.cells;

        let in_bounds_rows = if range.start.row >= row_count {
            0
        } else {
            let remaining = (row_count - range.start.row) as usize;
            height.min(remaining)
        };
        let in_bounds_cols = if range.start.col >= col_count {
            0
        } else {
            let remaining = (col_count - range.start.col) as usize;
            width.min(remaining)
        };

        // Fill any out-of-bounds cells with `#REF!` to mirror `get_cell_value`.
        if in_bounds_rows < height {
            for row_out in out.iter_mut().skip(in_bounds_rows) {
                row_out.fill(Value::Error(ErrorKind::Ref));
            }
        }
        if in_bounds_cols < width {
            for row_out in out.iter_mut().take(in_bounds_rows) {
                for cell in row_out.iter_mut().skip(in_bounds_cols) {
                    *cell = Value::Error(ErrorKind::Ref);
                }
            }
        }

        // Short-circuit when the entire requested rectangle is out of bounds.
        if in_bounds_rows == 0 || in_bounds_cols == 0 {
            return Ok(out);
        }

        // When an external provider is configured, missing cells can contain non-blank values; we
        // must query it for each missing coordinate, so fall back to per-cell lookup.
        if let Some(provider) = self.external_value_provider.as_deref() {
            let provider_sheet_name = self.workbook.sheet_key_name(sheet_id);
            for row_off in 0..in_bounds_rows {
                let row = range.start.row + row_off as u32;
                for col_off in 0..in_bounds_cols {
                    let col = range.start.col + col_off as u32;
                    let addr = CellAddr { row, col };
                    let key = CellKey {
                        sheet: sheet_id,
                        addr,
                    };

                    if let Some(v) = self.spilled_cell_value(key) {
                        out[row_off][col_off] = v;
                        continue;
                    }

                    if let Some(cell) = cells.get(&addr) {
                        // See `get_cell_value`: style-only blank cells should not hide provider values.
                        if cell.formula.is_some() || cell.value != Value::Blank {
                            out[row_off][col_off] = cell.value.clone();
                            continue;
                        }
                    }

                    if let Some(sheet_name) = provider_sheet_name {
                        if let Some(v) = provider.get(sheet_name, addr) {
                            out[row_off][col_off] = v;
                        }
                    }
                }
            }
            return Ok(out);
        }

        // Heuristic: if the requested rectangle is larger than the number of stored cells, it's
        // often faster to iterate the sparse `Sheet.cells` map and populate a pre-filled output
        // buffer than to perform per-cell HashMap lookups.
        let range_cells = (width as u64).saturating_mul(height as u64);
        let stored_cells = cells.len() as u64;
        let use_sparse_fill = range_cells >= stored_cells;

        if use_sparse_fill {
            for (addr, cell) in cells.iter() {
                if addr.row < range.start.row
                    || addr.row > range.end.row
                    || addr.col < range.start.col
                    || addr.col > range.end.col
                {
                    continue;
                }
                if addr.row >= row_count || addr.col >= col_count {
                    continue;
                }
                let row_off = (addr.row - range.start.row) as usize;
                let col_off = (addr.col - range.start.col) as usize;
                if row_off < in_bounds_rows && col_off < in_bounds_cols {
                    out[row_off][col_off] = cell.value.clone();
                }
            }

            // Overlay spilled values (spill cells override the workbook map for blank/style-only
            // cells).
            for (origin, spill) in &self.spills.by_origin {
                if origin.sheet != sheet_id {
                    continue;
                }
                let spill_start = origin.addr;
                let spill_end = spill.end;
                let start_row = spill_start.row.max(range.start.row);
                let start_col = spill_start.col.max(range.start.col);
                let end_row = spill_end.row.min(range.end.row);
                let end_col = spill_end.col.min(range.end.col);
                if start_row > end_row || start_col > end_col {
                    continue;
                }

                for row in start_row..=end_row {
                    if row >= row_count {
                        break;
                    }
                    let row_off = (row - range.start.row) as usize;
                    if row_off >= in_bounds_rows {
                        continue;
                    }
                    for col in start_col..=end_col {
                        if col >= col_count {
                            break;
                        }
                        let col_off = (col - range.start.col) as usize;
                        if col_off >= in_bounds_cols {
                            continue;
                        }
                        let spill_row_off = (row - spill_start.row) as usize;
                        let spill_col_off = (col - spill_start.col) as usize;
                        if let Some(v) = spill.array.get(spill_row_off, spill_col_off) {
                            out[row_off][col_off] = v.clone();
                        }
                    }
                }
            }
        } else {
            // For small rectangles inside dense sheets, direct per-cell lookups avoid scanning the
            // entire sheet HashMap.
            for row_off in 0..in_bounds_rows {
                let row = range.start.row + row_off as u32;
                for col_off in 0..in_bounds_cols {
                    let col = range.start.col + col_off as u32;
                    let addr = CellAddr { row, col };
                    let key = CellKey {
                        sheet: sheet_id,
                        addr,
                    };

                    if let Some(v) = self.spilled_cell_value(key) {
                        out[row_off][col_off] = v;
                        continue;
                    }
                    if let Some(cell) = cells.get(&addr) {
                        out[row_off][col_off] = cell.value.clone();
                    }
                }
            }
        }

        Ok(out)
    }

    /// Returns the spill range (origin inclusive) for a cell if it is an array-spill
    /// origin or belongs to a spilled range.
    pub fn spill_range(&self, sheet: &str, addr: &str) -> Option<(CellAddr, CellAddr)> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        if let Some(sheet_state) = self.workbook.sheets.get(sheet_id) {
            if addr.row >= sheet_state.row_count || addr.col >= sheet_state.col_count {
                return None;
            }
        }
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let origin = self.spill_origin_key(key)?;
        let spill = self.spills.by_origin.get(&origin)?;
        Some((origin.addr, spill.end))
    }

    /// Returns the spill origin for a cell if it is an array-spill origin or belongs
    /// to a spilled range.
    pub fn spill_origin(&self, sheet: &str, addr: &str) -> Option<(SheetId, CellAddr)> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        if let Some(sheet_state) = self.workbook.sheets.get(sheet_id) {
            if addr.row >= sheet_state.row_count || addr.col >= sheet_state.col_count {
                return None;
            }
        }
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let origin = self.spill_origin_key(key)?;
        Some((origin.sheet, origin.addr))
    }

    pub fn get_cell_formula(&self, sheet: &str, addr: &str) -> Option<&str> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        self.workbook.get_cell(key)?.formula.as_deref()
    }

    pub fn get_cell_phonetic(&self, sheet: &str, addr: &str) -> Option<&str> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        self.workbook.get_cell(key)?.phonetic.as_deref()
    }

    /// Returns the formula for `addr` localized to `locale`.
    ///
    /// The engine persists formulas as canonical A1 strings. This translates the stored formula
    /// on demand using [`crate::locale::localize_formula`].
    pub fn get_cell_formula_localized(
        &self,
        sheet: &str,
        addr: &str,
        locale: &FormulaLocale,
    ) -> Option<String> {
        let formula = self.get_cell_formula(sheet, addr)?;
        localize_formula(formula, locale).ok()
    }

    /// Returns the formula for `addr` rendered in R1C1 reference style.
    ///
    /// The engine persists formulas as canonical A1 strings; this converts them on demand using the
    /// syntax-only parser/serializer.
    pub fn get_cell_formula_r1c1(&self, sheet: &str, addr: &str) -> Option<String> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let formula = self.workbook.get_cell(key)?.formula.as_deref()?;

        let origin = crate::CellAddr::new(addr.row, addr.col);
        let ast = crate::parse_formula(
            formula,
            crate::ParseOptions {
                locale: crate::LocaleConfig::en_us(),
                reference_style: crate::ReferenceStyle::A1,
                normalize_relative_to: Some(origin),
            },
        )
        .ok()?;

        ast.to_string(crate::SerializeOptions {
            locale: crate::LocaleConfig::en_us(),
            reference_style: crate::ReferenceStyle::R1C1,
            include_xlfn_prefix: true,
            origin: Some(origin),
            omit_equals: false,
        })
        .ok()
    }

    /// Returns the formula for `addr` rendered in localized R1C1 reference style.
    ///
    /// The engine stores canonical A1 strings. This first converts the formula to canonical R1C1
    /// using [`Engine::get_cell_formula_r1c1`], then translates it using
    /// [`crate::locale::localize_formula_with_style`].
    pub fn get_cell_formula_localized_r1c1(
        &self,
        sheet: &str,
        addr: &str,
        locale: &FormulaLocale,
    ) -> Option<String> {
        let formula_r1c1 = self.get_cell_formula_r1c1(sheet, addr)?;
        localize_formula_with_style(&formula_r1c1, locale, crate::ReferenceStyle::R1C1).ok()
    }

    pub fn apply_operation(&mut self, op: EditOp) -> Result<EditResult, EditError> {
        let before = self.workbook.clone();
        // We only need to snapshot the pivot registry if the edit can shift worksheet coordinates.
        // Avoid cloning on operations that don't touch pivot destinations (e.g. CopyRange/Fill).
        let mut pivot_registry_before: Option<crate::pivot_registry::PivotRegistry> = None;
        let op_clone = op.clone();
        let mut formula_rewrites = Vec::new();
        let mut moved_ranges = Vec::new();

        let sheet_names = sheet_names_by_id(&self.workbook);
        // For pivot-definition structural edits, treat stable sheet keys and user-visible display
        // names as aliases for the same worksheet.
        let mut sheet_name_to_id: HashMap<String, SheetId> = self.workbook.sheet_key_to_id.clone();
        for (k, v) in &self.workbook.sheet_display_name_to_id {
            sheet_name_to_id.entry(k.clone()).or_insert(*v);
        }
        let edited_sheet_id: SheetId;

        match op {
            EditOp::InsertRows { sheet, row, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                shift_rows(&mut self.workbook.sheets[sheet_id], row, count, true);
                let edit = StructuralEdit::InsertRows {
                    sheet: sheet.clone(),
                    row,
                    count,
                };
                self.rewrite_defined_names_structural(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                if pivot_registry_before.is_none() {
                    pivot_registry_before = Some(self.pivot_registry.clone());
                }
                self.pivot_registry
                    .apply_structural_edit(&edit, &sheet_names);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    edit,
                ));
            }
            EditOp::DeleteRows { sheet, row, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                shift_rows(&mut self.workbook.sheets[sheet_id], row, count, false);
                let edit = StructuralEdit::DeleteRows {
                    sheet: sheet.clone(),
                    row,
                    count,
                };
                self.rewrite_defined_names_structural(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                if pivot_registry_before.is_none() {
                    pivot_registry_before = Some(self.pivot_registry.clone());
                }
                self.pivot_registry
                    .apply_structural_edit(&edit, &sheet_names);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    edit,
                ));
            }
            EditOp::InsertCols { sheet, col, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                shift_cols(&mut self.workbook.sheets[sheet_id], col, count, true);
                update_tables_for_insert_cols(&mut self.workbook.sheets[sheet_id], col, count);
                let edit = StructuralEdit::InsertCols {
                    sheet: sheet.clone(),
                    col,
                    count,
                };
                self.rewrite_defined_names_structural(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                if pivot_registry_before.is_none() {
                    pivot_registry_before = Some(self.pivot_registry.clone());
                }
                self.pivot_registry
                    .apply_structural_edit(&edit, &sheet_names);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    edit,
                ));
            }
            EditOp::DeleteCols { sheet, col, count } => {
                if count == 0 {
                    return Err(EditError::InvalidCount);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                shift_cols(&mut self.workbook.sheets[sheet_id], col, count, false);
                update_tables_for_delete_cols(&mut self.workbook.sheets[sheet_id], col, count);
                let edit = StructuralEdit::DeleteCols {
                    sheet: sheet.clone(),
                    col,
                    count,
                };
                self.rewrite_defined_names_structural(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                if pivot_registry_before.is_none() {
                    pivot_registry_before = Some(self.pivot_registry.clone());
                }
                self.pivot_registry
                    .apply_structural_edit(&edit, &sheet_names);
                formula_rewrites.extend(rewrite_all_formulas_structural(
                    &mut self.workbook,
                    &sheet_names,
                    edit,
                ));
            }
            EditOp::InsertCellsShiftRight { sheet, range } => {
                let width = range.width();
                if width == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                insert_cells_shift_right(&mut self.workbook.sheets[sheet_id], range, width);
                let edit = RangeMapEdit {
                    sheet,
                    moved_region: GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        u32::MAX,
                    ),
                    delta_row: 0,
                    delta_col: width as i32,
                    deleted_region: None,
                };
                self.rewrite_defined_names_range_map(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                if pivot_registry_before.is_none() {
                    pivot_registry_before = Some(self.pivot_registry.clone());
                }
                self.pivot_registry
                    .apply_range_map_edit(&edit, &sheet_names);
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::InsertCellsShiftDown { sheet, range } => {
                let height = range.height();
                if height == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                insert_cells_shift_down(&mut self.workbook.sheets[sheet_id], range, height);
                let edit = RangeMapEdit {
                    sheet,
                    moved_region: GridRange::new(
                        range.start.row,
                        range.start.col,
                        u32::MAX,
                        range.end.col,
                    ),
                    delta_row: height as i32,
                    delta_col: 0,
                    deleted_region: None,
                };
                self.rewrite_defined_names_range_map(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                if pivot_registry_before.is_none() {
                    pivot_registry_before = Some(self.pivot_registry.clone());
                }
                self.pivot_registry
                    .apply_range_map_edit(&edit, &sheet_names);
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::DeleteCellsShiftLeft { sheet, range } => {
                let width = range.width();
                if width == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                delete_cells_shift_left(&mut self.workbook.sheets[sheet_id], range, width);
                let start_col = range.end.col.saturating_add(1);
                let edit = RangeMapEdit {
                    sheet,
                    moved_region: GridRange::new(
                        range.start.row,
                        start_col,
                        range.end.row,
                        u32::MAX,
                    ),
                    delta_row: 0,
                    delta_col: -(width as i32),
                    deleted_region: Some(GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        range.end.col,
                    )),
                };
                self.rewrite_defined_names_range_map(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                if pivot_registry_before.is_none() {
                    pivot_registry_before = Some(self.pivot_registry.clone());
                }
                self.pivot_registry
                    .apply_range_map_edit(&edit, &sheet_names);
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::DeleteCellsShiftUp { sheet, range } => {
                let height = range.height();
                if height == 0 {
                    return Err(EditError::InvalidRange);
                }
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                delete_cells_shift_up(&mut self.workbook.sheets[sheet_id], range, height);
                let start_row = range.end.row.saturating_add(1);
                let edit = RangeMapEdit {
                    sheet,
                    moved_region: GridRange::new(
                        start_row,
                        range.start.col,
                        u32::MAX,
                        range.end.col,
                    ),
                    delta_row: -(height as i32),
                    delta_col: 0,
                    deleted_region: Some(GridRange::new(
                        range.start.row,
                        range.start.col,
                        range.end.row,
                        range.end.col,
                    )),
                };
                self.rewrite_defined_names_range_map(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                if pivot_registry_before.is_none() {
                    pivot_registry_before = Some(self.pivot_registry.clone());
                }
                self.pivot_registry
                    .apply_range_map_edit(&edit, &sheet_names);
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
            }
            EditOp::MoveRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                if src.width() == 0 || src.height() == 0 {
                    return Err(EditError::InvalidRange);
                }
                let dst = Range::new(
                    dst_top_left,
                    CellRef::new(
                        dst_top_left.row + src.height() - 1,
                        dst_top_left.col + src.width() - 1,
                    ),
                );
                if ranges_overlap(src, dst) {
                    return Err(EditError::OverlappingMove);
                }
                move_range(&mut self.workbook.sheets[sheet_id], src, dst_top_left);
                let edit = RangeMapEdit {
                    sheet: sheet.clone(),
                    moved_region: GridRange::new(
                        src.start.row,
                        src.start.col,
                        src.end.row,
                        src.end.col,
                    ),
                    delta_row: dst.start.row as i32 - src.start.row as i32,
                    delta_col: dst.start.col as i32 - src.start.col as i32,
                    deleted_region: None,
                };
                self.rewrite_defined_names_range_map(&sheet_names, &edit)
                    .map_err(|e| EditError::Engine(e.to_string()))?;
                if pivot_registry_before.is_none() {
                    pivot_registry_before = Some(self.pivot_registry.clone());
                }
                self.pivot_registry
                    .apply_range_map_edit(&edit, &sheet_names);
                formula_rewrites.extend(rewrite_all_formulas_range_map(
                    &mut self.workbook,
                    &sheet_names,
                    &edit,
                ));
                moved_ranges.push(MovedRange {
                    sheet,
                    from: src,
                    to: dst,
                });
            }
            EditOp::CopyRange {
                sheet,
                src,
                dst_top_left,
            } => {
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                if src.width() == 0 || src.height() == 0 {
                    return Err(EditError::InvalidRange);
                }
                copy_range(
                    &mut self.workbook.sheets[sheet_id],
                    &sheet,
                    src,
                    dst_top_left,
                    &mut formula_rewrites,
                );
            }
            EditOp::Fill { sheet, src, dst } => {
                let sheet_id = self
                    .workbook
                    .sheet_id(&sheet)
                    .ok_or_else(|| EditError::SheetNotFound(sheet.clone()))?;
                edited_sheet_id = sheet_id;
                fill_range(
                    &mut self.workbook.sheets[sheet_id],
                    &sheet,
                    src,
                    dst,
                    &mut formula_rewrites,
                );
            }
        }

        // Keep pivot table definitions in sync with the structural workbook edit.
        let mut resolve_sheet_id =
            |name: &str| sheet_name_to_id.get(&Workbook::sheet_key(name)).copied();
        for pivot in self.workbook.pivots.values_mut() {
            pivot.apply_edit_op_with_sheet_resolver(&op_clone, &mut resolve_sheet_id);
        }

        if let Err(err) = self.grow_sheet_dimensions_to_fit_cells(edited_sheet_id) {
            // `apply_operation` performs in-place edits to the workbook and only rebuilds the
            // dependency graph after the edit succeeds. If we hit a bounds error here, roll back
            // the workbook so the engine does not end up with mismatched workbook/graph state.
            self.workbook = before;
            if let Some(pivot_registry_before) = pivot_registry_before {
                self.pivot_registry = pivot_registry_before;
            }
            return Err(err);
        }

        self.rebuild_graph()
            .map_err(|e| EditError::Engine(e.to_string()))?;

        let sheet_names_after = sheet_names_by_id(&self.workbook);
        let changed_cells = diff_workbooks(&before, &self.workbook, &sheet_names_after);

        Ok(EditResult {
            changed_cells,
            moved_ranges,
            formula_rewrites,
        })
    }

    fn grow_sheet_dimensions_to_fit_cells(&mut self, sheet_id: SheetId) -> Result<(), EditError> {
        let (max_row, max_col) = {
            let Some(sheet) = self.workbook.sheets.get(sheet_id) else {
                return Ok(());
            };
            let mut max_row: Option<u32> = None;
            let mut max_col: Option<u32> = None;
            for addr in sheet.cells.keys() {
                max_row = Some(max_row.map_or(addr.row, |v| v.max(addr.row)));
                max_col = Some(max_col.map_or(addr.col, |v| v.max(addr.col)));
            }
            match (max_row, max_col) {
                (Some(r), Some(c)) => (r, c),
                _ => return Ok(()),
            }
        };

        // The engine enforces Excel's fixed 16,384-column grid. Editing operations should not be
        // able to create cells beyond that bound.
        if max_col >= EXCEL_MAX_COLS {
            return Err(EditError::Engine(
                crate::eval::AddressParseError::ColumnOutOfRange.to_string(),
            ));
        }
        if max_row >= i32::MAX as u32 {
            return Err(EditError::Engine(
                crate::eval::AddressParseError::RowOutOfRange.to_string(),
            ));
        }

        let max_addr = CellAddr {
            row: max_row,
            col: max_col,
        };
        if self.workbook.grow_sheet_dimensions(sheet_id, max_addr) {
            self.sheet_dims_generation = self.sheet_dims_generation.wrapping_add(1);
        }
        Ok(())
    }

    pub fn recalculate(&mut self) {
        #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
        {
            self.recalculate_with_mode(RecalcMode::MultiThreaded);
        }
        #[cfg(any(target_arch = "wasm32", not(feature = "parallel")))]
        {
            self.recalculate_with_mode(RecalcMode::SingleThreaded);
        }
    }

    pub fn recalculate_single_threaded(&mut self) {
        self.recalculate_with_mode(RecalcMode::SingleThreaded);
    }

    pub fn recalculate_multi_threaded(&mut self) {
        #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
        {
            self.recalculate_with_mode(RecalcMode::MultiThreaded);
        }
        #[cfg(any(target_arch = "wasm32", not(feature = "parallel")))]
        {
            self.recalculate_with_mode(RecalcMode::SingleThreaded);
        }
    }

    fn recalculate_with_mode(&mut self, mode: RecalcMode) {
        self.recalculate_with_mode_and_value_changes(mode, None);
    }

    pub fn recalculate_with_value_changes(&mut self, mode: RecalcMode) -> Vec<RecalcValueChange> {
        let mut changes = RecalcValueChangeCollector::default();
        self.recalculate_with_mode_and_value_changes(mode, Some(&mut changes));
        changes.into_sorted_changes(&self.workbook)
    }

    pub fn recalculate_with_value_changes_single_threaded(&mut self) -> Vec<RecalcValueChange> {
        self.recalculate_with_value_changes(RecalcMode::SingleThreaded)
    }

    pub fn recalculate_with_value_changes_multi_threaded(&mut self) -> Vec<RecalcValueChange> {
        self.recalculate_with_value_changes(RecalcMode::MultiThreaded)
    }

    fn recalculate_with_mode_and_value_changes(
        &mut self,
        mode: RecalcMode,
        value_changes: Option<&mut RecalcValueChangeCollector>,
    ) {
        #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
        {
            if mode == RecalcMode::MultiThreaded {
                if let Some(pool) = recalc_thread_pool() {
                    pool.install(|| {
                        self.recalculate_with_mode_and_value_changes_inner(mode, value_changes)
                    });
                    return;
                }

                // If we can't create a Rayon thread pool (e.g. due to OS resource limits), fall
                // back to single-threaded recalc instead of panicking.
                self.recalculate_with_mode_and_value_changes_inner(
                    RecalcMode::SingleThreaded,
                    value_changes,
                );
                return;
            }
        }

        self.recalculate_with_mode_and_value_changes_inner(mode, value_changes);
    }

    fn recalculate_with_mode_and_value_changes_inner(
        &mut self,
        mode: RecalcMode,
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) {
        let date_system = self.date_system;
        // Spill recalculation can introduce new dirty cells (spill outputs becoming
        // computed/cleared). These should be resolved as part of the same recalc "tick",
        // sharing a single `RecalcContext` so volatile functions remain stable.
        let mut recalc_ctx: Option<crate::eval::RecalcContext> = None;
        loop {
            // Single-threaded recalc does not need dependency levels (parallel batches). Using the
            // cached calculation chain avoids allocating many tiny level vectors for deep chains.
            if mode == RecalcMode::SingleThreaded {
                let order = match self.calc_graph.calc_order_for_dirty() {
                    Ok(order) => order,
                    Err(_) => {
                        self.recalculate_with_cycles(mode, value_changes);
                        return;
                    }
                };

                if order.is_empty() {
                    return;
                }

                // Some workloads (range aggregations / dynamic-deps formulas) benefit from the
                // level schedule because it lets us build per-level range caches and ensures
                // dynamic reference evaluation happens after other work in the same batch.
                let sheet_dims_generation = self.sheet_dims_generation;
                let mut needs_levels = false;
                for &cell_id in &order {
                    let key = cell_key_from_id(cell_id);
                    let Some(cell) = self.workbook.get_cell(key) else {
                        continue;
                    };
                    if cell.dynamic_deps {
                        needs_levels = true;
                        break;
                    }

                    let Some(CompiledFormula::Bytecode(bc)) = cell.compiled.as_ref() else {
                        continue;
                    };

                    // Sheet-dim changes force AST fallback, which does not use the bytecode range cache.
                    if bc.sheet_dims_generation != sheet_dims_generation {
                        continue;
                    }
                    if !bc.program.range_refs.is_empty() || !bc.program.multi_range_refs.is_empty()
                    {
                        needs_levels = true;
                        break;
                    }
                }

                if !needs_levels {
                    let recalc_ctx = recalc_ctx.get_or_insert_with(|| self.begin_recalc_context());
                    let (spill_dirty_roots, dynamic_dirty_roots) = self.recalculate_order(
                        order,
                        recalc_ctx,
                        date_system,
                        value_changes.as_deref_mut(),
                    );
                    if spill_dirty_roots.is_empty() && dynamic_dirty_roots.is_empty() {
                        return;
                    }

                    for cell in spill_dirty_roots.into_iter().chain(dynamic_dirty_roots) {
                        self.calc_graph.mark_dirty(cell);
                    }
                    continue;
                }
            }

            let levels = match self.calc_graph.calc_levels_for_dirty() {
                Ok(levels) => levels,
                Err(_) => {
                    self.recalculate_with_cycles(mode, value_changes);
                    return;
                }
            };

            if levels.is_empty() {
                return;
            }

            let recalc_ctx = recalc_ctx.get_or_insert_with(|| self.begin_recalc_context());

            let (spill_dirty_roots, dynamic_dirty_roots) = self.recalculate_levels(
                levels,
                mode,
                recalc_ctx,
                date_system,
                value_changes.as_deref_mut(),
            );
            if spill_dirty_roots.is_empty() && dynamic_dirty_roots.is_empty() {
                return;
            }

            // Spills can change which cells are considered inputs vs computed spill outputs.
            // When the spill footprint changes (new cells gain/lose values), mark the affected
            // coordinates dirty so any dependents recalculate with the updated spill state.
            for cell in spill_dirty_roots.into_iter().chain(dynamic_dirty_roots) {
                self.calc_graph.mark_dirty(cell);
            }
        }
    }

    fn recalculate_order(
        &mut self,
        order: Vec<CellId>,
        recalc_ctx: &crate::eval::RecalcContext,
        date_system: ExcelDateSystem,
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) -> (Vec<CellId>, Vec<CellId>) {
        self.circular_references.clear();
        let value_locale = self.value_locale;
        let locale_config = self.locale_config.clone();

        let mut snapshot = Snapshot::from_workbook(
            &self.workbook,
            &self.spills,
            self.external_value_provider.clone(),
            self.external_data_provider.clone(),
            self.info.clone(),
            self.pivot_registry.clone(),
        );
        let sheet_dims_generation = self.sheet_dims_generation;
        let mut spill_dirty_roots: Vec<CellId> = Vec::new();
        let dynamic_dirty_roots: Vec<CellId> = Vec::new();
        let text_codepage = self.text_codepage;

        let sheet_count = self.workbook.sheets.len();
        let empty_cols: HashMap<i32, BytecodeColumn> = HashMap::new();
        let empty_cols_by_sheet: Vec<HashMap<i32, BytecodeColumn>> =
            vec![HashMap::new(); sheet_count];
        let cols_by_sheet = empty_cols_by_sheet.as_slice();

        let mut vm = bytecode::Vm::with_capacity(32);
        let _eval_ctx_guard = bytecode::runtime::set_thread_eval_context(
            date_system,
            value_locale,
            recalc_ctx.now_utc.clone(),
            recalc_ctx.recalc_id,
        );

        for cell_id in order {
            let key = cell_key_from_id(cell_id);

            let value = {
                let Some(cell) = self.workbook.get_cell(key) else {
                    continue;
                };
                let Some(compiled_cell) = cell.compiled.as_ref() else {
                    continue;
                };

                let ctx = crate::eval::EvalContext {
                    current_sheet: key.sheet,
                    current_cell: key.addr,
                };

                match compiled_cell {
                    CompiledFormula::Ast(expr) => {
                        let evaluator = crate::eval::Evaluator::new_with_date_system_and_locales(
                            &snapshot,
                            ctx,
                            recalc_ctx,
                            date_system,
                            value_locale,
                            locale_config.clone(),
                        )
                        .with_text_codepage(text_codepage);
                        evaluator.eval_formula(expr)
                    }
                    CompiledFormula::Bytecode(bc) => {
                        if bc.sheet_dims_generation != sheet_dims_generation {
                            let evaluator =
                                crate::eval::Evaluator::new_with_date_system_and_locales(
                                    &snapshot,
                                    ctx,
                                    recalc_ctx,
                                    date_system,
                                    value_locale,
                                    locale_config.clone(),
                                )
                                .with_text_codepage(text_codepage);
                            evaluator.eval_formula(&bc.ast)
                        } else {
                            let cols = cols_by_sheet.get(key.sheet).unwrap_or(&empty_cols);
                            let slice_mode = slice_mode_for_program(&bc.program);
                            let grid = EngineBytecodeGrid {
                                snapshot: &snapshot,
                                sheet_id: key.sheet,
                                cols,
                                cols_by_sheet,
                                slice_mode,
                                trace: None,
                            };
                            let base = bytecode::CellCoord {
                                row: key.addr.row as i32,
                                col: key.addr.col as i32,
                            };
                            let v = vm.eval(&bc.program, &grid, key.sheet, base, &locale_config);
                            bytecode_value_to_engine(v)
                        }
                    }
                }
            };

            self.apply_eval_result(
                key,
                value,
                &mut snapshot,
                &mut spill_dirty_roots,
                value_changes.as_deref_mut(),
            );
        }

        self.calc_graph.clear_dirty();
        self.dirty.clear();
        self.dirty_reasons.clear();

        (spill_dirty_roots, dynamic_dirty_roots)
    }

    fn recalculate_levels(
        &mut self,
        levels: Vec<Vec<CellId>>,
        mode: RecalcMode,
        recalc_ctx: &crate::eval::RecalcContext,
        date_system: ExcelDateSystem,
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) -> (Vec<CellId>, Vec<CellId>) {
        self.circular_references.clear();
        let value_locale = self.value_locale;
        let locale_config = self.locale_config.clone();

        let mut snapshot = Snapshot::from_workbook(
            &self.workbook,
            &self.spills,
            self.external_value_provider.clone(),
            self.external_data_provider.clone(),
            self.info.clone(),
            self.pivot_registry.clone(),
        );
        let sheet_dims_generation = self.sheet_dims_generation;
        let mut spill_dirty_roots: Vec<CellId> = Vec::new();
        let mut dynamic_dirty_roots: Vec<CellId> = Vec::new();
        let text_codepage = self.text_codepage;
        let sheet_count = self.workbook.sheets.len();
        let empty_cols: HashMap<i32, BytecodeColumn> = HashMap::new();
        let empty_cols_by_sheet: Vec<HashMap<i32, BytecodeColumn>> =
            vec![HashMap::new(); sheet_count];

        for level in levels {
            let mut keys: Vec<CellKey> = level.into_iter().map(cell_key_from_id).collect();
            keys.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));

            let mut parallel_tasks: Vec<(CellKey, CompiledFormula)> = Vec::new();
            let mut serial_tasks: Vec<(CellKey, CompiledFormula)> = Vec::new();
            let mut dynamic_tasks: Vec<(CellKey, CompiledFormula)> = Vec::new();
            let mut needs_column_cache = false;

            for &k in &keys {
                let Some(cell) = self.workbook.get_cell(k) else {
                    continue;
                };
                let Some(compiled_cell) = cell.compiled.as_ref() else {
                    continue;
                };
                let compiled = match compiled_cell {
                    CompiledFormula::Ast(expr) => CompiledFormula::Ast(expr.clone()),
                    CompiledFormula::Bytecode(bc) => {
                        if bc.sheet_dims_generation != sheet_dims_generation {
                            // Sheet dimensions changed since this program was compiled; fall back to
                            // AST evaluation so whole-row/whole-column references stay consistent.
                            CompiledFormula::Ast(bc.ast.clone())
                        } else {
                            CompiledFormula::Bytecode(bc.clone())
                        }
                    }
                };

                if let CompiledFormula::Bytecode(bc) = &compiled {
                    if !bc.program.range_refs.is_empty() || !bc.program.multi_range_refs.is_empty()
                    {
                        needs_column_cache = true;
                    }
                }

                if cell.dynamic_deps {
                    dynamic_tasks.push((k, compiled));
                } else if cell.thread_safe {
                    parallel_tasks.push((k, compiled));
                } else {
                    serial_tasks.push((k, compiled));
                }
            }

            let column_cache = if needs_column_cache {
                let mut all_tasks: Vec<(CellKey, CompiledFormula)> = Vec::with_capacity(
                    parallel_tasks.len() + serial_tasks.len() + dynamic_tasks.len(),
                );
                all_tasks.extend(parallel_tasks.iter().cloned());
                all_tasks.extend(serial_tasks.iter().cloned());
                all_tasks.extend(dynamic_tasks.iter().cloned());
                Some(BytecodeColumnCache::build(
                    sheet_count,
                    &snapshot,
                    &all_tasks,
                ))
            } else {
                None
            };
            let cols_by_sheet = column_cache
                .as_ref()
                .map(|cache| cache.by_sheet.as_slice())
                .unwrap_or(empty_cols_by_sheet.as_slice());

            let mut results: Vec<(CellKey, Value)> =
                Vec::with_capacity(parallel_tasks.len() + serial_tasks.len());
            let eval_parallel_tasks_serial = |results: &mut Vec<(CellKey, Value)>| {
                let mut vm = bytecode::Vm::with_capacity(32);
                let _eval_ctx_guard = bytecode::runtime::set_thread_eval_context(
                    date_system,
                    value_locale,
                    recalc_ctx.now_utc.clone(),
                    recalc_ctx.recalc_id,
                );
                for (k, compiled) in &parallel_tasks {
                    let ctx = crate::eval::EvalContext {
                        current_sheet: k.sheet,
                        current_cell: k.addr,
                    };
                    let value = match compiled {
                        CompiledFormula::Ast(expr) => {
                            let evaluator =
                                crate::eval::Evaluator::new_with_date_system_and_locales(
                                    &snapshot,
                                    ctx,
                                    recalc_ctx,
                                    date_system,
                                    value_locale,
                                    locale_config.clone(),
                                )
                                .with_text_codepage(text_codepage);
                            evaluator.eval_formula(expr)
                        }
                        CompiledFormula::Bytecode(bc) => {
                            let cols = cols_by_sheet.get(k.sheet).unwrap_or(&empty_cols);
                            let slice_mode = slice_mode_for_program(&bc.program);
                            let grid = EngineBytecodeGrid {
                                snapshot: &snapshot,
                                sheet_id: k.sheet,
                                cols,
                                cols_by_sheet,
                                slice_mode,
                                trace: None,
                            };
                            let base = bytecode::CellCoord {
                                row: k.addr.row as i32,
                                col: k.addr.col as i32,
                            };
                            let v = vm.eval(&bc.program, &grid, k.sheet, base, &locale_config);
                            bytecode_value_to_engine(v)
                        }
                    };
                    results.push((*k, value));
                }
            };

            if mode == RecalcMode::MultiThreaded {
                #[cfg(all(feature = "parallel", not(target_arch = "wasm32")))]
                {
                    if let Some(pool) = crate::parallel::rayon_pool() {
                        results.extend(pool.install(|| {
                            parallel_tasks
                                .par_iter()
                                .map_init(
                                    || {
                                        (
                                            bytecode::Vm::with_capacity(32),
                                            bytecode::runtime::set_thread_eval_context(
                                                date_system,
                                                value_locale,
                                                recalc_ctx.now_utc.clone(),
                                                recalc_ctx.recalc_id,
                                            ),
                                        )
                                    },
                                    |(vm, _eval_ctx_guard), (k, compiled)| {
                                        let ctx = crate::eval::EvalContext {
                                            current_sheet: k.sheet,
                                            current_cell: k.addr,
                                        };

                                        match compiled {
                                            CompiledFormula::Ast(expr) => {
                                                let evaluator =
                                                    crate::eval::Evaluator::new_with_date_system_and_locales(
                                                        &snapshot,
                                                        ctx,
                                                        recalc_ctx,
                                                        date_system,
                                                        value_locale,
                                                        locale_config.clone(),
                                                    )
                                                    .with_text_codepage(text_codepage);
                                                (*k, evaluator.eval_formula(expr))
                                            }
                                            CompiledFormula::Bytecode(bc) => {
                                                let cols =
                                                    cols_by_sheet.get(k.sheet).unwrap_or(&empty_cols);
                                                let slice_mode = slice_mode_for_program(&bc.program);
                                                let grid = EngineBytecodeGrid {
                                                    snapshot: &snapshot,
                                                    sheet_id: k.sheet,
                                                    cols,
                                                    cols_by_sheet,
                                                    slice_mode,
                                                    trace: None,
                                                };
                                                let base = bytecode::CellCoord {
                                                    row: k.addr.row as i32,
                                                    col: k.addr.col as i32,
                                                };
                                                let v = vm.eval(
                                                    &bc.program,
                                                    &grid,
                                                    k.sheet,
                                                    base,
                                                    &locale_config,
                                                );
                                                (*k, bytecode_value_to_engine(v))
                                            }
                                        }
                                    },
                                )
                                .collect::<Vec<_>>()
                        }));
                    } else {
                        // If we can't initialize a thread pool (e.g. under thread-constrained CI
                        // environments), fall back to single-threaded evaluation rather than
                        // panicking inside Rayon.
                        eval_parallel_tasks_serial(&mut results);
                    }
                }

                #[cfg(not(all(feature = "parallel", not(target_arch = "wasm32"))))]
                {
                    eval_parallel_tasks_serial(&mut results);
                }
            } else {
                eval_parallel_tasks_serial(&mut results);
            }

            // Non-thread-safe tasks are always serialized.
            let mut vm = bytecode::Vm::with_capacity(32);
            let _eval_ctx_guard = bytecode::runtime::set_thread_eval_context(
                date_system,
                value_locale,
                recalc_ctx.now_utc.clone(),
                recalc_ctx.recalc_id,
            );
            for (k, compiled) in &serial_tasks {
                let ctx = crate::eval::EvalContext {
                    current_sheet: k.sheet,
                    current_cell: k.addr,
                };
                let value = match compiled {
                    CompiledFormula::Ast(expr) => {
                        let evaluator = crate::eval::Evaluator::new_with_date_system_and_locales(
                            &snapshot,
                            ctx,
                            recalc_ctx,
                            date_system,
                            value_locale,
                            locale_config.clone(),
                        )
                        .with_text_codepage(text_codepage);
                        evaluator.eval_formula(expr)
                    }
                    CompiledFormula::Bytecode(bc) => {
                        let cols = cols_by_sheet.get(k.sheet).unwrap_or(&empty_cols);
                        let slice_mode = slice_mode_for_program(&bc.program);
                        let grid = EngineBytecodeGrid {
                            snapshot: &snapshot,
                            sheet_id: k.sheet,
                            cols,
                            cols_by_sheet,
                            slice_mode,
                            trace: None,
                        };
                        let base = bytecode::CellCoord {
                            row: k.addr.row as i32,
                            col: k.addr.col as i32,
                        };
                        let v = vm.eval(&bc.program, &grid, k.sheet, base, &locale_config);
                        bytecode_value_to_engine(v)
                    }
                };
                results.push((*k, value));
            }

            results.sort_by_key(|(k, _)| (k.sheet, k.addr.row, k.addr.col));

            for (k, v) in results {
                self.apply_eval_result(
                    k,
                    v,
                    &mut snapshot,
                    &mut spill_dirty_roots,
                    value_changes.as_deref_mut(),
                );
            }

            // Dynamic-reference formulas (e.g. INDIRECT/OFFSET) must be evaluated serially so we
            // can trace their runtime precedents and update the dependency graph deterministically.
            for (k, compiled) in &dynamic_tasks {
                let ctx = crate::eval::EvalContext {
                    current_sheet: k.sheet,
                    current_cell: k.addr,
                };

                let trace = RefCell::new(crate::eval::DependencyTrace::default());
                let bytecode_trace = Mutex::new(crate::eval::DependencyTrace::default());
                let mut used_bytecode_trace = false;

                let value = match compiled {
                    CompiledFormula::Ast(expr) => {
                        let evaluator = crate::eval::Evaluator::new_with_date_system_and_locales(
                            &snapshot,
                            ctx,
                            recalc_ctx,
                            date_system,
                            value_locale,
                            locale_config.clone(),
                        )
                        .with_text_codepage(text_codepage)
                        .with_dependency_trace(&trace);
                        evaluator.eval_formula(expr)
                    }
                    CompiledFormula::Bytecode(bc) => {
                        used_bytecode_trace = true;
                        let cols = cols_by_sheet.get(k.sheet).unwrap_or(&empty_cols);
                        let slice_mode = slice_mode_for_program(&bc.program);
                        let grid = EngineBytecodeGrid {
                            snapshot: &snapshot,
                            sheet_id: k.sheet,
                            cols,
                            cols_by_sheet,
                            slice_mode,
                            trace: Some(&bytecode_trace),
                        };
                        let base = bytecode::CellCoord {
                            row: k.addr.row as i32,
                            col: k.addr.col as i32,
                        };
                        let v = vm.eval(&bc.program, &grid, k.sheet, base, &locale_config);
                        bytecode_value_to_engine(v)
                    }
                };

                self.apply_eval_result(
                    *k,
                    value,
                    &mut snapshot,
                    &mut spill_dirty_roots,
                    value_changes.as_deref_mut(),
                );

                let traced_precedents = {
                    let tab_index_by_sheet_id = self.workbook.tab_index_by_sheet_id();
                    if used_bytecode_trace {
                        let guard = match bytecode_trace.lock() {
                            Ok(g) => g,
                            Err(poisoned) => poisoned.into_inner(),
                        };
                        guard.precedents(|sheet_id| {
                            tab_index_by_sheet_id
                                .get(sheet_id)
                                .copied()
                                .unwrap_or(usize::MAX)
                        })
                    } else {
                        trace.borrow().precedents(|sheet_id| {
                            tab_index_by_sheet_id
                                .get(sheet_id)
                                .copied()
                                .unwrap_or(usize::MAX)
                        })
                    }
                };
                let expr = compiled.ast();

                let cell_id = cell_id_from_key(*k);
                let old_precedents: HashSet<Precedent> =
                    self.calc_graph.precedents_of(cell_id).into_iter().collect();

                let mut new_precedents: HashSet<Precedent> = analyze_calc_precedents(
                    expr,
                    *k,
                    &snapshot.tables,
                    &self.workbook,
                    &self.spills,
                );
                let static_cell_precedents: HashSet<CellId> = new_precedents
                    .iter()
                    .filter_map(|p| match p {
                        Precedent::Cell(c) => Some(*c),
                        _ => None,
                    })
                    .collect();
                let mut dynamic_external_precedents: HashSet<crate::functions::Reference> =
                    HashSet::new();
                for reference in traced_precedents {
                    let crate::functions::Reference {
                        sheet_id,
                        start,
                        end,
                    } = reference;
                    match sheet_id {
                        crate::functions::SheetId::Local(sheet_id) => {
                            let sheet_id = sheet_id_for_graph(sheet_id);
                            if start == end {
                                new_precedents.insert(Precedent::Cell(CellId::new(
                                    sheet_id, start.row, start.col,
                                )));
                            } else {
                                let range = Range::new(
                                    CellRef::new(start.row, start.col),
                                    CellRef::new(end.row, end.col),
                                );
                                new_precedents
                                    .insert(Precedent::Range(SheetRange::new(sheet_id, range)));
                            }
                        }
                        crate::functions::SheetId::External(key) => {
                            // External references can't be represented in the internal dependency graph yet.
                            //
                            // By default, these are handled via the engine's volatile external-ref semantics;
                            // when external refs are configured as non-volatile, callers must explicitly
                            // invalidate dependents via `mark_external_*_dirty`.
                            //
                            // Persist them separately so auditing + explicit invalidation can still surface them.
                            if crate::eval::split_external_sheet_key(&key).is_some() {
                                dynamic_external_precedents.insert(crate::functions::Reference {
                                    sheet_id: crate::functions::SheetId::External(key),
                                    start,
                                    end,
                                });
                            }
                        }
                    }
                }
                self.set_cell_dynamic_external_precedents(*k, dynamic_external_precedents);

                // Dynamic dependency tracing can record both a range and individual cells within
                // that range (e.g. INDEX(OFFSET(...), ...) records the OFFSET range, then the
                // evaluator dereferences a single cell from the INDEX result).
                //
                // Once a range precedent exists, the contained cells are redundant for calculation.
                // However, preserve any cell precedents that were present in the static analysis
                // (direct references in the formula) so auditing remains informative.
                let ranges: Vec<SheetRange> = new_precedents
                    .iter()
                    .filter_map(|p| match p {
                        Precedent::Range(r) => Some(*r),
                        _ => None,
                    })
                    .collect();
                if !ranges.is_empty() {
                    new_precedents.retain(|p| match p {
                        Precedent::Cell(cell) => {
                            static_cell_precedents.contains(cell)
                                || !ranges.iter().any(|r| r.contains(*cell))
                        }
                        _ => true,
                    });
                }

                if new_precedents != old_precedents {
                    let mut vec: Vec<Precedent> = new_precedents.into_iter().collect();
                    vec.sort_by_key(|p| match p {
                        Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col, 0u32, 0u32),
                        Precedent::Range(r) => (
                            1u8,
                            r.sheet_id,
                            r.range.start.row,
                            r.range.start.col,
                            r.range.end.row,
                            r.range.end.col,
                        ),
                    });

                    let is_volatile = self
                        .workbook
                        .get_cell(*k)
                        .map(|c| c.volatile)
                        .unwrap_or(false);
                    self.calc_graph.update_cell_dependencies(
                        cell_id,
                        CellDeps::new(vec).volatile(is_volatile),
                    );
                    dynamic_dirty_roots.push(cell_id);
                }
            }
        }

        self.calc_graph.clear_dirty();
        self.dirty.clear();
        self.dirty_reasons.clear();

        (spill_dirty_roots, dynamic_dirty_roots)
    }

    fn recalculate_with_cycles(
        &mut self,
        _mode: RecalcMode,
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) {
        let mut impacted_ids: HashSet<CellId> = self.calc_graph.dirty_cells().into_iter().collect();
        impacted_ids.extend(self.calc_graph.volatile_cells());

        if impacted_ids.is_empty() {
            return;
        }

        let recalc_ctx = self.begin_recalc_context();

        self.circular_references.clear();

        let mut impacted: Vec<CellKey> = impacted_ids.into_iter().map(cell_key_from_id).collect();
        impacted.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));

        let impacted_set: HashSet<CellKey> = impacted.iter().copied().collect();
        let mut edges: HashMap<CellKey, Vec<CellKey>> = HashMap::new();
        for &cell in &impacted {
            let cell_id = cell_id_from_key(cell);
            let mut out: Vec<CellKey> = self
                .calc_graph
                .direct_dependents(cell_id)
                .into_iter()
                .map(cell_key_from_id)
                .filter(|d| impacted_set.contains(d))
                .collect();
            if out.is_empty() {
                continue;
            }
            out.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));
            edges.insert(cell, out);
        }

        let sccs = iterative::strongly_connected_components(&impacted, &edges);
        let order = iterative::topo_sort_sccs(&sccs, &edges);

        let mut snapshot = Snapshot::from_workbook(
            &self.workbook,
            &self.spills,
            self.external_value_provider.clone(),
            self.external_data_provider.clone(),
            self.info.clone(),
            self.pivot_registry.clone(),
        );
        let mut spill_dirty_roots: Vec<CellId> = Vec::new();
        let date_system = self.date_system;
        let value_locale = self.value_locale;
        let locale_config = self.locale_config.clone();
        let text_codepage = self.text_codepage;

        for scc_idx in order {
            let mut scc = sccs[scc_idx].clone();
            scc.sort_by_key(|k| (k.sheet, k.addr.row, k.addr.col));

            let is_cycle = match scc.as_slice() {
                [] => continue,
                [only] => edges
                    .get(only)
                    .map(|deps| deps.contains(only))
                    .unwrap_or(false),
                _ => true,
            };

            if !is_cycle {
                let k = scc[0];
                let Some(expr) = self
                    .workbook
                    .get_cell(k)
                    .and_then(|c| c.compiled.as_ref().map(|compiled| compiled.ast().clone()))
                else {
                    continue;
                };
                let ctx = crate::eval::EvalContext {
                    current_sheet: k.sheet,
                    current_cell: k.addr,
                };
                let evaluator = crate::eval::Evaluator::new_with_date_system_and_locales(
                    &snapshot,
                    ctx,
                    &recalc_ctx,
                    date_system,
                    value_locale,
                    locale_config.clone(),
                )
                .with_text_codepage(text_codepage);
                let v = evaluator.eval_formula(&expr);
                self.apply_eval_result(
                    k,
                    v,
                    &mut snapshot,
                    &mut spill_dirty_roots,
                    value_changes.as_deref_mut(),
                );
                continue;
            }

            for &k in &scc {
                self.circular_references.insert(k);
            }

            if !self.calc_settings.iterative.enabled {
                for &k in &scc {
                    let v = Value::Number(0.0);
                    self.apply_eval_result(
                        k,
                        v,
                        &mut snapshot,
                        &mut spill_dirty_roots,
                        value_changes.as_deref_mut(),
                    );
                }
                continue;
            }

            let max_iters = max(1, self.calc_settings.iterative.max_iterations) as usize;
            let tol = self.calc_settings.iterative.max_change.max(0.0);

            for _ in 0..max_iters {
                let mut max_delta: f64 = 0.0;
                for &k in &scc {
                    let Some(expr) = self
                        .workbook
                        .get_cell(k)
                        .and_then(|c| c.compiled.as_ref().map(|compiled| compiled.ast().clone()))
                    else {
                        continue;
                    };
                    let old = snapshot.values.get(&k).cloned().unwrap_or(Value::Blank);
                    let ctx = crate::eval::EvalContext {
                        current_sheet: k.sheet,
                        current_cell: k.addr,
                    };
                    let evaluator = crate::eval::Evaluator::new_with_date_system_and_locales(
                        &snapshot,
                        ctx,
                        &recalc_ctx,
                        date_system,
                        value_locale,
                        locale_config.clone(),
                    )
                    .with_text_codepage(text_codepage);
                    let new_val = evaluator.eval_formula(&expr);
                    max_delta = max_delta.max(value_delta(&old, &new_val));
                    self.apply_eval_result(
                        k,
                        new_val,
                        &mut snapshot,
                        &mut spill_dirty_roots,
                        value_changes.as_deref_mut(),
                    );
                }

                if max_delta <= tol {
                    break;
                }
            }
        }

        self.calc_graph.clear_dirty();
        self.dirty.clear();
        self.dirty_reasons.clear();
    }

    fn apply_eval_result(
        &mut self,
        key: CellKey,
        value: Value,
        snapshot: &mut Snapshot,
        spill_dirty_roots: &mut Vec<CellId>,
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) {
        // Clear any previously tracked spill blockage for this origin before applying the new
        // evaluation result.
        self.clear_blocked_spill_for_origin(key);

        let value = match value {
            Value::Lambda(_) => Value::Error(ErrorKind::Calc),
            other => other,
        };

        // The number-format lookup is only needed for Excel's "precision as displayed" mode.
        // In full-precision mode (the Excel default), `round_number_as_displayed` ignores the
        // format string, so avoid the repeated style/format lookups on hot recalc paths.
        let format_pattern = if self.calc_settings.full_precision {
            None
        } else {
            self.number_format_pattern_for_rounding(key)
        };

        match value {
            Value::Array(mut array) => {
                if array.rows == 0 || array.cols == 0 {
                    self.apply_eval_result(
                        key,
                        Value::Error(ErrorKind::Calc),
                        snapshot,
                        spill_dirty_roots,
                        value_changes,
                    );
                    return;
                }

                // Excel treats bare lambda values (not invoked) as `#CALC!`.
                // This applies to each spilled cell as well, so coerce any lambda elements in a
                // dynamic array result to `#CALC!` before the spill is materialized.
                for value in &mut array.values {
                    if matches!(value, Value::Lambda(_)) {
                        *value = Value::Error(ErrorKind::Calc);
                    }
                }

                if !self.calc_settings.full_precision {
                    for value in &mut array.values {
                        if let Value::Number(n) = value {
                            *n = self.round_number_as_displayed(*n, format_pattern);
                        }
                    }
                }

                let (sheet_rows, sheet_cols) = self
                    .workbook
                    .sheets
                    .get(key.sheet)
                    .map(|s| (s.row_count, s.col_count))
                    .unwrap_or((EXCEL_MAX_ROWS, EXCEL_MAX_COLS));

                let mut spill_too_big = || {
                    let cleared = self.clear_spill_for_origin(key);
                    snapshot.spill_end_by_origin.remove(&key);
                    for cleared_key in cleared {
                        if let Some(changes) = value_changes.as_deref_mut() {
                            let before =
                                snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                            snapshot.remove_value(&cleared_key);
                            let after =
                                snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                            changes.record(cleared_key, before, after);
                        } else {
                            snapshot.remove_value(&cleared_key);
                        }
                        snapshot.spill_origin_by_cell.remove(&cleared_key);
                        spill_dirty_roots.push(cell_id_from_key(cleared_key));
                        self.append_blocked_spill_dirty_roots(cleared_key, spill_dirty_roots);
                    }

                    let after = Value::Error(ErrorKind::Spill);
                    if let Some(changes) = value_changes.as_deref_mut() {
                        let before = snapshot.get_cell_value(key.sheet, key.addr);
                        changes.record(key, before, after.clone());
                    }
                    let cell = self.workbook.get_or_create_cell_mut(key);
                    cell.value = after.clone();
                    snapshot.insert_value(key, after);
                };

                let row_delta = match u32::try_from(array.rows.saturating_sub(1)) {
                    Ok(v) => v,
                    Err(_) => {
                        spill_too_big();
                        return;
                    }
                };
                let col_delta = match u32::try_from(array.cols.saturating_sub(1)) {
                    Ok(v) => v,
                    Err(_) => {
                        spill_too_big();
                        return;
                    }
                };

                let Some(end_row) = key.addr.row.checked_add(row_delta) else {
                    spill_too_big();
                    return;
                };
                let Some(end_col) = key.addr.col.checked_add(col_delta) else {
                    spill_too_big();
                    return;
                };

                let end = CellAddr {
                    row: end_row,
                    col: end_col,
                };

                // If the spilled result would extend beyond the sheet bounds, the origin evaluates
                // to `#SPILL!` ("Spill range is too big") and no out-of-bounds spill cells should
                // be materialized.
                if end_row >= sheet_rows || end_col >= sheet_cols {
                    spill_too_big();
                    return;
                }

                // Fast path: if the spill shape is unchanged, update the stored array and
                // overwrite spill cell values in the snapshot without reshaping dependency nodes.
                if let Some(existing) = self.spills.by_origin.get_mut(&key) {
                    if existing.end == end {
                        existing.array = array.clone();

                        let top_left = array.top_left();
                        if let Some(changes) = value_changes.as_deref_mut() {
                            let before = snapshot.get_cell_value(key.sheet, key.addr);
                            changes.record(key, before, top_left.clone());
                        }
                        let cell = self.workbook.get_or_create_cell_mut(key);
                        cell.value = top_left.clone();
                        snapshot.insert_value(key, top_left);

                        for r in 0..array.rows {
                            for c in 0..array.cols {
                                if r == 0 && c == 0 {
                                    continue;
                                }
                                let addr = CellAddr {
                                    row: key.addr.row + r as u32,
                                    col: key.addr.col + c as u32,
                                };
                                let spill_key = CellKey {
                                    sheet: key.sheet,
                                    addr,
                                };
                                if let Some(v) = array.get(r, c).cloned() {
                                    if let Some(changes) = value_changes.as_deref_mut() {
                                        let before = snapshot
                                            .get_cell_value(spill_key.sheet, spill_key.addr);
                                        changes.record(spill_key, before, v.clone());
                                    }
                                    snapshot.insert_value(spill_key, v);
                                }
                            }
                        }
                        return;
                    }
                }

                // Spill footprint change: clear the previous spill range (if any) before attempting
                // to write the new one.
                let cleared = self.clear_spill_for_origin(key);
                snapshot.spill_end_by_origin.remove(&key);
                for cleared_key in cleared {
                    if let Some(changes) = value_changes.as_deref_mut() {
                        let before = snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                        snapshot.remove_value(&cleared_key);
                        let after = snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                        changes.record(cleared_key, before, after);
                    } else {
                        snapshot.remove_value(&cleared_key);
                    }
                    snapshot.spill_origin_by_cell.remove(&cleared_key);
                    spill_dirty_roots.push(cell_id_from_key(cleared_key));
                    self.append_blocked_spill_dirty_roots(cleared_key, spill_dirty_roots);
                }

                if let Some(blocker) = self.spill_blocker(key, &array) {
                    self.record_blocked_spill(key, blocker);
                    let after = Value::Error(ErrorKind::Spill);
                    if let Some(changes) = value_changes.as_deref_mut() {
                        let before = snapshot.get_cell_value(key.sheet, key.addr);
                        changes.record(key, before, after.clone());
                    }
                    let cell = self.workbook.get_or_create_cell_mut(key);
                    cell.value = after.clone();
                    snapshot.insert_value(key, after);
                    return;
                }

                self.apply_new_spill(key, end, array, snapshot, spill_dirty_roots, value_changes);
            }
            other => {
                let other = match other {
                    Value::Number(n) => {
                        Value::Number(self.round_number_as_displayed(n, format_pattern))
                    }
                    v => v,
                };

                let cleared = self.clear_spill_for_origin(key);
                snapshot.spill_end_by_origin.remove(&key);
                for cleared_key in cleared {
                    if let Some(changes) = value_changes.as_deref_mut() {
                        let before = snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                        snapshot.remove_value(&cleared_key);
                        let after = snapshot.get_cell_value(cleared_key.sheet, cleared_key.addr);
                        changes.record(cleared_key, before, after);
                    } else {
                        snapshot.remove_value(&cleared_key);
                    }
                    snapshot.spill_origin_by_cell.remove(&cleared_key);
                    spill_dirty_roots.push(cell_id_from_key(cleared_key));
                    self.append_blocked_spill_dirty_roots(cleared_key, spill_dirty_roots);
                }

                if let Some(changes) = value_changes.as_deref_mut() {
                    let before = snapshot.get_cell_value(key.sheet, key.addr);
                    changes.record(key, before, other.clone());
                }
                let cell = self.workbook.get_or_create_cell_mut(key);
                cell.value = other.clone();
                snapshot.insert_value(key, other);
            }
        }
    }

    fn spill_origin_key(&self, key: CellKey) -> Option<CellKey> {
        if self.spills.by_origin.contains_key(&key) {
            return Some(key);
        }
        self.spills.origin_by_cell.get(&key).copied()
    }

    fn spilled_cell_value(&self, key: CellKey) -> Option<Value> {
        let origin = self.spills.origin_by_cell.get(&key).copied()?;
        let spill = self.spills.by_origin.get(&origin)?;
        let row_off = key.addr.row.checked_sub(origin.addr.row)? as usize;
        let col_off = key.addr.col.checked_sub(origin.addr.col)? as usize;
        spill.array.get(row_off, col_off).cloned()
    }

    fn clear_blocked_spill_for_origin(&mut self, origin: CellKey) {
        let Some(blocked) = self.spills.blocked_by_origin.remove(&origin) else {
            return;
        };

        if let Some(origins) = self
            .spills
            .blocked_origins_by_cell
            .get_mut(&blocked.blocker)
        {
            origins.remove(&origin);
            if origins.is_empty() {
                self.spills.blocked_origins_by_cell.remove(&blocked.blocker);
            }
        }
    }

    fn record_blocked_spill(&mut self, origin: CellKey, blocker: CellKey) {
        self.spills
            .blocked_by_origin
            .insert(origin, BlockedSpill { blocker });
        self.spills
            .blocked_origins_by_cell
            .entry(blocker)
            .or_default()
            .insert(origin);
    }

    fn mark_dirty_blocked_spill_origins_for_cell(&mut self, cell: CellKey) {
        let Some(origins) = self.spills.blocked_origins_by_cell.get(&cell) else {
            return;
        };

        // Clone so we can freely mutate dirty bookkeeping while iterating.
        let origins: Vec<CellKey> = origins.iter().copied().collect();
        for origin in origins {
            let origin_id = cell_id_from_key(origin);
            self.calc_graph.mark_dirty(origin_id);

            if self.dirty.insert(origin) {
                self.dirty_reasons.insert(origin, DirtyReason::Cell(cell));
            }
            self.mark_dirty_dependents_with_reasons(origin);
        }
    }

    fn append_blocked_spill_dirty_roots(&self, cell: CellKey, out: &mut Vec<CellId>) {
        let Some(origins) = self.spills.blocked_origins_by_cell.get(&cell) else {
            return;
        };
        for &origin in origins {
            out.push(cell_id_from_key(origin));
        }
    }

    fn clear_spill_for_cell(&mut self, key: CellKey) {
        let origin = match self.spill_origin_key(key) {
            Some(origin) => origin,
            None => return,
        };

        let cleared = self.clear_spill_for_origin(origin);
        for cleared_key in cleared {
            self.calc_graph.mark_dirty(cell_id_from_key(cleared_key));
            self.mark_dirty_blocked_spill_origins_for_cell(cleared_key);
        }

        // If a user edits any cell in a spill range, the origin needs to be re-evaluated
        // to either re-spill (if the blockage was removed) or surface #SPILL! (if blocked).
        if origin != key {
            self.calc_graph.mark_dirty(cell_id_from_key(origin));
        }
    }

    fn clear_spill_for_origin(&mut self, origin: CellKey) -> Vec<CellKey> {
        let Some(spill) = self.spills.by_origin.remove(&origin) else {
            return Vec::new();
        };

        let mut cleared = Vec::new();
        for r in 0..spill.array.rows {
            for c in 0..spill.array.cols {
                if r == 0 && c == 0 {
                    continue;
                }
                let addr = CellAddr {
                    row: origin.addr.row + r as u32,
                    col: origin.addr.col + c as u32,
                };
                let key = CellKey {
                    sheet: origin.sheet,
                    addr,
                };
                self.spills.origin_by_cell.remove(&key);
                self.calc_graph.remove_cell(cell_id_from_key(key));
                cleared.push(key);
            }
        }
        cleared
    }

    fn spill_blocker(&self, origin: CellKey, array: &Array) -> Option<CellKey> {
        for r in 0..array.rows {
            for c in 0..array.cols {
                if r == 0 && c == 0 {
                    continue;
                }
                let addr = CellAddr {
                    row: origin.addr.row + r as u32,
                    col: origin.addr.col + c as u32,
                };
                let key = CellKey {
                    sheet: origin.sheet,
                    addr,
                };

                // Blocked by non-empty user cell (literal or formula).
                if let Some(cell) = self.workbook.get_cell(key) {
                    if cell.formula.is_some() {
                        return Some(key);
                    }
                    if cell.value != Value::Blank {
                        return Some(key);
                    }
                }

                // Blocked by external provider values (e.g. columnar backing). Style-only cells
                // (blank + no formula) should not mask provider-backed blockers.
                if let Some(provider) = &self.external_value_provider {
                    if let Some(sheet_name) = self.workbook.sheet_key_name(origin.sheet) {
                        if let Some(v) = provider.get(sheet_name, addr) {
                            if v != Value::Blank {
                                return Some(key);
                            }
                        }
                    }
                }

                // Blocked by another spill.
                if let Some(other_origin) = self.spills.origin_by_cell.get(&key) {
                    if *other_origin != origin {
                        return Some(key);
                    }
                }
            }
        }
        None
    }

    fn apply_new_spill(
        &mut self,
        origin: CellKey,
        end: CellAddr,
        array: Array,
        snapshot: &mut Snapshot,
        spill_dirty_roots: &mut Vec<CellId>,
        mut value_changes: Option<&mut RecalcValueChangeCollector>,
    ) {
        let top_left = array.top_left();

        if let Some(changes) = value_changes.as_deref_mut() {
            let before = snapshot.get_cell_value(origin.sheet, origin.addr);
            changes.record(origin, before, top_left.clone());
        }

        let cell = self.workbook.get_or_create_cell_mut(origin);
        cell.value = top_left.clone();
        snapshot.insert_value(origin, top_left);

        self.spills.by_origin.insert(
            origin,
            Spill {
                end,
                array: array.clone(),
            },
        );
        snapshot.spill_end_by_origin.insert(origin, end);

        let origin_id = cell_id_from_key(origin);
        for r in 0..array.rows {
            for c in 0..array.cols {
                if r == 0 && c == 0 {
                    continue;
                }
                let addr = CellAddr {
                    row: origin.addr.row + r as u32,
                    col: origin.addr.col + c as u32,
                };
                let key = CellKey {
                    sheet: origin.sheet,
                    addr,
                };
                self.spills.origin_by_cell.insert(key, origin);
                snapshot.spill_origin_by_cell.insert(key, origin);

                if let Some(v) = array.get(r, c).cloned() {
                    if let Some(changes) = value_changes.as_deref_mut() {
                        let before = snapshot.get_cell_value(key.sheet, key.addr);
                        changes.record(key, before, v.clone());
                    }
                    snapshot.insert_value(key, v);
                }

                // Register spill cells as formula nodes that depend on the origin so they participate in
                // calculation ordering and dirty marking.
                let deps = CellDeps::new(vec![Precedent::Cell(origin_id)]);
                self.calc_graph
                    .update_cell_dependencies(cell_id_from_key(key), deps);

                spill_dirty_roots.push(cell_id_from_key(key));
            }
        }
    }

    fn try_compile_bytecode(
        &self,
        expr: &crate::Expr,
        key: CellKey,
        thread_safe: bool,
        _dynamic_deps: bool,
    ) -> Result<Arc<bytecode::Program>, BytecodeCompileReason> {
        if !self.bytecode_enabled {
            return Err(BytecodeCompileReason::Disabled);
        }
        if !thread_safe {
            return Err(BytecodeCompileReason::NotThreadSafe);
        }

        let origin_ast = crate::CellAddr::new(key.addr.row, key.addr.col);
        let origin = bytecode::CellCoord {
            row: i32::try_from(key.addr.row)
                .map_err(|_| BytecodeCompileReason::ExceedsGridLimits)?,
            col: i32::try_from(key.addr.col)
                .map_err(|_| BytecodeCompileReason::ExceedsGridLimits)?,
        };
        // Inline any defined names that resolve to static expressions/values before lowering.
        // This improves bytecode eligibility and keeps the bytecode backend deterministic.
        let expr = self.inline_static_defined_names_for_bytecode(expr, key.sheet);

        // Structured references depend on table metadata. Resolve them against the current
        // workbook tables at compile time and rewrite into concrete cell/range references that
        // the bytecode lowering step understands.
        let maybe_structured_rewritten = if canonical_expr_contains_structured_refs(&expr) {
            let tables_by_sheet: Vec<Vec<Table>> = self
                .workbook
                .sheets
                .iter()
                .map(|s| s.tables.clone())
                .collect();
            Some(
                rewrite_structured_refs_for_bytecode(&expr, key.sheet, key.addr, &tables_by_sheet)
                    .ok_or(BytecodeCompileReason::IneligibleExpr)?,
            )
        } else {
            None
        };

        let expr_after_structured = maybe_structured_rewritten.as_ref().unwrap_or(&expr);

        let workbook = &self.workbook;
        let mut resolve_sheet_id = |name: &str| workbook.sheet_id(name);
        let mut expand_sheet_span =
            |start_id: SheetId, end_id: SheetId| workbook.sheet_span_ids(start_id, end_id);
        let rewritten_names = rewrite_defined_name_constants_for_bytecode(
            expr_after_structured,
            key.sheet,
            &self.workbook,
        );
        let expr_to_lower = rewritten_names.as_ref().unwrap_or(expr_after_structured);

        // External workbook references and invalid sheet prefixes can be introduced via defined
        // name inlining (or eliminated by it). Run the prefix check on the final expression shape
        // that will be lowered so bytecode eligibility and diagnostics reflect the actual lowered
        // references.
        if let Some(lower_error) = canonical_expr_depends_on_lowering_prefix_error(
            expr_to_lower,
            key.sheet,
            &self.workbook,
        ) {
            return Err(BytecodeCompileReason::LowerError(lower_error));
        }

        let mut sheet_dimensions = |sheet_id: SheetId| {
            workbook
                .sheets
                .get(sheet_id)
                .map(|s| (s.row_count, s.col_count))
        };
        let expr = bytecode::lower_canonical_expr_with_sheet_span(
            expr_to_lower,
            origin_ast,
            key.sheet,
            &mut resolve_sheet_id,
            &mut expand_sheet_span,
            &mut sheet_dimensions,
        )
        .map_err(|e| match e {
            // The lowering layer uses `Unsupported` as a catch-all for expression shapes the
            // bytecode backend doesn't implement. Surface these as `IneligibleExpr` so compile
            // reports can distinguish "missing implementation" from structural lowering errors
            // like cross-sheet references.
            bytecode::LowerError::Unsupported => BytecodeCompileReason::IneligibleExpr,
            other => BytecodeCompileReason::LowerError(other),
        })?;
        if let Some(name) = bytecode_expr_first_unsupported_function(&expr) {
            return Err(BytecodeCompileReason::UnsupportedFunction(name));
        }
        if !bytecode_expr_is_eligible(&expr) {
            return Err(BytecodeCompileReason::IneligibleExpr);
        }
        let (sheet_rows, sheet_cols) = self
            .workbook
            .sheets
            .get(key.sheet)
            .map(|sheet| {
                (
                    i32::try_from(sheet.row_count).unwrap_or(i32::MAX),
                    i32::try_from(sheet.col_count).unwrap_or(i32::MAX),
                )
            })
            .unwrap_or((0, 0));
        let mut sheet_bounds = |sheet: &bytecode::SheetId| match sheet {
            bytecode::SheetId::Local(sheet_id) => {
                let (rows, cols) = workbook
                    .sheets
                    .get(*sheet_id)
                    .map(|sheet| (sheet.row_count, sheet.col_count))
                    .unwrap_or((0, 0));
                let rows = i32::try_from(rows).unwrap_or(i32::MAX);
                let cols = i32::try_from(cols).unwrap_or(i32::MAX);
                (rows, cols)
            }
            bytecode::SheetId::External(_) => (EXCEL_MAX_ROWS_I32, EXCEL_MAX_COLS_I32),
        };
        bytecode_expr_within_grid_limits(
            &expr,
            origin,
            (sheet_rows, sheet_cols),
            &mut sheet_bounds,
        )?;
        Ok(self.bytecode_cache.get_or_compile(&expr))
    }

    fn inline_static_defined_names_for_bytecode(
        &self,
        expr: &crate::Expr,
        current_sheet: SheetId,
    ) -> crate::Expr {
        let mut visiting: HashSet<(SheetId, String)> = HashSet::new();
        let mut lexical_scopes: Vec<HashSet<String>> = Vec::new();
        self.inline_static_defined_names_for_bytecode_inner(
            expr,
            current_sheet,
            &mut visiting,
            &mut lexical_scopes,
        )
    }

    fn inline_static_defined_names_for_bytecode_inner(
        &self,
        expr: &crate::Expr,
        current_sheet: SheetId,
        visiting: &mut HashSet<(SheetId, String)>,
        lexical_scopes: &mut Vec<HashSet<String>>,
    ) -> crate::Expr {
        fn name_is_local(scopes: &[HashSet<String>], name_key: &str) -> bool {
            scopes.iter().rev().any(|scope| scope.contains(name_key))
        }

        fn bare_identifier(expr: &crate::Expr) -> Option<String> {
            match expr {
                crate::Expr::NameRef(nref) if nref.workbook.is_none() && nref.sheet.is_none() => {
                    let name_key = normalize_defined_name(&nref.name);
                    (!name_key.is_empty()).then_some(name_key)
                }
                _ => None,
            }
        }

        match expr {
            crate::Expr::FieldAccess(access) => crate::Expr::FieldAccess(crate::FieldAccessExpr {
                base: Box::new(self.inline_static_defined_names_for_bytecode_inner(
                    access.base.as_ref(),
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )),
                field: access.field.clone(),
            }),
            crate::Expr::NameRef(nref) => {
                let name_key = normalize_defined_name(&nref.name);
                if name_key.is_empty() {
                    return expr.clone();
                }

                // LET/LAMBDA lexical bindings are only visible for unqualified identifiers.
                // If a name reference is explicitly sheet-qualified (e.g. `Sheet1!X`), it should
                // bypass the local LET/LAMBDA scope and resolve as a defined name.
                if nref.workbook.is_none()
                    && nref.sheet.is_none()
                    && name_is_local(lexical_scopes, &name_key)
                {
                    return expr.clone();
                }

                self.try_inline_defined_name_ref_for_bytecode(
                    nref,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )
                .unwrap_or_else(|| expr.clone())
            }
            crate::Expr::Array(arr) => crate::Expr::Array(crate::ArrayLiteral {
                rows: arr
                    .rows
                    .iter()
                    .map(|r| {
                        r.iter()
                            .map(|e| {
                                self.inline_static_defined_names_for_bytecode_inner(
                                    e,
                                    current_sheet,
                                    visiting,
                                    lexical_scopes,
                                )
                            })
                            .collect()
                    })
                    .collect(),
            }),
            crate::Expr::FunctionCall(call) if call.name.name_upper == "LET" => {
                if call.args.len() < 3 || call.args.len() % 2 == 0 {
                    return expr.clone();
                }

                lexical_scopes.push(HashSet::new());
                let mut args = Vec::with_capacity(call.args.len());

                for pair in call.args[..call.args.len() - 1].chunks_exact(2) {
                    // LET binding identifiers are not evaluated; keep them as written.
                    let name_expr = pair[0].clone();
                    let value_expr = self.inline_static_defined_names_for_bytecode_inner(
                        &pair[1],
                        current_sheet,
                        visiting,
                        lexical_scopes,
                    );
                    args.push(name_expr.clone());
                    args.push(value_expr);

                    if let Some(name_key) = bare_identifier(&name_expr) {
                        lexical_scopes
                            .last_mut()
                            .expect("pushed scope")
                            .insert(name_key);
                    }
                }

                let body = self.inline_static_defined_names_for_bytecode_inner(
                    &call.args[call.args.len() - 1],
                    current_sheet,
                    visiting,
                    lexical_scopes,
                );
                args.push(body);

                lexical_scopes.pop();

                crate::Expr::FunctionCall(crate::FunctionCall {
                    name: call.name.clone(),
                    args,
                })
            }
            crate::Expr::FunctionCall(call) if call.name.name_upper == "LAMBDA" => {
                if call.args.is_empty() {
                    return expr.clone();
                }

                let mut scope = HashSet::new();
                for param in &call.args[..call.args.len() - 1] {
                    let Some(name_key) = bare_identifier(param) else {
                        return expr.clone();
                    };
                    if !scope.insert(name_key) {
                        return expr.clone();
                    }
                }

                lexical_scopes.push(scope);
                let mut args = Vec::with_capacity(call.args.len());
                args.extend(call.args[..call.args.len() - 1].iter().cloned());
                let body = self.inline_static_defined_names_for_bytecode_inner(
                    &call.args[call.args.len() - 1],
                    current_sheet,
                    visiting,
                    lexical_scopes,
                );
                args.push(body);
                lexical_scopes.pop();

                crate::Expr::FunctionCall(crate::FunctionCall {
                    name: call.name.clone(),
                    args,
                })
            }
            crate::Expr::FunctionCall(call) => crate::Expr::FunctionCall(crate::FunctionCall {
                name: call.name.clone(),
                args: call
                    .args
                    .iter()
                    .map(|arg| {
                        self.inline_static_defined_names_for_bytecode_inner(
                            arg,
                            current_sheet,
                            visiting,
                            lexical_scopes,
                        )
                    })
                    .collect(),
            }),
            crate::Expr::Call(call) => crate::Expr::Call(crate::CallExpr {
                callee: Box::new(self.inline_static_defined_names_for_bytecode_inner(
                    &call.callee,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )),
                args: call
                    .args
                    .iter()
                    .map(|arg| {
                        self.inline_static_defined_names_for_bytecode_inner(
                            arg,
                            current_sheet,
                            visiting,
                            lexical_scopes,
                        )
                    })
                    .collect(),
            }),
            crate::Expr::Unary(u) => crate::Expr::Unary(crate::UnaryExpr {
                op: u.op,
                expr: Box::new(self.inline_static_defined_names_for_bytecode_inner(
                    &u.expr,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )),
            }),
            crate::Expr::Postfix(p) => crate::Expr::Postfix(crate::PostfixExpr {
                op: p.op,
                expr: Box::new(self.inline_static_defined_names_for_bytecode_inner(
                    &p.expr,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )),
            }),
            crate::Expr::Binary(b) => crate::Expr::Binary(crate::BinaryExpr {
                op: b.op,
                left: Box::new(self.inline_static_defined_names_for_bytecode_inner(
                    &b.left,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )),
                right: Box::new(self.inline_static_defined_names_for_bytecode_inner(
                    &b.right,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )),
            }),
            crate::Expr::CellRef(_)
            | crate::Expr::ColRef(_)
            | crate::Expr::RowRef(_)
            | crate::Expr::StructuredRef(_)
            | crate::Expr::Number(_)
            | crate::Expr::String(_)
            | crate::Expr::Boolean(_)
            | crate::Expr::Error(_)
            | crate::Expr::Missing => expr.clone(),
        }
    }

    fn try_inline_defined_name_ref_for_bytecode(
        &self,
        nref: &crate::NameRef,
        current_sheet: SheetId,
        visiting: &mut HashSet<(SheetId, String)>,
        lexical_scopes: &mut Vec<HashSet<String>>,
    ) -> Option<crate::Expr> {
        if nref.workbook.is_some() {
            return None;
        }
        let name_key = normalize_defined_name(&nref.name);
        if name_key.is_empty() {
            return None;
        }

        let sheet_id = match nref.sheet.as_ref() {
            None => Some(current_sheet),
            Some(sheet_ref) => sheet_ref
                .as_single_sheet()
                .and_then(|name| self.workbook.sheet_id(name)),
        }?;

        self.resolve_defined_name_expr_for_bytecode(sheet_id, &name_key, visiting, lexical_scopes)
    }

    fn resolve_defined_name_expr_for_bytecode(
        &self,
        sheet_id: SheetId,
        name_key: &str,
        visiting: &mut HashSet<(SheetId, String)>,
        lexical_scopes: &mut Vec<HashSet<String>>,
    ) -> Option<crate::Expr> {
        let def = resolve_defined_name(&self.workbook, sheet_id, name_key)?;
        let visit_key = (sheet_id, name_key.to_string());
        if !visiting.insert(visit_key.clone()) {
            return None;
        }

        let result = match &def.definition {
            NameDefinition::Reference(formula) => {
                let ast = match crate::parse_formula(
                    formula,
                    crate::ParseOptions {
                        locale: crate::LocaleConfig::en_us(),
                        reference_style: crate::ReferenceStyle::A1,
                        normalize_relative_to: None,
                    },
                ) {
                    Ok(ast) => ast,
                    Err(_) => {
                        visiting.remove(&visit_key);
                        return None;
                    }
                };
                match &ast.expr {
                    crate::Expr::CellRef(_) => self.extract_static_ref_expr_for_bytecode(
                        &ast.expr,
                        sheet_id,
                        visiting,
                        lexical_scopes,
                    ),
                    crate::Expr::StructuredRef(_) => self.extract_static_ref_expr_for_bytecode(
                        &ast.expr,
                        sheet_id,
                        visiting,
                        lexical_scopes,
                    ),
                    crate::Expr::NameRef(_) => {
                        // Allow reference definitions to alias other defined names as long as they
                        // ultimately resolve to a static reference that can be lowered to bytecode.
                        self.extract_static_ref_expr_for_bytecode(
                            &ast.expr,
                            sheet_id,
                            visiting,
                            lexical_scopes,
                        )
                    }
                    crate::Expr::Postfix(p) if p.op == crate::PostfixOp::SpillRange => {
                        // Spill-range references (e.g. `A1#`) can be lowered to bytecode.
                        self.extract_static_ref_expr_for_bytecode(
                            &ast.expr,
                            sheet_id,
                            visiting,
                            lexical_scopes,
                        )
                    }
                    crate::Expr::Binary(b) if b.op == crate::BinaryOp::Union => {
                        // Multi-area reference definitions (e.g. `Sheet1!A1,Sheet1!B1`) can be
                        // lowered to a bytecode `MultiRangeRef`.
                        self.extract_static_ref_expr_for_bytecode(
                            &ast.expr,
                            sheet_id,
                            visiting,
                            lexical_scopes,
                        )
                    }
                    crate::Expr::Binary(b) if b.op == crate::BinaryOp::Intersect => {
                        // Reference intersection definitions (e.g. `A1:C3 B2:D4`) can be lowered
                        // to bytecode using reference algebra operators.
                        self.extract_static_ref_expr_for_bytecode(
                            &ast.expr,
                            sheet_id,
                            visiting,
                            lexical_scopes,
                        )
                    }
                    crate::Expr::Binary(b) if b.op == crate::BinaryOp::Range => {
                        // Reference definitions must be a direct cell/range reference.
                        if !matches!(
                            b.left.as_ref(),
                            crate::Expr::CellRef(_)
                                | crate::Expr::ColRef(_)
                                | crate::Expr::RowRef(_)
                                | crate::Expr::NameRef(_)
                        ) || !matches!(
                            b.right.as_ref(),
                            crate::Expr::CellRef(_)
                                | crate::Expr::ColRef(_)
                                | crate::Expr::RowRef(_)
                                | crate::Expr::NameRef(_)
                        ) {
                            None
                        } else {
                            self.extract_static_ref_expr_for_bytecode(
                                &ast.expr,
                                sheet_id,
                                visiting,
                                lexical_scopes,
                            )
                        }
                    }
                    _ => None,
                }
            }
            NameDefinition::Formula(formula) => {
                let ast = match crate::parse_formula(
                    formula,
                    crate::ParseOptions {
                        locale: crate::LocaleConfig::en_us(),
                        reference_style: crate::ReferenceStyle::A1,
                        normalize_relative_to: None,
                    },
                ) {
                    Ok(ast) => ast,
                    Err(_) => {
                        visiting.remove(&visit_key);
                        return None;
                    }
                };
                let inlined = self.inline_static_defined_names_for_bytecode_inner(
                    &ast.expr,
                    sheet_id,
                    visiting,
                    lexical_scopes,
                );
                Some(self.normalize_defined_name_formula_refs_for_bytecode(&inlined, sheet_id))
            }
            NameDefinition::Constant(_) => None,
        };

        visiting.remove(&visit_key);
        result
    }

    /// Defined-name formulas preserve `SheetReference::Current` so they can be evaluated relative
    /// to the sheet where the name is *used* (or explicitly sheet-qualified).
    ///
    /// When we inline a name formula into the canonical AST for bytecode compilation, we need to
    /// make that "current sheet" context explicit so later lowering treats references correctly
    /// (especially for sheet-qualified name uses like `Sheet2!MyName`).
    ///
    /// This walks an expression and fills in missing `sheet` prefixes on references using
    /// `current_sheet`, while preserving Excel's range-prefix semantics (`Sheet1!A1:B2` is parsed
    /// as `Sheet1!A1` + `B2`, where the prefix applies to both endpoints).
    fn normalize_defined_name_formula_refs_for_bytecode(
        &self,
        expr: &crate::Expr,
        current_sheet: SheetId,
    ) -> crate::Expr {
        fn ref_is_unprefixed(expr: &crate::Expr) -> bool {
            match expr {
                crate::Expr::CellRef(r) => r.workbook.is_none() && r.sheet.is_none(),
                crate::Expr::ColRef(r) => r.workbook.is_none() && r.sheet.is_none(),
                crate::Expr::RowRef(r) => r.workbook.is_none() && r.sheet.is_none(),
                _ => false,
            }
        }

        fn fill_ref_sheet(expr: &mut crate::Expr, sheet_name: &str) {
            match expr {
                crate::Expr::CellRef(r) => {
                    if r.workbook.is_none() && r.sheet.is_none() {
                        r.sheet = Some(crate::SheetRef::Sheet(sheet_name.to_string()));
                    }
                }
                crate::Expr::ColRef(r) => {
                    if r.workbook.is_none() && r.sheet.is_none() {
                        r.sheet = Some(crate::SheetRef::Sheet(sheet_name.to_string()));
                    }
                }
                crate::Expr::RowRef(r) => {
                    if r.workbook.is_none() && r.sheet.is_none() {
                        r.sheet = Some(crate::SheetRef::Sheet(sheet_name.to_string()));
                    }
                }
                _ => {}
            }
        }

        fn normalize_inner(
            expr: &crate::Expr,
            sheet_name: &str,
            fill_unprefixed: bool,
        ) -> crate::Expr {
            match expr {
                crate::Expr::CellRef(r) => {
                    let mut r = r.clone();
                    if fill_unprefixed && r.workbook.is_none() && r.sheet.is_none() {
                        r.sheet = Some(crate::SheetRef::Sheet(sheet_name.to_string()));
                    }
                    crate::Expr::CellRef(r)
                }
                crate::Expr::ColRef(r) => {
                    let mut r = r.clone();
                    if fill_unprefixed && r.workbook.is_none() && r.sheet.is_none() {
                        r.sheet = Some(crate::SheetRef::Sheet(sheet_name.to_string()));
                    }
                    crate::Expr::ColRef(r)
                }
                crate::Expr::RowRef(r) => {
                    let mut r = r.clone();
                    if fill_unprefixed && r.workbook.is_none() && r.sheet.is_none() {
                        r.sheet = Some(crate::SheetRef::Sheet(sheet_name.to_string()));
                    }
                    crate::Expr::RowRef(r)
                }
                crate::Expr::Binary(b) if b.op == crate::BinaryOp::Range => {
                    // Preserve unprefixed endpoints so range prefixes can be merged by the lowerer,
                    // then fill both endpoints when they are truly unprefixed (e.g. `A1:B2`).
                    let mut left = normalize_inner(&b.left, sheet_name, false);
                    let mut right = normalize_inner(&b.right, sheet_name, false);
                    if ref_is_unprefixed(&left) && ref_is_unprefixed(&right) {
                        fill_ref_sheet(&mut left, sheet_name);
                        fill_ref_sheet(&mut right, sheet_name);
                    }
                    crate::Expr::Binary(crate::BinaryExpr {
                        op: crate::BinaryOp::Range,
                        left: Box::new(left),
                        right: Box::new(right),
                    })
                }
                crate::Expr::Binary(b) => crate::Expr::Binary(crate::BinaryExpr {
                    op: b.op,
                    left: Box::new(normalize_inner(&b.left, sheet_name, true)),
                    right: Box::new(normalize_inner(&b.right, sheet_name, true)),
                }),
                crate::Expr::Postfix(p) => crate::Expr::Postfix(crate::PostfixExpr {
                    op: p.op,
                    expr: Box::new(normalize_inner(&p.expr, sheet_name, true)),
                }),
                crate::Expr::Unary(u) => crate::Expr::Unary(crate::UnaryExpr {
                    op: u.op,
                    expr: Box::new(normalize_inner(&u.expr, sheet_name, true)),
                }),
                crate::Expr::FunctionCall(call) => crate::Expr::FunctionCall(crate::FunctionCall {
                    name: call.name.clone(),
                    args: call
                        .args
                        .iter()
                        .map(|arg| normalize_inner(arg, sheet_name, true))
                        .collect(),
                }),
                crate::Expr::Call(call) => crate::Expr::Call(crate::CallExpr {
                    callee: Box::new(normalize_inner(&call.callee, sheet_name, true)),
                    args: call
                        .args
                        .iter()
                        .map(|arg| normalize_inner(arg, sheet_name, true))
                        .collect(),
                }),
                crate::Expr::FieldAccess(access) => {
                    crate::Expr::FieldAccess(crate::FieldAccessExpr {
                        base: Box::new(normalize_inner(&access.base, sheet_name, true)),
                        field: access.field.clone(),
                    })
                }
                crate::Expr::Array(arr) => crate::Expr::Array(crate::ArrayLiteral {
                    rows: arr
                        .rows
                        .iter()
                        .map(|row| {
                            row.iter()
                                .map(|el| normalize_inner(el, sheet_name, true))
                                .collect()
                        })
                        .collect(),
                }),
                crate::Expr::NameRef(_)
                | crate::Expr::StructuredRef(_)
                | crate::Expr::Number(_)
                | crate::Expr::String(_)
                | crate::Expr::Boolean(_)
                | crate::Expr::Error(_)
                | crate::Expr::Missing => expr.clone(),
            }
        }

        let Some(sheet_name) = self.workbook.sheet_name(current_sheet) else {
            return expr.clone();
        };
        normalize_inner(expr, sheet_name, true)
    }

    fn extract_static_ref_expr_for_bytecode(
        &self,
        expr: &crate::Expr,
        current_sheet: SheetId,
        visiting: &mut HashSet<(SheetId, String)>,
        lexical_scopes: &mut Vec<HashSet<String>>,
    ) -> Option<crate::Expr> {
        match expr {
            crate::Expr::CellRef(r) => Some(crate::Expr::CellRef(
                self.normalize_cell_ref_for_bytecode(r, current_sheet)?,
            )),
            crate::Expr::Postfix(p) if p.op == crate::PostfixOp::SpillRange => {
                let inner = self.extract_static_ref_expr_for_bytecode(
                    &p.expr,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )?;
                Some(crate::Expr::Postfix(crate::PostfixExpr {
                    op: crate::PostfixOp::SpillRange,
                    expr: Box::new(inner),
                }))
            }
            crate::Expr::Binary(b) if b.op == crate::BinaryOp::Union => {
                let left = self.extract_static_ref_expr_for_bytecode(
                    &b.left,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )?;
                let right = self.extract_static_ref_expr_for_bytecode(
                    &b.right,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )?;
                Some(crate::Expr::Binary(crate::BinaryExpr {
                    op: crate::BinaryOp::Union,
                    left: Box::new(left),
                    right: Box::new(right),
                }))
            }
            crate::Expr::Binary(b) if b.op == crate::BinaryOp::Intersect => {
                let left = self.extract_static_ref_expr_for_bytecode(
                    &b.left,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )?;
                let right = self.extract_static_ref_expr_for_bytecode(
                    &b.right,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )?;
                Some(crate::Expr::Binary(crate::BinaryExpr {
                    op: crate::BinaryOp::Intersect,
                    left: Box::new(left),
                    right: Box::new(right),
                }))
            }
            crate::Expr::Binary(b) if b.op == crate::BinaryOp::Range => {
                // Preserve unprefixed endpoints when the opposite endpoint carries a sheet prefix
                // (e.g. the parser represents `Sheet1!A1:B2` as `Sheet1!A1` + `B2`). This keeps
                // the range shape compatible with the bytecode lowerer's prefix-merging logic,
                // which applies the explicit prefix to both endpoints.
                let mut left = self.normalize_range_endpoint_for_bytecode_preserve_sheet(
                    &b.left,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )?;
                let mut right = self.normalize_range_endpoint_for_bytecode_preserve_sheet(
                    &b.right,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )?;

                let left_unprefixed = match &left {
                    crate::Expr::CellRef(r) => r.sheet.is_none(),
                    crate::Expr::ColRef(r) => r.sheet.is_none(),
                    crate::Expr::RowRef(r) => r.sheet.is_none(),
                    _ => false,
                };
                let right_unprefixed = match &right {
                    crate::Expr::CellRef(r) => r.sheet.is_none(),
                    crate::Expr::ColRef(r) => r.sheet.is_none(),
                    crate::Expr::RowRef(r) => r.sheet.is_none(),
                    _ => false,
                };
                let both_unprefixed = left_unprefixed && right_unprefixed;

                // When both endpoints are unprefixed, interpret them relative to the sheet context
                // used to resolve the defined name (which may differ from the formula's sheet when
                // the name reference is explicitly sheet-qualified).
                if both_unprefixed {
                    let sheet_ref = crate::SheetRef::Sheet(
                        self.workbook.sheet_name(current_sheet)?.to_string(),
                    );
                    match &mut left {
                        crate::Expr::CellRef(r) => r.sheet = Some(sheet_ref.clone()),
                        crate::Expr::ColRef(r) => r.sheet = Some(sheet_ref.clone()),
                        crate::Expr::RowRef(r) => r.sheet = Some(sheet_ref.clone()),
                        _ => {}
                    }
                    match &mut right {
                        crate::Expr::CellRef(r) => r.sheet = Some(sheet_ref.clone()),
                        crate::Expr::ColRef(r) => r.sheet = Some(sheet_ref.clone()),
                        crate::Expr::RowRef(r) => r.sheet = Some(sheet_ref.clone()),
                        _ => {}
                    }
                }
                Some(crate::Expr::Binary(crate::BinaryExpr {
                    op: crate::BinaryOp::Range,
                    left: Box::new(left),
                    right: Box::new(right),
                }))
            }
            crate::Expr::StructuredRef(sref) => Some(crate::Expr::StructuredRef(sref.clone())),
            crate::Expr::NameRef(nref) => self.try_inline_defined_name_ref_for_bytecode(
                nref,
                current_sheet,
                visiting,
                lexical_scopes,
            ),
            _ => None,
        }
    }

    fn normalize_range_endpoint_for_bytecode_preserve_sheet(
        &self,
        expr: &crate::Expr,
        current_sheet: SheetId,
        visiting: &mut HashSet<(SheetId, String)>,
        lexical_scopes: &mut Vec<HashSet<String>>,
    ) -> Option<crate::Expr> {
        match expr {
            crate::Expr::CellRef(r) => {
                let col = match r.col {
                    crate::Coord::A1 { index, .. } => crate::Coord::A1 { index, abs: true },
                    crate::Coord::Offset(_) => return None,
                };
                let row = match r.row {
                    crate::Coord::A1 { index, .. } => crate::Coord::A1 { index, abs: true },
                    crate::Coord::Offset(_) => return None,
                };
                Some(crate::Expr::CellRef(crate::CellRef {
                    workbook: r.workbook.clone(),
                    sheet: r.sheet.clone(),
                    col,
                    row,
                }))
            }
            crate::Expr::ColRef(r) => {
                let col = match r.col {
                    crate::Coord::A1 { index, .. } => crate::Coord::A1 { index, abs: true },
                    crate::Coord::Offset(_) => return None,
                };
                Some(crate::Expr::ColRef(crate::ColRef {
                    workbook: r.workbook.clone(),
                    sheet: r.sheet.clone(),
                    col,
                }))
            }
            crate::Expr::RowRef(r) => {
                let row = match r.row {
                    crate::Coord::A1 { index, .. } => crate::Coord::A1 { index, abs: true },
                    crate::Coord::Offset(_) => return None,
                };
                Some(crate::Expr::RowRef(crate::RowRef {
                    workbook: r.workbook.clone(),
                    sheet: r.sheet.clone(),
                    row,
                }))
            }
            crate::Expr::NameRef(nref) => {
                let resolved = self.try_inline_defined_name_ref_for_bytecode(
                    nref,
                    current_sheet,
                    visiting,
                    lexical_scopes,
                )?;
                match resolved {
                    crate::Expr::CellRef(_) | crate::Expr::ColRef(_) | crate::Expr::RowRef(_) => {
                        Some(resolved)
                    }
                    _ => None,
                }
            }
            _ => None,
        }
    }

    fn normalize_cell_ref_for_bytecode(
        &self,
        r: &crate::CellRef,
        current_sheet: SheetId,
    ) -> Option<crate::CellRef> {
        let workbook = r.workbook.clone();
        let sheet = match r.sheet.as_ref() {
            None => {
                // Unqualified references in defined-name definitions are evaluated relative to the
                // sheet where the name is *used*. When we inline for bytecode, make that context
                // explicit by filling the current sheet name.
                if workbook.is_none() {
                    Some(crate::SheetRef::Sheet(
                        self.workbook.sheet_name(current_sheet)?.to_string(),
                    ))
                } else {
                    None
                }
            }
            Some(crate::SheetRef::Sheet(name)) => Some(crate::SheetRef::Sheet(name.clone())),
            Some(crate::SheetRef::SheetRange { start, end }) => Some(crate::SheetRef::SheetRange {
                start: start.clone(),
                end: end.clone(),
            }),
        };

        let col = match r.col {
            crate::Coord::A1 { index, .. } => crate::Coord::A1 { index, abs: true },
            crate::Coord::Offset(_) => return None,
        };
        let row = match r.row {
            crate::Coord::A1 { index, .. } => crate::Coord::A1 { index, abs: true },
            crate::Coord::Offset(_) => return None,
        };

        Some(crate::CellRef {
            workbook,
            sheet,
            col,
            row,
        })
    }

    fn begin_recalc_context(&mut self) -> crate::eval::RecalcContext {
        let id = self.next_recalc_id;
        self.next_recalc_id = self.next_recalc_id.wrapping_add(1);
        let mut ctx = crate::eval::RecalcContext::new(id);
        ctx.calculation_mode = self.calc_settings.calculation_mode;
        let separators = self.value_locale.separators;
        ctx.number_locale =
            crate::value::NumberLocale::new(separators.decimal_sep, Some(separators.thousands_sep));
        ctx
    }

    fn compile_name_expr(&self, expr: &Expr<String>) -> CompiledExpr {
        let mut map = |sref: &SheetReference<String>| match sref {
            SheetReference::Current => SheetReference::Current,
            SheetReference::Sheet(name) => self
                .workbook
                .sheet_id(name)
                .map(SheetReference::Sheet)
                .unwrap_or_else(|| SheetReference::External(name.clone())),
            SheetReference::SheetRange(start, end) => {
                let start_id = self.workbook.sheet_id(start);
                let end_id = self.workbook.sheet_id(end);
                match (start_id, end_id) {
                    (Some(a), Some(b)) => SheetReference::SheetRange(a, b),
                    _ => SheetReference::External(format!("{start}:{end}")),
                }
            }
            SheetReference::External(wb) => SheetReference::External(wb.clone()),
        };
        expr.map_sheets(&mut map)
    }

    fn rewrite_defined_names_structural(
        &mut self,
        sheet_names: &HashMap<SheetId, String>,
        edit: &StructuralEdit,
    ) -> Result<(), EngineError> {
        let edit_sheet = match edit {
            StructuralEdit::InsertRows { sheet, .. }
            | StructuralEdit::DeleteRows { sheet, .. }
            | StructuralEdit::InsertCols { sheet, .. }
            | StructuralEdit::DeleteCols { sheet, .. } => sheet.as_str(),
        };

        let sheet_order_indices = build_sheet_order_indices(&self.workbook);
        let mut updates: Vec<(Option<SheetId>, String, NameDefinition, CompiledExpr)> = Vec::new();

        for (name, def) in &self.workbook.names {
            let Some((new_def, compiled)) =
                rewrite_defined_name_structural(self, def, edit_sheet, edit, &sheet_order_indices)?
            else {
                continue;
            };
            updates.push((None, name.clone(), new_def, compiled));
        }

        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            let Some(ctx_sheet) = sheet_names.get(&sheet_id) else {
                continue;
            };
            for (name, def) in &sheet.names {
                let Some((new_def, compiled)) = rewrite_defined_name_structural(
                    self,
                    def,
                    ctx_sheet,
                    edit,
                    &sheet_order_indices,
                )?
                else {
                    continue;
                };
                updates.push((Some(sheet_id), name.clone(), new_def, compiled));
            }
        }

        for (scope_sheet, name, new_def, compiled) in updates {
            match scope_sheet {
                None => {
                    if let Some(def) = self.workbook.names.get_mut(&name) {
                        def.definition = new_def;
                        def.compiled = Some(compiled);
                    }
                }
                Some(sheet_id) => {
                    if let Some(def) = self.workbook.sheets[sheet_id].names.get_mut(&name) {
                        def.definition = new_def;
                        def.compiled = Some(compiled);
                    }
                }
            }
        }

        Ok(())
    }

    fn rewrite_defined_names_range_map(
        &mut self,
        sheet_names: &HashMap<SheetId, String>,
        edit: &RangeMapEdit,
    ) -> Result<(), EngineError> {
        let sheet_order_indices = build_sheet_order_indices(&self.workbook);
        let mut updates: Vec<(Option<SheetId>, String, NameDefinition, CompiledExpr)> = Vec::new();

        for (name, def) in &self.workbook.names {
            let Some((new_def, compiled)) =
                rewrite_defined_name_range_map(self, def, &edit.sheet, edit, &sheet_order_indices)?
            else {
                continue;
            };
            updates.push((None, name.clone(), new_def, compiled));
        }

        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            let Some(ctx_sheet) = sheet_names.get(&sheet_id) else {
                continue;
            };
            for (name, def) in &sheet.names {
                let Some((new_def, compiled)) = rewrite_defined_name_range_map(
                    self,
                    def,
                    ctx_sheet,
                    edit,
                    &sheet_order_indices,
                )?
                else {
                    continue;
                };
                updates.push((Some(sheet_id), name.clone(), new_def, compiled));
            }
        }

        for (scope_sheet, name, new_def, compiled) in updates {
            match scope_sheet {
                None => {
                    if let Some(def) = self.workbook.names.get_mut(&name) {
                        def.definition = new_def;
                        def.compiled = Some(compiled);
                    }
                }
                Some(sheet_id) => {
                    if let Some(def) = self.workbook.sheets[sheet_id].names.get_mut(&name) {
                        def.definition = new_def;
                        def.compiled = Some(compiled);
                    }
                }
            }
        }

        Ok(())
    }

    fn clear_cell_name_refs(&mut self, cell: CellKey) {
        let Some(names) = self.cell_name_refs.remove(&cell) else {
            return;
        };
        for name in names {
            let should_remove = self
                .name_dependents
                .get_mut(&name)
                .map(|deps| {
                    deps.remove(&cell);
                    deps.is_empty()
                })
                .unwrap_or(false);
            if should_remove {
                self.name_dependents.remove(&name);
            }
        }
    }

    fn set_cell_name_refs(&mut self, cell: CellKey, names: HashSet<String>) {
        self.clear_cell_name_refs(cell);
        if names.is_empty() {
            return;
        }
        self.cell_name_refs.insert(cell, names.clone());
        for name in names {
            self.name_dependents.entry(name).or_default().insert(cell);
        }
    }

    fn clear_cell_external_refs(&mut self, cell: CellKey) {
        if let Some(keys) = self.cell_external_sheet_refs.remove(&cell) {
            for key in keys {
                let should_remove = self
                    .external_sheet_dependents
                    .get_mut(&key)
                    .map(|deps| {
                        deps.remove(&cell);
                        deps.is_empty()
                    })
                    .unwrap_or(false);
                if should_remove {
                    self.external_sheet_dependents.remove(&key);
                }
            }
        }

        if let Some(workbooks) = self.cell_external_workbook_refs.remove(&cell) {
            for workbook in workbooks {
                let should_remove = self
                    .external_workbook_dependents
                    .get_mut(&workbook)
                    .map(|deps| {
                        deps.remove(&cell);
                        deps.is_empty()
                    })
                    .unwrap_or(false);
                if should_remove {
                    self.external_workbook_dependents.remove(&workbook);
                }
            }
        }
    }

    fn set_cell_external_refs(
        &mut self,
        cell: CellKey,
        sheet_keys: HashSet<String>,
        workbook_keys: HashSet<String>,
    ) {
        self.clear_cell_external_refs(cell);

        if !sheet_keys.is_empty() {
            self.cell_external_sheet_refs
                .insert(cell, sheet_keys.clone());
            for key in sheet_keys {
                self.external_sheet_dependents
                    .entry(key)
                    .or_default()
                    .insert(cell);
            }
        }

        if !workbook_keys.is_empty() {
            self.cell_external_workbook_refs
                .insert(cell, workbook_keys.clone());
            for workbook in workbook_keys {
                self.external_workbook_dependents
                    .entry(workbook)
                    .or_default()
                    .insert(cell);
            }
        }
    }

    fn clear_cell_dynamic_external_precedents(&mut self, cell: CellKey) {
        self.cell_dynamic_external_precedents.remove(&cell);

        if let Some(keys) = self.cell_dynamic_external_sheet_refs.remove(&cell) {
            for key in keys {
                let should_remove = self
                    .dynamic_external_sheet_dependents
                    .get_mut(&key)
                    .map(|deps| {
                        deps.remove(&cell);
                        deps.is_empty()
                    })
                    .unwrap_or(false);
                if should_remove {
                    self.dynamic_external_sheet_dependents.remove(&key);
                }
            }
        }

        if let Some(workbooks) = self.cell_dynamic_external_workbook_refs.remove(&cell) {
            for workbook in workbooks {
                let should_remove = self
                    .dynamic_external_workbook_dependents
                    .get_mut(&workbook)
                    .map(|deps| {
                        deps.remove(&cell);
                        deps.is_empty()
                    })
                    .unwrap_or(false);
                if should_remove {
                    self.dynamic_external_workbook_dependents.remove(&workbook);
                }
            }
        }
    }

    fn set_cell_dynamic_external_precedents(
        &mut self,
        cell: CellKey,
        precedents: HashSet<crate::functions::Reference>,
    ) {
        self.clear_cell_dynamic_external_precedents(cell);
        if precedents.is_empty() {
            return;
        }

        let mut sheet_keys: HashSet<String> = HashSet::new();
        let mut workbook_keys: HashSet<String> = HashSet::new();
        for reference in &precedents {
            let crate::functions::SheetId::External(key) = &reference.sheet_id else {
                continue;
            };

            if let Some((workbook, _sheet)) = crate::eval::split_external_sheet_key(key) {
                workbook_keys.insert(workbook.to_string());
            }

            if crate::eval::is_valid_external_sheet_key(key) {
                sheet_keys.insert(key.clone());
            } else if crate::eval::split_external_sheet_span_key(key).is_some() {
                if let Some(expanded) =
                    expand_external_sheet_span_key(key, self.external_value_provider.as_deref())
                {
                    sheet_keys.extend(expanded);
                }
            }
        }

        self.cell_dynamic_external_precedents
            .insert(cell, precedents);

        if !sheet_keys.is_empty() {
            self.cell_dynamic_external_sheet_refs
                .insert(cell, sheet_keys.clone());
            for key in sheet_keys {
                self.dynamic_external_sheet_dependents
                    .entry(key)
                    .or_default()
                    .insert(cell);
            }
        }

        if !workbook_keys.is_empty() {
            self.cell_dynamic_external_workbook_refs
                .insert(cell, workbook_keys.clone());
            for workbook in workbook_keys {
                self.dynamic_external_workbook_dependents
                    .entry(workbook)
                    .or_default()
                    .insert(cell);
            }
        }
    }

    fn refresh_cells_after_name_change(&mut self, name: &str) {
        let Some(cells) = self.name_dependents.get(name).cloned() else {
            return;
        };

        let tables_by_sheet: Vec<Vec<Table>> = self
            .workbook
            .sheets
            .iter()
            .map(|s| s.tables.clone())
            .collect();

        for key in cells {
            let Some((ast, formula)) = self.workbook.get_cell(key).and_then(|c| {
                Some((
                    c.compiled.as_ref()?.ast().clone(),
                    c.formula.as_ref()?.clone(),
                ))
            }) else {
                continue;
            };

            let cell_id = cell_id_from_key(key);
            self.clear_cell_dynamic_external_precedents(key);

            let (names, volatile, thread_safe, dynamic_deps, origin_deps) = analyze_expr_flags(
                &ast,
                key,
                &tables_by_sheet,
                &self.workbook,
                self.external_refs_volatile,
            );
            self.set_cell_name_refs(key, names);
            let (external_sheets, external_workbooks) = analyze_external_dependencies(
                &ast,
                key,
                &self.workbook,
                self.external_value_provider.as_deref(),
            );
            self.set_cell_external_refs(key, external_sheets, external_workbooks);

            let calc_precedents =
                analyze_calc_precedents(&ast, key, &tables_by_sheet, &self.workbook, &self.spills);
            let mut calc_vec: Vec<Precedent> = calc_precedents.into_iter().collect();
            calc_vec.sort_by_key(|p| match p {
                Precedent::Cell(c) => (0u8, c.sheet_id, c.cell.row, c.cell.col, 0u32, 0u32),
                Precedent::Range(r) => (
                    1u8,
                    r.sheet_id,
                    r.range.start.row,
                    r.range.start.col,
                    r.range.end.row,
                    r.range.end.col,
                ),
            });
            let deps = CellDeps::new(calc_vec).volatile(volatile);
            self.calc_graph.update_cell_dependencies(cell_id, deps);

            // Name definition changes can affect bytecode eligibility (and for constant inlining,
            // even the literal value baked into the program), so rebuild the compiled variant.
            let (compiled_formula, bytecode_compile_reason) = {
                let origin = crate::CellAddr::new(key.addr.row, key.addr.col);
                let parsed = crate::parse_formula(
                    &formula,
                    crate::ParseOptions {
                        locale: crate::LocaleConfig::en_us(),
                        reference_style: crate::ReferenceStyle::A1,
                        normalize_relative_to: Some(origin),
                    },
                );

                match parsed {
                    Ok(parsed) => match self.try_compile_bytecode(
                        &parsed.expr,
                        key,
                        thread_safe,
                        dynamic_deps,
                    ) {
                        Ok(program) => (
                            CompiledFormula::Bytecode(BytecodeFormula {
                                ast: ast.clone(),
                                program,
                                sheet_dims_generation: self.sheet_dims_generation,
                            }),
                            None,
                        ),
                        Err(reason) => (CompiledFormula::Ast(ast.clone()), Some(reason)),
                    },
                    Err(_) => (
                        CompiledFormula::Ast(ast.clone()),
                        Some(BytecodeCompileReason::IneligibleExpr),
                    ),
                }
            };

            {
                let cell = self.workbook.get_or_create_cell_mut(key);
                cell.compiled = Some(compiled_formula);
                cell.bytecode_compile_reason = bytecode_compile_reason;
                cell.volatile = volatile;
                cell.thread_safe = thread_safe;
                cell.dynamic_deps = dynamic_deps;
            }

            if let Some(sheet_state) = self.workbook.sheets.get_mut(key.sheet) {
                if origin_deps {
                    sheet_state.origin_dependents.insert(key.addr);
                } else {
                    sheet_state.origin_dependents.remove(&key.addr);
                }
            }

            self.mark_dirty_including_self_with_reasons(key);
            self.calc_graph.mark_dirty(cell_id);
        }
    }
    fn rebuild_graph(&mut self) -> Result<(), EngineError> {
        let sheet_names = sheet_names_by_id(&self.workbook);
        let mut formulas: Vec<(SheetId, String, CellAddr, Arc<str>, Option<String>)> = Vec::new();
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            let Some(sheet_name) = sheet_names.get(&sheet_id).cloned() else {
                continue;
            };
            for (addr, cell) in &sheet.cells {
                if let Some(formula) = &cell.formula {
                    formulas.push((
                        sheet_id,
                        sheet_name.clone(),
                        *addr,
                        formula.to_string(),
                        cell.phonetic.clone(),
                    ));
                }
            }
        }

        self.calc_graph = CalcGraph::new();
        self.name_dependents.clear();
        self.cell_name_refs.clear();
        self.external_sheet_dependents.clear();
        self.cell_external_sheet_refs.clear();
        self.external_workbook_dependents.clear();
        self.cell_external_workbook_refs.clear();
        self.cell_dynamic_external_precedents.clear();
        self.dynamic_external_sheet_dependents.clear();
        self.cell_dynamic_external_sheet_refs.clear();
        self.dynamic_external_workbook_dependents.clear();
        self.cell_dynamic_external_workbook_refs.clear();
        self.dirty.clear();
        self.dirty_reasons.clear();
        self.spills = SpillState::default();
        for sheet in &mut self.workbook.sheets {
            sheet.origin_dependents.clear();
        }

        for (sheet_id, sheet_name, addr, formula, phonetic) in formulas {
            let addr_a1 = cell_addr_to_a1(addr);
            self.set_cell_formula(&sheet_name, &addr_a1, &formula)?;
            if let Some(phonetic) = phonetic {
                // Rebuilding the dependency graph recompiles formulas but should not clear
                // cell-level phonetic metadata (used by `PHONETIC()`).
                if let Some(cell) = self
                    .workbook
                    .sheets
                    .get_mut(sheet_id)
                    .and_then(|sheet| sheet.cells.get_mut(&addr))
                {
                    cell.phonetic = Some(phonetic);
                }
            }
        }
        Ok(())
    }

    /// Returns whether a cell is currently marked dirty (needs recalculation).
    pub fn is_dirty(&self, sheet: &str, addr: &str) -> bool {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return false;
        };
        let Ok(addr) = parse_a1(addr) else {
            return false;
        };
        self.dirty.contains(&CellKey {
            sheet: sheet_id,
            addr,
        })
    }

    /// Direct precedents (cells and ranges referenced by the formula in `cell`).
    ///
    /// Note: external-workbook 3D spans like `[Book.xlsx]Sheet1:Sheet3!A1` are expanded into
    /// per-sheet precedents when the engine has an [`ExternalValueProvider`] that supplies
    /// [`ExternalValueProvider::sheet_order`]. If sheet order is unavailable, 3D spans are omitted
    /// from this auditing API (matching evaluation, which returns `#REF!`).
    pub fn precedents(&self, sheet: &str, addr: &str) -> Result<Vec<PrecedentNode>, EngineError> {
        self.precedents_impl(sheet, addr, false)
    }

    /// Transitive precedents (all precedents that can influence `cell`).
    pub fn precedents_transitive(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<PrecedentNode>, EngineError> {
        self.precedents_impl(sheet, addr, true)
    }

    /// Returns a cell-level view of `precedents`, expanding any ranges until `limit` cells have been
    /// produced.
    ///
    /// This is intended for UI tracing/highlighting. The returned list is deterministically ordered
    /// by workbook tab order, then `(row, col)`.
    pub fn precedents_expanded(
        &self,
        sheet: &str,
        addr: &str,
        limit: usize,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        let nodes = self.precedents(sheet, addr)?;
        Ok(expand_nodes_to_cells(&nodes, limit, &self.workbook))
    }

    /// Direct dependents (cells whose formulas reference `cell`).
    pub fn dependents(&self, sheet: &str, addr: &str) -> Result<Vec<PrecedentNode>, EngineError> {
        self.dependents_impl(sheet, addr, false)
    }

    /// Transitive dependents (all downstream cells that are affected by `cell`).
    pub fn dependents_transitive(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<Vec<PrecedentNode>, EngineError> {
        self.dependents_impl(sheet, addr, true)
    }

    /// Returns a cell-level view of `dependents`, expanding any range nodes until `limit` cells have
    /// been produced.
    #[must_use]
    pub fn dependents_expanded(
        &self,
        sheet: &str,
        addr: &str,
        limit: usize,
    ) -> Result<Vec<(SheetId, CellAddr)>, EngineError> {
        let nodes = self.dependents(sheet, addr)?;
        Ok(expand_nodes_to_cells(&nodes, limit, &self.workbook))
    }

    /// Returns a dependency path explaining why `cell` is currently dirty.
    ///
    /// The returned vector is ordered from the root cause (usually an edited input cell) to the
    /// provided `cell`.
    pub fn dirty_dependency_path(&self, sheet: &str, addr: &str) -> Option<Vec<PrecedentNode>> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        if !self.dirty.contains(&key) {
            return None;
        }

        let mut path = vec![PrecedentNode::Cell {
            sheet: key.sheet,
            addr: key.addr,
        }];
        let mut current = key;
        let mut guard = 0usize;
        while let Some(reason) = self.dirty_reasons.get(&current).cloned() {
            match reason {
                DirtyReason::Cell(prev) => {
                    path.push(PrecedentNode::Cell {
                        sheet: prev.sheet,
                        addr: prev.addr,
                    });
                    current = prev;
                }
                DirtyReason::ViaRange { from, range } => {
                    path.push(range);
                    path.push(PrecedentNode::Cell {
                        sheet: from.sheet,
                        addr: from.addr,
                    });
                    current = from;
                }
            }
            guard += 1;
            if guard > 10_000 {
                break;
            }
        }
        path.reverse();
        Some(path)
    }

    /// Deterministically evaluates a cell's formula while capturing a per-node trace.
    ///
    /// This is intended for on-demand debugging and does **not** mutate engine state.
    pub fn debug_evaluate(
        &self,
        sheet: &str,
        addr: &str,
    ) -> Result<crate::debug::DebugEvaluation, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Err(EngineError::Parse(FormulaParseError::UnexpectedToken(
                format!("unknown sheet '{sheet}'"),
            )));
        };
        let addr = parse_a1(addr)?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let cell = self.workbook.get_cell(key);
        let Some(formula) = cell.and_then(|c| c.formula.as_deref()) else {
            return Err(EngineError::Parse(FormulaParseError::UnexpectedToken(
                "cell has no formula".to_string(),
            )));
        };

        let snapshot = Snapshot::from_workbook(
            &self.workbook,
            &self.spills,
            self.external_value_provider.clone(),
            self.external_data_provider.clone(),
            self.info.clone(),
            self.pivot_registry.clone(),
        );
        let ctx = crate::eval::EvalContext {
            current_sheet: sheet_id,
            current_cell: addr,
        };

        // Parse with spans, compile sheet references without mutating the workbook,
        // then evaluate with tracing.
        let parsed = crate::debug::parse_spanned_formula(formula)?;
        let mut map = |sref: &SheetReference<String>| match sref {
            SheetReference::Current => SheetReference::Current,
            SheetReference::Sheet(name) => self
                .workbook
                .sheet_id(name)
                .map(SheetReference::Sheet)
                .unwrap_or_else(|| SheetReference::External(name.clone())),
            SheetReference::SheetRange(start, end) => {
                let start_id = self.workbook.sheet_id(start);
                let end_id = self.workbook.sheet_id(end);
                match (start_id, end_id) {
                    (Some(a), Some(b)) => SheetReference::SheetRange(a, b),
                    _ => SheetReference::External(format!("{start}:{end}")),
                }
            }
            SheetReference::External(wb) => SheetReference::External(wb.clone()),
        };
        let compiled = parsed.map_sheets(&mut map);

        let mut recalc_ctx = crate::eval::RecalcContext::new(0);
        let separators = self.value_locale.separators;
        recalc_ctx.number_locale =
            crate::value::NumberLocale::new(separators.decimal_sep, Some(separators.thousands_sep));
        recalc_ctx.calculation_mode = self.calc_settings.calculation_mode;
        let (value, trace) = crate::debug::evaluate_with_trace(
            &snapshot,
            ctx,
            &recalc_ctx,
            self.date_system,
            self.value_locale,
            &compiled,
        );

        Ok(crate::debug::DebugEvaluation {
            formula: formula.to_string(),
            value,
            trace,
        })
    }

    fn dynamic_external_precedent_nodes(&self, cell: CellKey) -> Vec<PrecedentNode> {
        let Some(precedents) = self.cell_dynamic_external_precedents.get(&cell) else {
            return Vec::new();
        };

        let mut out = Vec::new();
        for reference in precedents {
            let crate::functions::SheetId::External(key) = &reference.sheet_id else {
                continue;
            };
            if crate::eval::split_external_sheet_key(key).is_none() {
                continue;
            }

            let start = clamp_addr_to_excel_dimensions(reference.start);
            let end = clamp_addr_to_excel_dimensions(reference.end);
            let (start, end) = normalize_range(start, end);
            if start == end {
                out.push(PrecedentNode::ExternalCell {
                    sheet: key.clone(),
                    addr: start,
                });
            } else {
                out.push(PrecedentNode::ExternalRange {
                    sheet: key.clone(),
                    start,
                    end,
                });
            }
        }
        out
    }

    fn precedents_impl(
        &self,
        sheet: &str,
        addr: &str,
        transitive: bool,
    ) -> Result<Vec<PrecedentNode>, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(Vec::new());
        };
        let addr = parse_a1(addr)?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        if transitive {
            return Ok(self.precedents_transitive_nodes(key));
        }

        let cell_id = cell_id_from_key(key);
        let mut out: Vec<PrecedentNode> = self
            .calc_graph
            .precedents_of(cell_id)
            .into_iter()
            .map(|precedent| precedent_to_node(precedent, &self.workbook))
            .collect();
        if let Some(cell) = self.workbook.get_cell(key) {
            if let Some(compiled) = cell.compiled.as_ref() {
                out.extend(analyze_external_precedents(
                    compiled.ast(),
                    key,
                    &self.workbook,
                    self.external_value_provider.as_deref(),
                ));
            }
        }
        out.extend(self.dynamic_external_precedent_nodes(key));
        sort_and_dedup_nodes(&mut out, &self.workbook);
        Ok(out)
    }

    fn dependents_impl(
        &self,
        sheet: &str,
        addr: &str,
        transitive: bool,
    ) -> Result<Vec<PrecedentNode>, EngineError> {
        let Some(sheet_id) = self.workbook.sheet_id(sheet) else {
            return Ok(Vec::new());
        };
        let addr = parse_a1(addr)?;
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        if transitive {
            return Ok(self.dependents_transitive_nodes(key));
        }

        let cell_id = cell_id_from_key(key);
        let mut out: Vec<PrecedentNode> = self
            .calc_graph
            .direct_dependents(cell_id)
            .into_iter()
            .map(cell_key_from_id)
            .map(|key| PrecedentNode::Cell {
                sheet: key.sheet,
                addr: key.addr,
            })
            .collect();
        sort_and_dedup_nodes(&mut out, &self.workbook);
        Ok(out)
    }

    fn mark_dirty_including_self_with_reasons(&mut self, from: CellKey) {
        self.dirty.insert(from);
        self.dirty_reasons.remove(&from);
        self.mark_dirty_dependents_with_reasons(from);
    }

    fn mark_dirty_dependents_with_reasons(&mut self, from: CellKey) {
        let mut queue: VecDeque<(DirtyReason, CellKey)> = VecDeque::new();

        let from_id = cell_id_from_key(from);
        for edge in self.calc_graph.dependents_of(from_id) {
            let dep = cell_key_from_id(edge.dependent);
            let reason = match edge.kind {
                crate::graph::DependentEdgeKind::DirectCell => DirtyReason::Cell(from),
                crate::graph::DependentEdgeKind::Range(range) => DirtyReason::ViaRange {
                    from,
                    range: sheet_range_to_node(range, &self.workbook),
                },
            };
            queue.push_back((reason, dep));
        }

        while let Some((reason, cell)) = queue.pop_front() {
            if !self.dirty.insert(cell) {
                continue;
            }
            self.dirty_reasons.entry(cell).or_insert(reason);

            let cell_id = cell_id_from_key(cell);
            for edge in self.calc_graph.dependents_of(cell_id) {
                let dep = cell_key_from_id(edge.dependent);
                let next_reason = match edge.kind {
                    crate::graph::DependentEdgeKind::DirectCell => DirtyReason::Cell(cell),
                    crate::graph::DependentEdgeKind::Range(range) => DirtyReason::ViaRange {
                        from: cell,
                        range: sheet_range_to_node(range, &self.workbook),
                    },
                };
                queue.push_back((next_reason, dep));
            }
        }
    }

    fn dependents_transitive_nodes(&self, start: CellKey) -> Vec<PrecedentNode> {
        let mut visited: HashSet<CellKey> = HashSet::new();
        let mut out: Vec<CellKey> = Vec::new();
        let mut queue: VecDeque<CellKey> = VecDeque::new();

        visited.insert(start);
        queue.push_back(start);

        while let Some(cell) = queue.pop_front() {
            let cell_id = cell_id_from_key(cell);
            for dep_id in self.calc_graph.direct_dependents(cell_id) {
                let dep = cell_key_from_id(dep_id);
                if visited.insert(dep) {
                    out.push(dep);
                    queue.push_back(dep);
                }
            }
        }

        let mut out: Vec<PrecedentNode> = out
            .into_iter()
            .map(|k| PrecedentNode::Cell {
                sheet: k.sheet,
                addr: k.addr,
            })
            .collect();
        sort_and_dedup_nodes(&mut out, &self.workbook);
        out
    }

    fn precedents_transitive_nodes(&self, start: CellKey) -> Vec<PrecedentNode> {
        let start_node = PrecedentNode::Cell {
            sheet: start.sheet,
            addr: start.addr,
        };

        let mut visited: HashSet<PrecedentNode> = HashSet::new();
        let mut out: Vec<PrecedentNode> = Vec::new();
        let mut queue: VecDeque<PrecedentNode> = VecDeque::new();

        visited.insert(start_node.clone());
        queue.push_back(start_node);

        while let Some(node) = queue.pop_front() {
            let neighbors: Vec<PrecedentNode> = match node {
                PrecedentNode::Cell { sheet, addr } => {
                    let key = CellKey { sheet, addr };
                    let cell_id = cell_id_from_key(key);
                    let mut neighbors: Vec<PrecedentNode> = self
                        .calc_graph
                        .precedents_of(cell_id)
                        .into_iter()
                        .map(|precedent| precedent_to_node(precedent, &self.workbook))
                        .collect();
                    if let Some(cell) = self.workbook.get_cell(key) {
                        if let Some(compiled) = cell.compiled.as_ref() {
                            neighbors.extend(analyze_external_precedents(
                                compiled.ast(),
                                key,
                                &self.workbook,
                                self.external_value_provider.as_deref(),
                            ));
                        }
                    }
                    neighbors.extend(self.dynamic_external_precedent_nodes(key));
                    neighbors
                }
                PrecedentNode::Range { sheet, start, end } => {
                    let range = Range::new(cell_ref_from_addr(start), cell_ref_from_addr(end));
                    let sheet_range = SheetRange::new(sheet_id_for_graph(sheet), range);
                    self.calc_graph
                        .formula_cells_in_range(sheet_range)
                        .into_iter()
                        .map(|id| {
                            let key = cell_key_from_id(id);
                            PrecedentNode::Cell {
                                sheet: key.sheet,
                                addr: key.addr,
                            }
                        })
                        .collect()
                }
                PrecedentNode::SpillRange { sheet, origin, .. } => vec![PrecedentNode::Cell {
                    sheet,
                    addr: origin,
                }],
                PrecedentNode::ExternalCell { .. } | PrecedentNode::ExternalRange { .. } => {
                    Vec::new()
                }
            };

            for n in neighbors {
                if visited.insert(n.clone()) {
                    out.push(n.clone());
                    queue.push_back(n);
                }
            }
        }

        sort_and_dedup_nodes(&mut out, &self.workbook);
        out
    }

    fn sync_dirty_from_calc_graph(&mut self) {
        for id in self.calc_graph.dirty_cells() {
            self.dirty.insert(cell_key_from_id(id));
        }
    }
}

impl PivotRefreshContext for Engine {
    fn read_cell(&mut self, sheet: &str, addr: &str) -> Value {
        self.get_cell_value(sheet, addr)
    }

    fn read_cell_number_format(&self, sheet: &str, addr: &str) -> Option<String> {
        let sheet_id = self.workbook.sheet_id(sheet)?;
        let addr = parse_a1(addr).ok()?;
        self.number_format_pattern_for_rounding(CellKey {
            sheet: sheet_id,
            addr,
        })
        .map(|s| s.to_string())
    }

    fn date_system(&self) -> ExcelDateSystem {
        Engine::date_system(self)
    }

    fn intern_style(&mut self, style: Style) -> u32 {
        Engine::intern_style(self, style)
    }

    fn set_cell_style_id(
        &mut self,
        sheet: &str,
        addr: &str,
        style_id: u32,
    ) -> Result<(), EngineError> {
        Engine::set_cell_style_id(self, sheet, addr, style_id)
    }

    fn set_cell_style_ids(
        &mut self,
        sheet: &str,
        writes: &[(formula_model::CellRef, u32)],
    ) -> Result<(), EngineError> {
        Engine::set_cell_style_ids(self, sheet, writes)
    }

    fn write_cell(&mut self, sheet: &str, addr: &str, value: Value) -> Result<(), EngineError> {
        self.set_cell_value(sheet, addr, value)
    }

    fn clear_cell(&mut self, sheet: &str, addr: &str) -> Result<(), EngineError> {
        Engine::clear_cell(self, sheet, addr)
    }

    fn resolve_table(&mut self, table_id: u32) -> Option<(String, Range)> {
        for (sheet_id, sheet) in self.workbook.sheets.iter().enumerate() {
            if let Some(table) = sheet.tables.iter().find(|t| t.id == table_id) {
                let name = self
                    .workbook
                    .sheet_key_name(sheet_id)
                    .map(|n| n.to_string())
                    .unwrap_or_else(|| format!("Sheet{sheet_id}"));
                return Some((name, table.range));
            }
        }
        None
    }

    fn pivot_cache_from_range(
        &mut self,
        sheet: &str,
        range: Range,
    ) -> Result<crate::pivot::PivotCache, crate::pivot::PivotError> {
        Engine::pivot_cache_from_range(self, sheet, range)
    }

    fn clear_range(&mut self, sheet: &str, range: Range) -> Result<(), EngineError> {
        Engine::clear_range(self, sheet, range, false)
    }

    fn set_range_values(
        &mut self,
        sheet: &str,
        range: Range,
        values: &[Vec<Value>],
    ) -> Result<(), EngineError> {
        Engine::set_range_values(self, sheet, range, values, false)
    }

    fn register_pivot_table(
        &mut self,
        sheet: &str,
        destination: Range,
        pivot: crate::pivot::PivotTable,
    ) -> Result<(), crate::pivot_registry::PivotRegistryError> {
        Engine::register_pivot_table(self, sheet, destination, pivot)
    }

    fn unregister_pivot_table(&mut self, pivot_id: &str) {
        self.pivot_registry.unregister(pivot_id);
    }
}

fn sheet_names_by_id(workbook: &Workbook) -> HashMap<SheetId, String> {
    workbook
        .sheet_keys
        .iter()
        .enumerate()
        .filter_map(|(id, name)| name.as_ref().map(|name| (id, name.clone())))
        .collect()
}

fn rewrite_table_names_in_defined_names(
    names: &mut HashMap<String, DefinedName>,
    renames: &[(String, String)],
) {
    for def in names.values_mut() {
        let mut changed = false;
        match &mut def.definition {
            NameDefinition::Reference(formula) | NameDefinition::Formula(formula) => {
                let rewritten = rewrite_table_names_in_formula(formula, renames);
                if rewritten != *formula {
                    *formula = rewritten;
                    changed = true;
                }
            }
            NameDefinition::Constant(_) => {}
        }

        if changed {
            if let Some(compiled) = def.compiled.as_mut() {
                rewrite_table_names_in_compiled_expr(compiled, renames);
            }
        }
    }
}

fn rewrite_table_names_in_compiled_formula(
    compiled: &mut CompiledFormula,
    renames: &[(String, String)],
) {
    match compiled {
        CompiledFormula::Ast(expr) => rewrite_table_names_in_compiled_expr(expr, renames),
        CompiledFormula::Bytecode(bc) => rewrite_table_names_in_compiled_expr(&mut bc.ast, renames),
    }
}

fn rewrite_table_names_in_compiled_expr(expr: &mut CompiledExpr, renames: &[(String, String)]) {
    match expr {
        Expr::ArrayLiteral { values, .. } => {
            // `Arc<[T]>` does not expose mutable iteration over the slice, so clone to a vec and
            // re-wrap if any entries were rewritten.
            let mut out: Vec<Expr<usize>> = values.iter().cloned().collect();
            let mut changed = false;
            for el in &mut out {
                let before = el.clone();
                rewrite_table_names_in_compiled_expr(el, renames);
                changed |= *el != before;
            }
            if changed {
                *values = Arc::from(out);
            }
        }
        Expr::NameRef(nref) => {
            for (old, new) in renames {
                if nref.name.eq_ignore_ascii_case(old) {
                    nref.name = new.clone();
                    break;
                }
            }
        }
        Expr::StructuredRef(sref_expr) => {
            let Some(name) = sref_expr.sref.table_name.clone() else {
                return;
            };
            for (old, new) in renames {
                if name.eq_ignore_ascii_case(old) {
                    sref_expr.sref.table_name = Some(new.clone());
                    break;
                }
            }
        }
        Expr::FieldAccess { base, .. } => rewrite_table_names_in_compiled_expr(base, renames),
        Expr::Unary { expr, .. }
        | Expr::Postfix { expr, .. }
        | Expr::ImplicitIntersection(expr)
        | Expr::SpillRange(expr) => rewrite_table_names_in_compiled_expr(expr, renames),
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            rewrite_table_names_in_compiled_expr(left, renames);
            rewrite_table_names_in_compiled_expr(right, renames);
        }
        Expr::FunctionCall { args, .. } => {
            for arg in args {
                rewrite_table_names_in_compiled_expr(arg, renames);
            }
        }
        Expr::Call { callee, args } => {
            rewrite_table_names_in_compiled_expr(callee, renames);
            for arg in args {
                rewrite_table_names_in_compiled_expr(arg, renames);
            }
        }
        Expr::Number(_)
        | Expr::Text(_)
        | Expr::Bool(_)
        | Expr::Blank
        | Expr::Error(_)
        | Expr::CellRef(_)
        | Expr::RangeRef(_) => {}
    }
}

fn cell_ref_from_addr(addr: CellAddr) -> CellRef {
    CellRef::new(addr.row, addr.col)
}

fn cell_addr_from_cell_ref(cell: CellRef) -> CellAddr {
    CellAddr {
        row: cell.row,
        col: cell.col,
    }
}

fn cell_addr_to_a1(addr: CellAddr) -> String {
    addr.to_a1()
}

fn ranges_overlap(a: Range, b: Range) -> bool {
    !(a.end.row < b.start.row
        || a.start.row > b.end.row
        || a.end.col < b.start.col
        || a.start.col > b.end.col)
}

fn shift_rows(sheet: &mut Sheet, row: u32, count: u32, insert: bool) {
    let del_end = row.saturating_add(count.saturating_sub(1));
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if insert {
            if addr.row >= row {
                new_cells.insert(
                    CellAddr {
                        row: addr.row + count,
                        col: addr.col,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
            continue;
        }

        if addr.row < row {
            new_cells.insert(addr, cell);
        } else if addr.row > del_end {
            new_cells.insert(
                CellAddr {
                    row: addr.row - count,
                    col: addr.col,
                },
                cell,
            );
        }
    }
    sheet.cells = new_cells;

    // Shift row-level metadata alongside the cells for full-row edits.
    let mut new_props = BTreeMap::new();
    for (r, props) in std::mem::take(&mut sheet.row_properties) {
        if insert {
            if r >= row {
                new_props.insert(r.saturating_add(count), props);
            } else {
                new_props.insert(r, props);
            }
            continue;
        }

        if r < row {
            new_props.insert(r, props);
        } else if r > del_end {
            new_props.insert(r.saturating_sub(count), props);
        }
    }
    sheet.row_properties = new_props;
}

fn shift_cols(sheet: &mut Sheet, col: u32, count: u32, insert: bool) {
    let del_end = col.saturating_add(count.saturating_sub(1));
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if insert {
            if addr.col >= col {
                new_cells.insert(
                    CellAddr {
                        row: addr.row,
                        col: addr.col + count,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
            continue;
        }

        if addr.col < col {
            new_cells.insert(addr, cell);
        } else if addr.col > del_end {
            new_cells.insert(
                CellAddr {
                    row: addr.row,
                    col: addr.col - count,
                },
                cell,
            );
        }
    }
    sheet.cells = new_cells;

    // Shift column-level metadata alongside the cells for full-column edits.
    let mut new_props = BTreeMap::new();
    for (c, props) in std::mem::take(&mut sheet.col_properties) {
        if insert {
            if c >= col {
                new_props.insert(c.saturating_add(count), props);
            } else {
                new_props.insert(c, props);
            }
            continue;
        }

        if c < col {
            new_props.insert(c, props);
        } else if c > del_end {
            new_props.insert(c.saturating_sub(count), props);
        }
    }
    sheet.col_properties = new_props;
}

/// Ensure `table.columns.len()` matches `table.range.width()`.
///
/// When a table needs additional columns, we generate Excel-like default names (`Column1`,
/// `Column2`, ...) by picking the lowest positive integer that does not collide with an existing
/// column name (case-insensitive). This mirrors the behavior of
/// [`formula_model::table::Table::set_range`].
fn normalize_table_columns(table: &mut Table) {
    let target = table.range.width() as usize;
    let current = table.columns.len();
    if current == target {
        return;
    }
    if target < current {
        table.columns.truncate(target);
        return;
    }

    let mut used_names: HashSet<String> = table
        .columns
        .iter()
        .map(|c| c.name.to_ascii_lowercase())
        .collect();
    let mut next_id = table.columns.iter().map(|c| c.id).max().unwrap_or(0) + 1;
    let mut next_default_num: u32 = 1;

    for _ in current..target {
        let name = loop {
            let candidate = format!("Column{next_default_num}");
            next_default_num += 1;
            if used_names.insert(candidate.to_ascii_lowercase()) {
                break candidate;
            }
        };
        table.columns.push(TableColumn {
            id: next_id,
            name,
            formula: None,
            totals_formula: None,
        });
        next_id += 1;
    }
}

fn insert_default_table_columns(table: &mut Table, insert_idx: usize, count: u32) {
    if count == 0 {
        return;
    }

    let mut used_names: HashSet<String> = table
        .columns
        .iter()
        .map(|c| c.name.to_ascii_lowercase())
        .collect();
    let mut next_id = table.columns.iter().map(|c| c.id).max().unwrap_or(0) + 1;
    let mut next_default_num: u32 = 1;

    let mut inserted: Vec<TableColumn> = Vec::with_capacity(count as usize);
    for _ in 0..count {
        let name = loop {
            let candidate = format!("Column{next_default_num}");
            next_default_num += 1;
            if used_names.insert(candidate.to_ascii_lowercase()) {
                break candidate;
            }
        };
        inserted.push(TableColumn {
            id: next_id,
            name,
            formula: None,
            totals_formula: None,
        });
        next_id += 1;
    }

    let idx = insert_idx.min(table.columns.len());
    table.columns.splice(idx..idx, inserted);
}

fn adjust_range_for_insert_cols(mut range: Range, col: u32, count: u32) -> Range {
    if count == 0 {
        return range;
    }

    if col <= range.start.col {
        range.start.col = range.start.col.saturating_add(count);
        range.end.col = range.end.col.saturating_add(count);
    } else if col <= range.end.col {
        range.end.col = range.end.col.saturating_add(count);
    }

    range
}

fn adjust_range_for_delete_cols(range: Range, col: u32, count: u32) -> Option<Range> {
    if count == 0 {
        return Some(range);
    }

    let del_end = col.saturating_add(count.saturating_sub(1));

    if range.end.col < col {
        return Some(range);
    }

    if range.start.col > del_end {
        return Some(Range::new(
            CellRef::new(range.start.row, range.start.col.saturating_sub(count)),
            CellRef::new(range.end.row, range.end.col.saturating_sub(count)),
        ));
    }

    // Deletion overlaps the range's columns.
    let new_start_col = if range.start.col >= col {
        col
    } else {
        range.start.col
    };
    let new_end_col = if range.end.col > del_end {
        range.end.col.saturating_sub(count)
    } else {
        // The deletion wipes out the right edge of the range; it becomes empty.
        if col == 0 {
            return None;
        }
        col - 1
    };

    if new_start_col > new_end_col {
        return None;
    }

    Some(Range::new(
        CellRef::new(range.start.row, new_start_col),
        CellRef::new(range.end.row, new_end_col),
    ))
}

fn update_tables_for_insert_cols(sheet: &mut Sheet, col: u32, count: u32) {
    if count == 0 {
        return;
    }

    for table in &mut sheet.tables {
        // Keep invariants explicit: table column metadata should always match the table's width.
        normalize_table_columns(table);

        let start_col = table.range.start.col;
        let end_col = table.range.end.col;

        if col < start_col {
            // Inserting strictly before the table shifts its range, but doesn't add new table
            // columns.
            table.range.start.col = start_col.saturating_add(count);
            table.range.end.col = end_col.saturating_add(count);
        } else if col <= end_col.saturating_add(1) {
            // Inserting within the table adds columns to the table at the insertion point.
            //
            // We treat inserting at `end_col + 1` as inserting at the table's right edge (i.e.
            // appending new columns), matching how Excel expands tables when inserting columns at
            // the boundary.
            let insert_idx = (col - start_col) as usize;
            insert_default_table_columns(table, insert_idx, count);
            table.range.end.col = end_col.saturating_add(count);
        }

        if let Some(auto_filter) = table.auto_filter.as_mut() {
            // Best-effort keep AutoFilter metadata in sync with the table range.
            auto_filter.range = table.range;
            let filter_range = auto_filter.range;

            // Shift filter column ids for inserts that land within the table's column span.
            if col >= start_col && col <= end_col.saturating_add(1) {
                let insert_idx = col - start_col;
                for filter_column in &mut auto_filter.filter_columns {
                    if filter_column.col_id >= insert_idx {
                        filter_column.col_id = filter_column.col_id.saturating_add(count);
                    }
                }
            }

            if let Some(sort_state) = auto_filter.sort_state.as_mut() {
                for condition in &mut sort_state.conditions {
                    condition.range = adjust_range_for_insert_cols(condition.range, col, count);
                }
                sort_state
                    .conditions
                    .retain(|cond| cond.range.intersects(&filter_range));
            }
        }

        normalize_table_columns(table);
    }
}

fn update_tables_for_delete_cols(sheet: &mut Sheet, col: u32, count: u32) {
    if count == 0 {
        return;
    }

    let del_end = col.saturating_add(count.saturating_sub(1));

    sheet.tables.retain_mut(|table| {
        normalize_table_columns(table);

        let start_col = table.range.start.col;
        let end_col = table.range.end.col;

        if del_end < start_col {
            // Deleting strictly before the table shifts it left without removing columns.
            table.range.start.col = start_col.saturating_sub(count);
            table.range.end.col = end_col.saturating_sub(count);
            if let Some(auto_filter) = table.auto_filter.as_mut() {
                auto_filter.range = table.range;
                let filter_range = auto_filter.range;
                if let Some(sort_state) = auto_filter.sort_state.as_mut() {
                    sort_state.conditions.retain_mut(|cond| {
                        let Some(updated) = adjust_range_for_delete_cols(cond.range, col, count)
                        else {
                            return false;
                        };
                        cond.range = updated;
                        cond.range.intersects(&filter_range)
                    });
                }
            }
            normalize_table_columns(table);
            return true;
        }

        if col > end_col {
            // Deleting strictly after the table does not affect it.
            if let Some(auto_filter) = table.auto_filter.as_mut() {
                auto_filter.range = table.range;
                let filter_range = auto_filter.range;
                if let Some(sort_state) = auto_filter.sort_state.as_mut() {
                    sort_state.conditions.retain_mut(|cond| {
                        let Some(updated) = adjust_range_for_delete_cols(cond.range, col, count)
                        else {
                            return false;
                        };
                        cond.range = updated;
                        cond.range.intersects(&filter_range)
                    });
                }
            }
            normalize_table_columns(table);
            return true;
        }

        // Deletion overlaps the table's columns.
        let overlap_start = col.max(start_col);
        let overlap_end = del_end.min(end_col);
        let overlap_count = overlap_end.saturating_sub(overlap_start) + 1;
        let old_width = end_col.saturating_sub(start_col) + 1;
        let new_width = old_width.saturating_sub(overlap_count);

        // Deterministic behavior: if all columns are deleted, drop the table metadata.
        if new_width == 0 {
            return false;
        }

        let rel_start = (overlap_start - start_col) as usize;
        let rel_end_exclusive = (overlap_end - start_col + 1) as usize;
        let drain_end = rel_end_exclusive.min(table.columns.len());
        if rel_start < drain_end {
            table.columns.drain(rel_start..drain_end);
        }

        let deleted_before = if col < start_col {
            (start_col - col).min(count)
        } else {
            0
        };
        table.range.start.col = start_col.saturating_sub(deleted_before);
        table.range.end.col = table.range.start.col.saturating_add(new_width - 1);

        if let Some(auto_filter) = table.auto_filter.as_mut() {
            auto_filter.range = table.range;
            let filter_range = auto_filter.range;

            // Shift or drop filter column ids so they remain aligned with the table's columns.
            let rel_start = overlap_start.saturating_sub(start_col);
            let rel_end = overlap_end.saturating_sub(start_col);
            auto_filter.filter_columns.retain_mut(|filter_column| {
                if filter_column.col_id < rel_start {
                    return true;
                }
                if filter_column.col_id <= rel_end {
                    return false;
                }
                filter_column.col_id = filter_column.col_id.saturating_sub(overlap_count);
                true
            });

            if let Some(sort_state) = auto_filter.sort_state.as_mut() {
                sort_state.conditions.retain_mut(|cond| {
                    let Some(updated) = adjust_range_for_delete_cols(cond.range, col, count) else {
                        return false;
                    };
                    cond.range = updated;
                    cond.range.intersects(&filter_range)
                });
            }
        }

        normalize_table_columns(table);
        true
    });
}

fn insert_cells_shift_right(sheet: &mut Sheet, range: Range, width: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.row >= range.start.row && addr.row <= range.end.row && addr.col >= range.start.col {
            new_cells.insert(
                CellAddr {
                    row: addr.row,
                    col: addr.col + width,
                },
                cell,
            );
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn insert_cells_shift_down(sheet: &mut Sheet, range: Range, height: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.col >= range.start.col && addr.col <= range.end.col && addr.row >= range.start.row {
            new_cells.insert(
                CellAddr {
                    row: addr.row + height,
                    col: addr.col,
                },
                cell,
            );
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn delete_cells_shift_left(sheet: &mut Sheet, range: Range, width: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.row >= range.start.row && addr.row <= range.end.row {
            if addr.col >= range.start.col && addr.col <= range.end.col {
                continue;
            }
            if addr.col > range.end.col {
                new_cells.insert(
                    CellAddr {
                        row: addr.row,
                        col: addr.col - width,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn delete_cells_shift_up(sheet: &mut Sheet, range: Range, height: u32) {
    let mut new_cells = HashMap::with_capacity(sheet.cells.len());
    for (addr, cell) in std::mem::take(&mut sheet.cells) {
        if addr.col >= range.start.col && addr.col <= range.end.col {
            if addr.row >= range.start.row && addr.row <= range.end.row {
                continue;
            }
            if addr.row > range.end.row {
                new_cells.insert(
                    CellAddr {
                        row: addr.row - height,
                        col: addr.col,
                    },
                    cell,
                );
            } else {
                new_cells.insert(addr, cell);
            }
        } else {
            new_cells.insert(addr, cell);
        }
    }
    sheet.cells = new_cells;
}

fn move_range(sheet: &mut Sheet, src: Range, dst_top_left: CellRef) {
    let dst_top_left_addr = cell_addr_from_cell_ref(dst_top_left);
    let dst = Range::new(
        dst_top_left,
        CellRef::new(
            dst_top_left.row + src.height() - 1,
            dst_top_left.col + src.width() - 1,
        ),
    );

    let mut extracted: Vec<(CellRef, Option<Cell>)> = Vec::new();
    for cell in src.iter() {
        extracted.push((cell, sheet.cells.remove(&cell_addr_from_cell_ref(cell))));
    }

    for cell in dst.iter() {
        sheet.cells.remove(&cell_addr_from_cell_ref(cell));
    }

    for (cell, value) in extracted {
        let Some(value) = value else { continue };
        let dr = cell.row - src.start.row;
        let dc = cell.col - src.start.col;
        sheet.cells.insert(
            CellAddr {
                row: dst_top_left_addr.row + dr,
                col: dst_top_left_addr.col + dc,
            },
            value,
        );
    }
}

fn copy_range(
    sheet: &mut Sheet,
    sheet_name: &str,
    src: Range,
    dst_top_left: CellRef,
    formula_rewrites: &mut Vec<FormulaRewrite>,
) {
    let dst = Range::new(
        dst_top_left,
        CellRef::new(
            dst_top_left.row + src.height() - 1,
            dst_top_left.col + src.width() - 1,
        ),
    );
    let delta_row = dst.start.row as i32 - src.start.row as i32;
    let delta_col = dst.start.col as i32 - src.start.col as i32;

    let mut extracted: Vec<(CellRef, Option<Cell>)> = Vec::new();
    for cell in src.iter() {
        extracted.push((
            cell,
            sheet.cells.get(&cell_addr_from_cell_ref(cell)).cloned(),
        ));
    }

    for cell in dst.iter() {
        sheet.cells.remove(&cell_addr_from_cell_ref(cell));
    }

    for (cell, value) in extracted {
        let Some(mut value) = value else { continue };
        let dr = cell.row - src.start.row;
        let dc = cell.col - src.start.col;
        let target = CellRef::new(dst.start.row + dr, dst.start.col + dc);

        if let Some(formula) = &value.formula {
            let origin = crate::CellAddr::new(target.row, target.col);
            let (new_formula, _) =
                rewrite_formula_for_copy_delta(formula, sheet_name, origin, delta_row, delta_col);
            if new_formula != formula.as_ref() {
                formula_rewrites.push(FormulaRewrite {
                    sheet: sheet_name.to_string(),
                    cell: target,
                    before: formula.to_string(),
                    after: new_formula.clone(),
                });
            }
            value.formula = Some(new_formula.into());
        }

        // Copy/paste-style operations overwrite cell input but do not explicitly set phonetic
        // metadata. Clear it to avoid returning stale furigana via PHONETIC().
        value.phonetic = None;

        sheet.cells.insert(cell_addr_from_cell_ref(target), value);
    }
}

fn fill_range(
    sheet: &mut Sheet,
    sheet_name: &str,
    src: Range,
    dst: Range,
    formula_rewrites: &mut Vec<FormulaRewrite>,
) {
    let height = src.height() as i32;
    let width = src.width() as i32;
    if height <= 0 || width <= 0 {
        return;
    }

    for cell in dst.iter() {
        if src.contains(cell) {
            continue;
        }
        sheet.cells.remove(&cell_addr_from_cell_ref(cell));

        let rel_row = cell.row as i32 - src.start.row as i32;
        let rel_col = cell.col as i32 - src.start.col as i32;
        let src_row = src.start.row + rel_row.rem_euclid(height) as u32;
        let src_col = src.start.col + rel_col.rem_euclid(width) as u32;
        let src_cell = CellRef::new(src_row, src_col);

        let Some(mut value) = sheet.cells.get(&cell_addr_from_cell_ref(src_cell)).cloned() else {
            continue;
        };
        if let Some(formula) = &value.formula {
            let delta_row = cell.row as i32 - src_cell.row as i32;
            let delta_col = cell.col as i32 - src_cell.col as i32;
            let origin = crate::CellAddr::new(cell.row, cell.col);
            let (new_formula, _) =
                rewrite_formula_for_copy_delta(formula, sheet_name, origin, delta_row, delta_col);
            if new_formula != formula.as_ref() {
                formula_rewrites.push(FormulaRewrite {
                    sheet: sheet_name.to_string(),
                    cell,
                    before: formula.to_string(),
                    after: new_formula.clone(),
                });
            }
            value.formula = Some(new_formula.into());
        }
        // Fill operations overwrite cell input but do not explicitly set phonetic metadata. Clear
        // it to avoid returning stale furigana via PHONETIC().
        value.phonetic = None;
        sheet.cells.insert(cell_addr_from_cell_ref(cell), value);
    }
}

fn build_sheet_order_indices(workbook: &Workbook) -> HashMap<String, usize> {
    // 3D references (`Sheet1:Sheet3!A1`) use sheet *tab order* to define span membership.
    // Produce a map from case-insensitive sheet name -> tab order index so formula rewrite helpers
    // can translate sheet spans consistently.
    // Note: Formulas may reference sheets by either their stable key or their user-visible display
    // name. Include both aliases so sheet-span rewrites and edit applicability checks can resolve
    // 3D spans regardless of which naming scheme appears in the formula text.
    let mut out: HashMap<String, usize> = HashMap::with_capacity(workbook.sheet_order.len() * 2);
    for (order_index, &sheet_id) in workbook.sheet_order.iter().enumerate() {
        if let Some(name) = workbook.sheet_key_name(sheet_id) {
            out.insert(Workbook::sheet_key(name), order_index);
        }
        if let Some(name) = workbook.sheet_name(sheet_id) {
            out.insert(Workbook::sheet_key(name), order_index);
        }
    }
    out
}

fn rewrite_all_formulas_structural(
    workbook: &mut Workbook,
    sheet_names: &HashMap<SheetId, String>,
    edit: StructuralEdit,
) -> Vec<FormulaRewrite> {
    // 3D references (`Sheet1:Sheet3!A1`) use sheet *tab order* to define span membership, so use
    // the workbook's current sheet ordering rather than stable sheet ids.
    let sheet_order_indices = build_sheet_order_indices(workbook);

    let mut rewrites = Vec::new();
    for (sheet_id, sheet) in workbook.sheets.iter_mut().enumerate() {
        let Some(ctx_sheet) = sheet_names.get(&sheet_id) else {
            continue;
        };
        for (addr, cell) in sheet.cells.iter_mut() {
            let Some(formula) = &cell.formula else {
                continue;
            };
            let origin = crate::CellAddr::new(addr.row, addr.col);
            let (new_formula, changed) = rewrite_formula_for_structural_edit_with_resolver(
                formula,
                ctx_sheet,
                origin,
                &edit,
                |name| sheet_order_indices.get(&Workbook::sheet_key(name)).copied(),
            );
            if changed {
                rewrites.push(FormulaRewrite {
                    sheet: ctx_sheet.clone(),
                    cell: cell_ref_from_addr(*addr),
                    before: formula.to_string(),
                    after: new_formula.clone(),
                });
                cell.formula = Some(new_formula.into());
            }
        }

        // Table column formulas live outside the worksheet cell map, but still need to be kept in
        // sync with structural edits so XLSX table metadata round-trips correctly.
        for table in &mut sheet.tables {
            let data_origin_row = table
                .data_range()
                .map(|r| r.start.row)
                .unwrap_or(table.range.start.row);
            let totals_origin_row = table
                .totals_range()
                .map(|r| r.start.row)
                .unwrap_or(table.range.end.row);

            for (idx, column) in table.columns.iter_mut().enumerate() {
                let col = table.range.start.col.saturating_add(idx as u32);

                if let Some(formula) = column.formula.as_mut() {
                    let origin = crate::CellAddr::new(data_origin_row, col);
                    let (new_formula, changed) = rewrite_formula_for_structural_edit_with_resolver(
                        formula,
                        ctx_sheet,
                        origin,
                        &edit,
                        |name| sheet_order_indices.get(&Workbook::sheet_key(name)).copied(),
                    );
                    if changed {
                        *formula = new_formula;
                    }
                }

                if let Some(formula) = column.totals_formula.as_mut() {
                    let origin = crate::CellAddr::new(totals_origin_row, col);
                    let (new_formula, changed) = rewrite_formula_for_structural_edit_with_resolver(
                        formula,
                        ctx_sheet,
                        origin,
                        &edit,
                        |name| sheet_order_indices.get(&Workbook::sheet_key(name)).copied(),
                    );
                    if changed {
                        *formula = new_formula;
                    }
                }
            }
        }
    }
    rewrites
}

fn rewrite_all_formulas_range_map(
    workbook: &mut Workbook,
    sheet_names: &HashMap<SheetId, String>,
    edit: &RangeMapEdit,
) -> Vec<FormulaRewrite> {
    // 3D references (`Sheet1:Sheet3!A1`) use sheet *tab order* to define span membership, so use
    // the workbook's current sheet ordering rather than stable sheet ids.
    let sheet_order_indices = build_sheet_order_indices(workbook);

    let mut rewrites = Vec::new();
    for (sheet_id, sheet) in workbook.sheets.iter_mut().enumerate() {
        let Some(ctx_sheet) = sheet_names.get(&sheet_id) else {
            continue;
        };
        for (addr, cell) in sheet.cells.iter_mut() {
            let Some(formula) = &cell.formula else {
                continue;
            };
            let origin = crate::CellAddr::new(addr.row, addr.col);
            let (new_formula, changed) = rewrite_formula_for_range_map_with_resolver(
                formula,
                ctx_sheet,
                origin,
                edit,
                |name| sheet_order_indices.get(&Workbook::sheet_key(name)).copied(),
            );
            if changed {
                rewrites.push(FormulaRewrite {
                    sheet: ctx_sheet.clone(),
                    cell: cell_ref_from_addr(*addr),
                    before: formula.to_string(),
                    after: new_formula.clone(),
                });
                cell.formula = Some(new_formula.into());
            }
        }
    }
    rewrites
}

fn diff_workbooks(
    before: &Workbook,
    after: &Workbook,
    sheet_names: &HashMap<SheetId, String>,
) -> Vec<CellChange> {
    let mut out = Vec::new();
    let max_sheets = max(before.sheets.len(), after.sheets.len());
    for sheet_id in 0..max_sheets {
        let sheet_name = sheet_names
            .get(&sheet_id)
            .cloned()
            .unwrap_or_else(|| format!("Sheet{sheet_id}"));
        let before_sheet = before.sheets.get(sheet_id);
        let after_sheet = after.sheets.get(sheet_id);
        let mut addrs: BTreeSet<CellAddr> = BTreeSet::new();
        if let Some(sheet) = before_sheet {
            addrs.extend(sheet.cells.keys().copied());
        }
        if let Some(sheet) = after_sheet {
            addrs.extend(sheet.cells.keys().copied());
        }
        for addr in addrs {
            let before_cell = before_sheet.and_then(|s| s.cells.get(&addr));
            let after_cell = after_sheet.and_then(|s| s.cells.get(&addr));
            let before_snap = before_cell.map(cell_snapshot);
            let after_snap = after_cell.map(cell_snapshot);
            if before_snap == after_snap {
                continue;
            }
            out.push(CellChange {
                sheet: sheet_name.clone(),
                cell: cell_ref_from_addr(addr),
                before: before_snap,
                after: after_snap,
            });
        }
    }
    out
}

fn cell_snapshot(cell: &Cell) -> CellSnapshot {
    CellSnapshot {
        value: cell.value.clone(),
        formula: cell.formula.as_deref().map(str::to_string),
    }
}

fn sheet_id_from_graph(sheet: u32) -> SheetId {
    usize::try_from(sheet).expect("sheet id exceeds usize")
}

fn clamp_addr_to_excel_dimensions(addr: CellAddr) -> CellAddr {
    let max_row = EXCEL_MAX_ROWS.saturating_sub(1);
    let max_col = EXCEL_MAX_COLS.saturating_sub(1);
    CellAddr {
        row: if addr.row == CellAddr::SHEET_END {
            max_row
        } else {
            addr.row
        },
        col: if addr.col == CellAddr::SHEET_END {
            max_col
        } else {
            addr.col
        },
    }
}

fn clamp_addr_to_sheet_dimensions(workbook: &Workbook, sheet: SheetId, addr: CellAddr) -> CellAddr {
    let (max_row, max_col) = workbook
        .sheets
        .get(sheet)
        .map(|s| (s.row_count.saturating_sub(1), s.col_count.saturating_sub(1)))
        .unwrap_or((
            EXCEL_MAX_ROWS.saturating_sub(1),
            EXCEL_MAX_COLS.saturating_sub(1),
        ));
    CellAddr {
        row: if addr.row == CellAddr::SHEET_END {
            max_row
        } else {
            addr.row
        },
        col: if addr.col == CellAddr::SHEET_END {
            max_col
        } else {
            addr.col
        },
    }
}

fn sheet_range_to_node(range: SheetRange, workbook: &Workbook) -> PrecedentNode {
    let sheet = sheet_id_from_graph(range.sheet_id);
    PrecedentNode::Range {
        sheet,
        start: clamp_addr_to_sheet_dimensions(
            workbook,
            sheet,
            CellAddr {
                row: range.range.start.row,
                col: range.range.start.col,
            },
        ),
        end: clamp_addr_to_sheet_dimensions(
            workbook,
            sheet,
            CellAddr {
                row: range.range.end.row,
                col: range.range.end.col,
            },
        ),
    }
}

fn precedent_to_node(precedent: Precedent, workbook: &Workbook) -> PrecedentNode {
    match precedent {
        Precedent::Cell(cell) => PrecedentNode::Cell {
            sheet: sheet_id_from_graph(cell.sheet_id),
            addr: CellAddr {
                row: cell.cell.row,
                col: cell.cell.col,
            },
        },
        Precedent::Range(range) => sheet_range_to_node(range, workbook),
    }
}

fn sheet_tab_key(sheet: SheetId, tab_index_by_sheet: &[usize]) -> (usize, SheetId) {
    (
        tab_index_by_sheet.get(sheet).copied().unwrap_or(usize::MAX),
        sheet,
    )
}

fn precedent_node_cmp(
    a: &PrecedentNode,
    b: &PrecedentNode,
    tab_index_by_sheet: &[usize],
) -> Ordering {
    let rank = |node: &PrecedentNode| match node {
        PrecedentNode::Cell { .. } => 0u8,
        PrecedentNode::Range { .. } => 1,
        PrecedentNode::SpillRange { .. } => 2,
        PrecedentNode::ExternalCell { .. } => 3,
        PrecedentNode::ExternalRange { .. } => 4,
    };

    rank(a).cmp(&rank(b)).then_with(|| match (a, b) {
        (
            PrecedentNode::Cell {
                sheet: a_sheet,
                addr: a_addr,
            },
            PrecedentNode::Cell {
                sheet: b_sheet,
                addr: b_addr,
            },
        ) => sheet_tab_key(*a_sheet, tab_index_by_sheet)
            .cmp(&sheet_tab_key(*b_sheet, tab_index_by_sheet))
            .then_with(|| a_addr.row.cmp(&b_addr.row))
            .then_with(|| a_addr.col.cmp(&b_addr.col)),
        (
            PrecedentNode::Range {
                sheet: a_sheet,
                start: a_start,
                end: a_end,
            },
            PrecedentNode::Range {
                sheet: b_sheet,
                start: b_start,
                end: b_end,
            },
        ) => sheet_tab_key(*a_sheet, tab_index_by_sheet)
            .cmp(&sheet_tab_key(*b_sheet, tab_index_by_sheet))
            .then_with(|| a_start.row.cmp(&b_start.row))
            .then_with(|| a_start.col.cmp(&b_start.col))
            .then_with(|| a_end.row.cmp(&b_end.row))
            .then_with(|| a_end.col.cmp(&b_end.col)),
        (
            PrecedentNode::SpillRange {
                sheet: a_sheet,
                origin: a_origin,
                start: a_start,
                end: a_end,
            },
            PrecedentNode::SpillRange {
                sheet: b_sheet,
                origin: b_origin,
                start: b_start,
                end: b_end,
            },
        ) => sheet_tab_key(*a_sheet, tab_index_by_sheet)
            .cmp(&sheet_tab_key(*b_sheet, tab_index_by_sheet))
            .then_with(|| a_origin.row.cmp(&b_origin.row))
            .then_with(|| a_origin.col.cmp(&b_origin.col))
            .then_with(|| a_start.row.cmp(&b_start.row))
            .then_with(|| a_start.col.cmp(&b_start.col))
            .then_with(|| a_end.row.cmp(&b_end.row))
            .then_with(|| a_end.col.cmp(&b_end.col)),
        (
            PrecedentNode::ExternalCell {
                sheet: a_sheet,
                addr: a_addr,
            },
            PrecedentNode::ExternalCell {
                sheet: b_sheet,
                addr: b_addr,
            },
        ) => a_sheet
            .cmp(b_sheet)
            .then_with(|| a_addr.row.cmp(&b_addr.row))
            .then_with(|| a_addr.col.cmp(&b_addr.col)),
        (
            PrecedentNode::ExternalRange {
                sheet: a_sheet,
                start: a_start,
                end: a_end,
            },
            PrecedentNode::ExternalRange {
                sheet: b_sheet,
                start: b_start,
                end: b_end,
            },
        ) => a_sheet
            .cmp(b_sheet)
            .then_with(|| a_start.row.cmp(&b_start.row))
            .then_with(|| a_start.col.cmp(&b_start.col))
            .then_with(|| a_end.row.cmp(&b_end.row))
            .then_with(|| a_end.col.cmp(&b_end.col)),
        _ => Ordering::Equal,
    })
}

fn sort_and_dedup_nodes(nodes: &mut Vec<PrecedentNode>, workbook: &Workbook) {
    let tab_index_by_sheet = workbook.tab_index_by_sheet_id();
    nodes.sort_by(|a, b| precedent_node_cmp(a, b, tab_index_by_sheet));
    nodes.dedup();
}

fn normalize_range(start: CellAddr, end: CellAddr) -> (CellAddr, CellAddr) {
    let start_row = start.row.min(end.row);
    let end_row = start.row.max(end.row);
    let start_col = start.col.min(end.col);
    let end_col = start.col.max(end.col);
    (
        CellAddr {
            row: start_row,
            col: start_col,
        },
        CellAddr {
            row: end_row,
            col: end_col,
        },
    )
}

fn expand_nodes_to_cells(
    nodes: &[PrecedentNode],
    limit: usize,
    workbook: &Workbook,
) -> Vec<(SheetId, CellAddr)> {
    #[derive(Debug, Clone)]
    enum Stream {
        Empty,
        Single {
            sheet: SheetId,
            addr: CellAddr,
            done: bool,
        },
        Range {
            sheet: SheetId,
            start: CellAddr,
            end: CellAddr,
            cur: CellAddr,
            done: bool,
        },
    }

    impl Stream {
        fn from_node(node: PrecedentNode) -> Self {
            match node {
                PrecedentNode::Cell { sheet, addr } => Stream::Single {
                    sheet,
                    addr,
                    done: false,
                },
                PrecedentNode::Range { sheet, start, end } => {
                    let (start, end) = normalize_range(start, end);
                    Stream::Range {
                        sheet,
                        start,
                        end,
                        cur: start,
                        done: false,
                    }
                }
                PrecedentNode::SpillRange {
                    sheet, start, end, ..
                } => {
                    let (start, end) = normalize_range(start, end);
                    Stream::Range {
                        sheet,
                        start,
                        end,
                        cur: start,
                        done: false,
                    }
                }
                PrecedentNode::ExternalCell { .. } | PrecedentNode::ExternalRange { .. } => {
                    Stream::Empty
                }
            }
        }

        fn peek(&self) -> Option<(SheetId, CellAddr)> {
            match self {
                Stream::Empty => None,
                Stream::Single { sheet, addr, done } => (!*done).then_some((*sheet, *addr)),
                Stream::Range {
                    sheet, cur, done, ..
                } => (!*done).then_some((*sheet, *cur)),
            }
        }

        fn advance(&mut self) {
            match self {
                Stream::Empty => {}
                Stream::Single { done, .. } => *done = true,
                Stream::Range {
                    start,
                    end,
                    cur,
                    done,
                    ..
                } => {
                    if *done {
                        return;
                    }
                    if cur.row == end.row && cur.col == end.col {
                        *done = true;
                        return;
                    }
                    if cur.col < end.col {
                        cur.col += 1;
                    } else {
                        cur.col = start.col;
                        cur.row += 1;
                        if cur.row > end.row {
                            *done = true;
                        }
                    }
                }
            }
        }
    }

    if limit == 0 || nodes.is_empty() {
        return Vec::new();
    }

    let mut nodes: Vec<PrecedentNode> = nodes.to_vec();
    sort_and_dedup_nodes(&mut nodes, workbook);

    let mut streams: Vec<Stream> = nodes.into_iter().map(Stream::from_node).collect();
    let tab_index_by_sheet = workbook.tab_index_by_sheet_id();
    let mut heap: std::collections::BinaryHeap<
        std::cmp::Reverse<(usize, u32, u32, SheetId, usize)>,
    > = std::collections::BinaryHeap::new();

    for (idx, stream) in streams.iter().enumerate() {
        if let Some((sheet, addr)) = stream.peek() {
            heap.push(std::cmp::Reverse((
                tab_index_by_sheet.get(sheet).copied().unwrap_or(usize::MAX),
                addr.row,
                addr.col,
                sheet,
                idx,
            )));
        }
    }

    let mut seen: HashSet<CellKey> = HashSet::new();
    let mut out: Vec<(SheetId, CellAddr)> = Vec::new();
    out.reserve(limit.min(1024));

    while out.len() < limit {
        let Some(std::cmp::Reverse((_tab, row, col, sheet, idx))) = heap.pop() else {
            break;
        };

        let addr = CellAddr { row, col };
        if seen.insert(CellKey { sheet, addr }) {
            out.push((sheet, addr));
        }

        let stream = streams
            .get_mut(idx)
            .expect("heap indices are valid stream indices");
        stream.advance();
        if let Some((sheet, addr)) = stream.peek() {
            heap.push(std::cmp::Reverse((
                tab_index_by_sheet.get(sheet).copied().unwrap_or(usize::MAX),
                addr.row,
                addr.col,
                sheet,
                idx,
            )));
        }
    }

    out
}

fn canonical_expr_contains_structured_refs(expr: &crate::Expr) -> bool {
    match expr {
        crate::Expr::StructuredRef(_) => true,
        crate::Expr::FieldAccess(access) => {
            canonical_expr_contains_structured_refs(access.base.as_ref())
        }
        crate::Expr::FunctionCall(call) => call
            .args
            .iter()
            .any(|arg| canonical_expr_contains_structured_refs(arg)),
        crate::Expr::Call(call) => {
            canonical_expr_contains_structured_refs(call.callee.as_ref())
                || call
                    .args
                    .iter()
                    .any(|arg| canonical_expr_contains_structured_refs(arg))
        }
        crate::Expr::Unary(u) => canonical_expr_contains_structured_refs(&u.expr),
        crate::Expr::Postfix(p) => canonical_expr_contains_structured_refs(&p.expr),
        crate::Expr::Binary(b) => {
            canonical_expr_contains_structured_refs(&b.left)
                || canonical_expr_contains_structured_refs(&b.right)
        }
        crate::Expr::Array(arr) => arr
            .rows
            .iter()
            .flat_map(|row| row.iter())
            .any(|el| canonical_expr_contains_structured_refs(el)),
        crate::Expr::Number(_)
        | crate::Expr::String(_)
        | crate::Expr::Boolean(_)
        | crate::Expr::Error(_)
        | crate::Expr::NameRef(_)
        | crate::Expr::CellRef(_)
        | crate::Expr::ColRef(_)
        | crate::Expr::RowRef(_)
        | crate::Expr::Missing => false,
    }
}

fn canonical_expr_contains_let_or_lambda(expr: &crate::Expr) -> bool {
    match expr {
        crate::Expr::FunctionCall(call) => {
            if matches!(call.name.name_upper.as_str(), "LET" | "LAMBDA") {
                return true;
            }
            call.args.iter().any(canonical_expr_contains_let_or_lambda)
        }
        crate::Expr::FieldAccess(access) => {
            canonical_expr_contains_let_or_lambda(access.base.as_ref())
        }
        crate::Expr::Call(call) => {
            canonical_expr_contains_let_or_lambda(call.callee.as_ref())
                || call
                    .args
                    .iter()
                    .any(|arg| canonical_expr_contains_let_or_lambda(arg))
        }
        crate::Expr::Unary(u) => canonical_expr_contains_let_or_lambda(&u.expr),
        crate::Expr::Postfix(p) => canonical_expr_contains_let_or_lambda(&p.expr),
        crate::Expr::Binary(b) => {
            canonical_expr_contains_let_or_lambda(&b.left)
                || canonical_expr_contains_let_or_lambda(&b.right)
        }
        crate::Expr::Array(arr) => arr
            .rows
            .iter()
            .flatten()
            .any(|el| canonical_expr_contains_let_or_lambda(el)),
        _ => false,
    }
}

#[derive(Clone, Copy, Default)]
struct PrefixLowerErrorFlags {
    external_reference: bool,
    unknown_sheet: bool,
}

fn canonical_expr_depends_on_lowering_prefix_error(
    expr: &crate::Expr,
    current_sheet: SheetId,
    workbook: &Workbook,
) -> Option<bytecode::LowerError> {
    let mut flags = PrefixLowerErrorFlags::default();
    canonical_expr_collect_sheet_prefix_errors(expr, current_sheet, workbook, &mut flags);

    if flags.external_reference {
        return Some(bytecode::LowerError::ExternalReference);
    }

    // Avoid chasing defined names when LET/LAMBDA might introduce lexical bindings that shadow
    // workbook/sheet defined names. Direct reference prefixes are still detected above.
    if canonical_expr_contains_let_or_lambda(expr) {
        if flags.unknown_sheet {
            return Some(bytecode::LowerError::UnknownSheet);
        }
        return None;
    }

    let mut visiting: HashSet<(SheetId, String)> = HashSet::new();
    canonical_expr_collect_defined_name_prefix_errors(
        expr,
        current_sheet,
        workbook,
        &mut visiting,
        &mut flags,
    );

    if flags.external_reference {
        Some(bytecode::LowerError::ExternalReference)
    } else if flags.unknown_sheet {
        Some(bytecode::LowerError::UnknownSheet)
    } else {
        None
    }
}

fn canonical_expr_collect_sheet_prefix_errors(
    expr: &crate::Expr,
    current_sheet: SheetId,
    workbook: &Workbook,
    flags: &mut PrefixLowerErrorFlags,
) {
    if flags.external_reference {
        return;
    }
    match expr {
        crate::Expr::CellRef(r) => {
            update_sheet_prefix_flags(
                &r.workbook,
                r.sheet.as_ref(),
                current_sheet,
                workbook,
                flags,
            );
        }
        crate::Expr::ColRef(r) => {
            update_sheet_prefix_flags(
                &r.workbook,
                r.sheet.as_ref(),
                current_sheet,
                workbook,
                flags,
            );
        }
        crate::Expr::RowRef(r) => {
            update_sheet_prefix_flags(
                &r.workbook,
                r.sheet.as_ref(),
                current_sheet,
                workbook,
                flags,
            );
        }
        crate::Expr::FieldAccess(access) => {
            canonical_expr_collect_sheet_prefix_errors(
                access.base.as_ref(),
                current_sheet,
                workbook,
                flags,
            );
        }
        crate::Expr::FunctionCall(call) => {
            for arg in &call.args {
                canonical_expr_collect_sheet_prefix_errors(arg, current_sheet, workbook, flags);
            }
        }
        crate::Expr::Call(call) => {
            canonical_expr_collect_sheet_prefix_errors(
                call.callee.as_ref(),
                current_sheet,
                workbook,
                flags,
            );
            for arg in &call.args {
                canonical_expr_collect_sheet_prefix_errors(arg, current_sheet, workbook, flags);
            }
        }
        crate::Expr::Unary(u) => {
            canonical_expr_collect_sheet_prefix_errors(&u.expr, current_sheet, workbook, flags);
        }
        crate::Expr::Postfix(p) => {
            canonical_expr_collect_sheet_prefix_errors(&p.expr, current_sheet, workbook, flags);
        }
        crate::Expr::Binary(b) => {
            canonical_expr_collect_sheet_prefix_errors(&b.left, current_sheet, workbook, flags);
            canonical_expr_collect_sheet_prefix_errors(&b.right, current_sheet, workbook, flags);
        }
        crate::Expr::Array(arr) => {
            for el in arr.rows.iter().flatten() {
                canonical_expr_collect_sheet_prefix_errors(el, current_sheet, workbook, flags);
            }
        }
        crate::Expr::NameRef(_)
        | crate::Expr::StructuredRef(_)
        | crate::Expr::Number(_)
        | crate::Expr::String(_)
        | crate::Expr::Boolean(_)
        | crate::Expr::Error(_)
        | crate::Expr::Missing => {}
    }
}

fn update_sheet_prefix_flags(
    workbook_prefix: &Option<String>,
    sheet: Option<&crate::SheetRef>,
    _current_sheet: SheetId,
    workbook: &Workbook,
    flags: &mut PrefixLowerErrorFlags,
) {
    if let Some(book) = workbook_prefix.as_ref() {
        // External workbook reference. This is supported by the bytecode backend as long as we can
        // construct a canonical external sheet key that matches what `ExternalValueProvider`
        // expects (e.g. `"[Book.xlsx]Sheet1"`).
        //
        // External 3D sheet spans (`[Book]Sheet1:Sheet3!A1`) cannot be represented via
        // `ExternalValueProvider`, so treat them as a lowering error to surface `#REF!` rather than
        // silently querying the provider with a misleading key.
        let key = match sheet {
            Some(crate::SheetRef::Sheet(sheet)) => format!("[{book}]{sheet}"),
            Some(crate::SheetRef::SheetRange { start, end }) => {
                if formula_model::sheet_name_eq_case_insensitive(start, end) {
                    format!("[{book}]{start}")
                } else {
                    format!("[{book}]{start}:{end}")
                }
            }
            None => format!("[{book}]"),
        };
        if !crate::eval::is_valid_external_sheet_key(&key) {
            flags.external_reference = true;
        }
        return;
    }

    let Some(sheet) = sheet else {
        return;
    };

    match sheet {
        crate::SheetRef::Sheet(name) => match workbook.sheet_id(name) {
            None => {
                flags.unknown_sheet = true;
            }
            Some(_sheet_id) => {}
        },
        crate::SheetRef::SheetRange { start, end } => {
            if workbook.sheet_id(start).is_none() || workbook.sheet_id(end).is_none() {
                flags.unknown_sheet = true;
            }
        }
    }
}

fn canonical_expr_collect_defined_name_prefix_errors(
    expr: &crate::Expr,
    current_sheet: SheetId,
    workbook: &Workbook,
    visiting: &mut HashSet<(SheetId, String)>,
    flags: &mut PrefixLowerErrorFlags,
) {
    if flags.external_reference {
        return;
    }

    match expr {
        crate::Expr::NameRef(nref) => {
            if nref.workbook.is_some() {
                flags.external_reference = true;
                return;
            }

            let name_key = normalize_defined_name(&nref.name);
            if name_key.is_empty() {
                return;
            }

            let sheet_id = match nref.sheet.as_ref() {
                None => Some(current_sheet),
                Some(sheet_ref) => sheet_ref
                    .as_single_sheet()
                    .and_then(|name| workbook.sheet_id(name)),
            };
            let Some(sheet_id) = sheet_id else {
                return;
            };

            canonical_expr_collect_defined_name_prefix_errors_for_name(
                sheet_id, &name_key, workbook, visiting, flags,
            );
        }
        crate::Expr::FieldAccess(access) => {
            canonical_expr_collect_defined_name_prefix_errors(
                access.base.as_ref(),
                current_sheet,
                workbook,
                visiting,
                flags,
            );
        }
        crate::Expr::FunctionCall(call) => {
            for arg in &call.args {
                canonical_expr_collect_defined_name_prefix_errors(
                    arg,
                    current_sheet,
                    workbook,
                    visiting,
                    flags,
                );
            }
        }
        crate::Expr::Call(call) => {
            canonical_expr_collect_defined_name_prefix_errors(
                call.callee.as_ref(),
                current_sheet,
                workbook,
                visiting,
                flags,
            );
            for arg in &call.args {
                canonical_expr_collect_defined_name_prefix_errors(
                    arg,
                    current_sheet,
                    workbook,
                    visiting,
                    flags,
                );
            }
        }
        crate::Expr::Unary(u) => {
            canonical_expr_collect_defined_name_prefix_errors(
                &u.expr,
                current_sheet,
                workbook,
                visiting,
                flags,
            );
        }
        crate::Expr::Postfix(p) => {
            canonical_expr_collect_defined_name_prefix_errors(
                &p.expr,
                current_sheet,
                workbook,
                visiting,
                flags,
            );
        }
        crate::Expr::Binary(b) => {
            canonical_expr_collect_defined_name_prefix_errors(
                &b.left,
                current_sheet,
                workbook,
                visiting,
                flags,
            );
            canonical_expr_collect_defined_name_prefix_errors(
                &b.right,
                current_sheet,
                workbook,
                visiting,
                flags,
            );
        }
        crate::Expr::Array(arr) => {
            for el in arr.rows.iter().flatten() {
                canonical_expr_collect_defined_name_prefix_errors(
                    el,
                    current_sheet,
                    workbook,
                    visiting,
                    flags,
                );
            }
        }
        crate::Expr::CellRef(_)
        | crate::Expr::ColRef(_)
        | crate::Expr::RowRef(_)
        | crate::Expr::StructuredRef(_)
        | crate::Expr::Number(_)
        | crate::Expr::String(_)
        | crate::Expr::Boolean(_)
        | crate::Expr::Error(_)
        | crate::Expr::Missing => {}
    }
}

fn canonical_expr_collect_defined_name_prefix_errors_for_name(
    sheet_id: SheetId,
    name_key: &str,
    workbook: &Workbook,
    visiting: &mut HashSet<(SheetId, String)>,
    flags: &mut PrefixLowerErrorFlags,
) {
    if flags.external_reference {
        return;
    }

    let visit_key = (sheet_id, name_key.to_string());
    if !visiting.insert(visit_key.clone()) {
        return;
    }

    let Some(def) = resolve_defined_name(workbook, sheet_id, name_key) else {
        visiting.remove(&visit_key);
        return;
    };

    match &def.definition {
        NameDefinition::Constant(_) => {}
        NameDefinition::Reference(formula) | NameDefinition::Formula(formula) => {
            if let Ok(ast) = crate::parse_formula(formula, crate::ParseOptions::default()) {
                canonical_expr_collect_sheet_prefix_errors(&ast.expr, sheet_id, workbook, flags);

                if !flags.external_reference && !canonical_expr_contains_let_or_lambda(&ast.expr) {
                    canonical_expr_collect_defined_name_prefix_errors(
                        &ast.expr, sheet_id, workbook, visiting, flags,
                    );
                }
            }
        }
    }

    visiting.remove(&visit_key);
}

fn rewrite_structured_refs_for_bytecode(
    expr: &crate::Expr,
    origin_sheet: usize,
    origin_cell: CellAddr,
    tables_by_sheet: &[Vec<Table>],
) -> Option<crate::Expr> {
    fn abs_cell_ref(addr: CellAddr) -> crate::Expr {
        crate::Expr::CellRef(crate::CellRef {
            workbook: None,
            sheet: None,
            col: crate::Coord::A1 {
                index: addr.col,
                abs: true,
            },
            row: crate::Coord::A1 {
                index: addr.row,
                abs: true,
            },
        })
    }

    fn this_row_cell_ref(col: u32) -> crate::Expr {
        crate::Expr::CellRef(crate::CellRef {
            workbook: None,
            sheet: None,
            // Structured refs like `[@Col]` should keep referring to the same *column*, even if the
            // formula is moved horizontally, but should follow the current row.
            col: crate::Coord::A1 {
                index: col,
                abs: true,
            },
            row: crate::Coord::Offset(0),
        })
    }

    fn build_union_expr(mut parts: Vec<crate::Expr>) -> crate::Expr {
        let mut iter = parts.drain(..);
        let mut acc = iter
            .next()
            .expect("caller must provide at least one union operand");
        for expr in iter {
            acc = crate::Expr::Binary(crate::BinaryExpr {
                op: crate::BinaryOp::Union,
                left: Box::new(acc),
                right: Box::new(expr),
            });
        }
        acc
    }

    match expr {
        crate::Expr::StructuredRef(r) => {
            // External workbook structured references are accepted syntactically but not supported.
            if r.workbook.is_some() {
                return None;
            }

            // The structured-ref resolver is table-name driven when available, so ignore any
            // explicit sheet prefix (matching the evaluation compiler behavior).
            let mut text = String::new();
            if let Some(table) = &r.table {
                text.push_str(table);
            }
            text.push('[');
            text.push_str(&r.spec);
            text.push(']');

            let (sref, end) = crate::structured_refs::parse_structured_ref(&text, 0)?;
            if end != text.len() {
                return None;
            }

            let ranges = crate::structured_refs::resolve_structured_ref(
                tables_by_sheet,
                origin_sheet,
                origin_cell,
                &sref,
            )
            .ok()?;

            // `[@Col]`/`Table1[@Col]` depends on the current row, so represent it as a row-relative
            // reference (row offset = 0) with an absolute column coordinate. This keeps the
            // compiled bytecode program reusable across table rows, while still producing the
            // correct row-dependent behavior at runtime.
            if matches!(
                sref.items.as_slice(),
                [crate::structured_refs::StructuredRefItem::ThisRow]
            ) {
                let mut parts = Vec::with_capacity(ranges.len());
                for (sheet_id, start, end) in &ranges {
                    if *sheet_id != origin_sheet {
                        return None;
                    }
                    if start.row != origin_cell.row || end.row != origin_cell.row {
                        return None;
                    }
                    if start == end {
                        parts.push(this_row_cell_ref(start.col));
                    } else {
                        parts.push(crate::Expr::Binary(crate::BinaryExpr {
                            op: crate::BinaryOp::Range,
                            left: Box::new(this_row_cell_ref(start.col)),
                            right: Box::new(this_row_cell_ref(end.col)),
                        }));
                    }
                }

                return if parts.len() == 1 {
                    parts.pop()
                } else {
                    Some(build_union_expr(parts))
                };
            }

            let mut parts = Vec::with_capacity(ranges.len());
            for (sheet_id, start, end) in &ranges {
                if *sheet_id != origin_sheet {
                    return None;
                }
                if start == end {
                    parts.push(abs_cell_ref(*start));
                } else {
                    parts.push(crate::Expr::Binary(crate::BinaryExpr {
                        op: crate::BinaryOp::Range,
                        left: Box::new(abs_cell_ref(*start)),
                        right: Box::new(abs_cell_ref(*end)),
                    }));
                }
            }
            if parts.len() == 1 {
                parts.pop()
            } else {
                Some(build_union_expr(parts))
            }
        }
        crate::Expr::FieldAccess(access) => {
            Some(crate::Expr::FieldAccess(crate::FieldAccessExpr {
                base: Box::new(rewrite_structured_refs_for_bytecode(
                    access.base.as_ref(),
                    origin_sheet,
                    origin_cell,
                    tables_by_sheet,
                )?),
                field: access.field.clone(),
            }))
        }
        crate::Expr::FunctionCall(call) => Some(crate::Expr::FunctionCall(crate::FunctionCall {
            name: call.name.clone(),
            args: call
                .args
                .iter()
                .map(|arg| {
                    rewrite_structured_refs_for_bytecode(
                        arg,
                        origin_sheet,
                        origin_cell,
                        tables_by_sheet,
                    )
                })
                .collect::<Option<Vec<_>>>()?,
        })),
        crate::Expr::Call(call) => Some(crate::Expr::Call(crate::CallExpr {
            callee: Box::new(rewrite_structured_refs_for_bytecode(
                call.callee.as_ref(),
                origin_sheet,
                origin_cell,
                tables_by_sheet,
            )?),
            args: call
                .args
                .iter()
                .map(|arg| {
                    rewrite_structured_refs_for_bytecode(
                        arg,
                        origin_sheet,
                        origin_cell,
                        tables_by_sheet,
                    )
                })
                .collect::<Option<Vec<_>>>()?,
        })),
        crate::Expr::Unary(u) => Some(crate::Expr::Unary(crate::UnaryExpr {
            op: u.op,
            expr: Box::new(rewrite_structured_refs_for_bytecode(
                &u.expr,
                origin_sheet,
                origin_cell,
                tables_by_sheet,
            )?),
        })),
        crate::Expr::Postfix(p) => Some(crate::Expr::Postfix(crate::PostfixExpr {
            op: p.op,
            expr: Box::new(rewrite_structured_refs_for_bytecode(
                &p.expr,
                origin_sheet,
                origin_cell,
                tables_by_sheet,
            )?),
        })),
        crate::Expr::Binary(b) => Some(crate::Expr::Binary(crate::BinaryExpr {
            op: b.op,
            left: Box::new(rewrite_structured_refs_for_bytecode(
                &b.left,
                origin_sheet,
                origin_cell,
                tables_by_sheet,
            )?),
            right: Box::new(rewrite_structured_refs_for_bytecode(
                &b.right,
                origin_sheet,
                origin_cell,
                tables_by_sheet,
            )?),
        })),
        crate::Expr::Array(arr) => {
            let mut rows: Vec<Vec<crate::Expr>> = Vec::with_capacity(arr.rows.len());
            for row in &arr.rows {
                let mut out_row = Vec::with_capacity(row.len());
                for el in row {
                    out_row.push(rewrite_structured_refs_for_bytecode(
                        el,
                        origin_sheet,
                        origin_cell,
                        tables_by_sheet,
                    )?);
                }
                rows.push(out_row);
            }
            Some(crate::Expr::Array(crate::ArrayLiteral { rows }))
        }
        crate::Expr::Number(_)
        | crate::Expr::String(_)
        | crate::Expr::Boolean(_)
        | crate::Expr::Error(_)
        | crate::Expr::NameRef(_)
        | crate::Expr::CellRef(_)
        | crate::Expr::ColRef(_)
        | crate::Expr::RowRef(_)
        | crate::Expr::Missing => Some(expr.clone()),
    }
}

struct Snapshot {
    sheet_keys_by_id: Vec<Option<String>>,
    sheet_display_names_by_id: Vec<Option<String>>,
    sheet_key_to_id: HashMap<String, SheetId>,
    sheet_display_name_to_id: HashMap<String, SheetId>,
    sheet_order: Vec<SheetId>,
    /// Mapping from stable sheet id to its current tab-order index.
    ///
    /// The vector length matches `sheet_keys_by_id.len()`. Deleted/missing sheet ids have a value
    /// of `usize::MAX`.
    tab_index_by_sheet_id: Vec<usize>,
    sheet_dimensions: Vec<(u32, u32)>,
    sheet_default_style_ids: Vec<Option<u32>>,
    sheet_default_col_width: Vec<Option<f32>>,
    format_runs_by_col: Vec<BTreeMap<u32, Vec<FormatRun>>>,
    text_codepage: u16,
    sheet_origin_cells: Vec<Option<CellAddr>>,
    values: HashMap<CellKey, Value>,
    style_ids: HashMap<CellKey, u32>,
    phonetics: HashMap<CellKey, String>,
    formulas: HashMap<CellKey, Arc<str>>,
    number_formats: HashMap<CellKey, String>,
    /// Stable ordering of stored cell keys (sheet, row, col) for deterministic sparse iteration.
    ///
    /// The evaluator's `iter_reference_cells` prefers iterating stored cells when the backend
    /// supports it. `HashMap` iteration order is non-deterministic across process runs, so keep a
    /// `BTreeSet` index that we can range-scan for a given sheet to preserve Excel-like row-major
    /// behavior (and stable error precedence) without scanning implicit blanks.
    ordered_cells: BTreeSet<CellKey>,
    spill_end_by_origin: HashMap<CellKey, CellAddr>,
    spill_origin_by_cell: HashMap<CellKey, CellKey>,
    tables: Vec<Vec<Table>>,
    workbook_names: HashMap<String, crate::eval::ResolvedName>,
    sheet_names: Vec<HashMap<String, crate::eval::ResolvedName>>,
    styles: StyleTable,
    row_properties: Vec<BTreeMap<u32, RowProperties>>,
    col_properties: Vec<BTreeMap<u32, ColProperties>>,
    workbook_directory: Option<String>,
    workbook_filename: Option<String>,
    external_value_provider: Option<Arc<dyn ExternalValueProvider>>,
    external_data_provider: Option<Arc<dyn ExternalDataProvider>>,
    info: EngineInfo,
    pivot_registry: crate::pivot_registry::PivotRegistry,
}

impl Snapshot {
    fn from_workbook(
        workbook: &Workbook,
        spills: &SpillState,
        external_value_provider: Option<Arc<dyn ExternalValueProvider>>,
        external_data_provider: Option<Arc<dyn ExternalDataProvider>>,
        info: EngineInfo,
        pivot_registry: crate::pivot_registry::PivotRegistry,
    ) -> Self {
        let sheet_order = workbook.sheet_ids_in_order().to_vec();
        let sheet_keys_by_id = workbook.sheet_keys.clone();
        let sheet_display_names_by_id = workbook.sheet_display_names.clone();
        let sheet_key_to_id = workbook.sheet_key_to_id.clone();
        let sheet_display_name_to_id = workbook.sheet_display_name_to_id.clone();
        let tab_index_by_sheet_id = workbook.tab_index_by_sheet_id().to_vec();
        let workbook_directory = workbook.workbook_directory.clone();
        let workbook_filename = workbook.workbook_filename.clone();
        let sheet_dimensions = workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if workbook.sheet_exists(sheet_id) {
                    (s.row_count, s.col_count)
                } else {
                    (0, 0)
                }
            })
            .collect();
        let sheet_default_style_ids = workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if workbook.sheet_exists(sheet_id) {
                    s.default_style_id
                } else {
                    None
                }
            })
            .collect();
        let sheet_default_col_width = workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if workbook.sheet_exists(sheet_id) {
                    s.default_col_width
                } else {
                    None
                }
            })
            .collect();
        let format_runs_by_col = workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if workbook.sheet_exists(sheet_id) {
                    s.format_runs_by_col.clone()
                } else {
                    BTreeMap::new()
                }
            })
            .collect();

        let sheet_origin_cells = workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if workbook.sheet_exists(sheet_id) {
                    s.origin
                } else {
                    None
                }
            })
            .collect();
        let mut cell_count = 0usize;
        for (sheet_id, sheet) in workbook.sheets.iter().enumerate() {
            if !workbook.sheet_exists(sheet_id) {
                continue;
            }
            cell_count = cell_count.saturating_add(sheet.cells.len());
        }
        for spill in spills.by_origin.values() {
            cell_count = cell_count.saturating_add(spill.array.values.len());
        }

        let mut values = HashMap::with_capacity(cell_count);
        let mut phonetics = HashMap::with_capacity(cell_count);
        let mut style_ids = HashMap::with_capacity(cell_count);
        let mut formulas = HashMap::with_capacity(cell_count);
        let mut number_formats = HashMap::new();
        let mut ordered_cells = BTreeSet::new();
        for (sheet_id, sheet) in workbook.sheets.iter().enumerate() {
            if !workbook.sheet_exists(sheet_id) {
                continue;
            }
            for (addr, cell) in &sheet.cells {
                let key = CellKey {
                    sheet: sheet_id,
                    addr: *addr,
                };
                // Mirror `Engine::get_cell_value` semantics:
                //
                // - For non-formula blank cells, treat the stored record as "style-only" so
                //   provider-backed values can flow through (and implicit blanks remain implicit).
                // - For formulas (even those producing blank), persist the computed result so it
                //   overrides any provider value.
                if cell.formula.is_some() || cell.value != Value::Blank {
                    values.insert(key, cell.value.clone());
                }
                if let Some(phonetic) = cell.phonetic.as_ref() {
                    phonetics.insert(key, phonetic.clone());
                }
                if cell.style_id != 0 {
                    style_ids.insert(key, cell.style_id);
                }
                if let Some(formula) = cell.formula.as_ref() {
                    formulas.insert(key, Arc::clone(formula));
                }
                if let Some(number_format) = cell.number_format.as_ref() {
                    number_formats.insert(key, number_format.clone());
                }
                ordered_cells.insert(key);
            }
        }

        // Overlay spilled values so formula evaluation can observe dynamic array results even
        // when the workbook map doesn't contain explicit cell records.
        let mut spill_end_by_origin = HashMap::new();
        for (origin, spill) in &spills.by_origin {
            spill_end_by_origin.insert(*origin, spill.end);
            for r in 0..spill.array.rows {
                for c in 0..spill.array.cols {
                    if r == 0 && c == 0 {
                        continue;
                    }
                    let addr = CellAddr {
                        row: origin.addr.row + r as u32,
                        col: origin.addr.col + c as u32,
                    };
                    let key = CellKey {
                        sheet: origin.sheet,
                        addr,
                    };
                    if let Some(v) = spill.array.get(r, c).cloned() {
                        values.insert(key, v);
                        ordered_cells.insert(key);
                    }
                }
            }
        }
        let spill_origin_by_cell = spills.origin_by_cell.clone();
        let tables = workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if workbook.sheet_exists(sheet_id) {
                    s.tables.clone()
                } else {
                    Vec::new()
                }
            })
            .collect();
        let row_properties = workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if workbook.sheet_exists(sheet_id) {
                    s.row_properties.clone()
                } else {
                    BTreeMap::new()
                }
            })
            .collect();
        let col_properties = workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(sheet_id, s)| {
                if workbook.sheet_exists(sheet_id) {
                    s.col_properties.clone()
                } else {
                    BTreeMap::new()
                }
            })
            .collect();

        let mut workbook_names = HashMap::new();
        for (name, def) in &workbook.names {
            workbook_names.insert(name.clone(), name_to_resolved(def));
        }

        let mut sheet_names = Vec::with_capacity(workbook.sheets.len());
        for (sheet_id, sheet) in workbook.sheets.iter().enumerate() {
            let mut names = HashMap::new();
            if workbook.sheet_exists(sheet_id) {
                for (name, def) in &sheet.names {
                    names.insert(name.clone(), name_to_resolved(def));
                }
            }
            sheet_names.push(names);
        }

        fn name_to_resolved(def: &DefinedName) -> crate::eval::ResolvedName {
            match &def.definition {
                NameDefinition::Constant(v) => crate::eval::ResolvedName::Constant(v.clone()),
                NameDefinition::Reference(_) | NameDefinition::Formula(_) => {
                    crate::eval::ResolvedName::Expr(
                        def.compiled
                            .clone()
                            .expect("non-constant defined name must have compiled expression"),
                    )
                }
            }
        }

        Self {
            sheet_keys_by_id,
            sheet_display_names_by_id,
            sheet_key_to_id,
            sheet_display_name_to_id,
            sheet_order,
            tab_index_by_sheet_id,
            workbook_directory,
            workbook_filename,
            sheet_dimensions,
            sheet_default_style_ids,
            sheet_default_col_width,
            format_runs_by_col,
            text_codepage: workbook.text_codepage,
            sheet_origin_cells,
            values,
            style_ids,
            phonetics,
            formulas,
            number_formats,
            ordered_cells,
            spill_end_by_origin,
            spill_origin_by_cell,
            tables,
            workbook_names,
            sheet_names,
            row_properties,
            col_properties,
            styles: workbook.styles.clone(),
            external_value_provider,
            external_data_provider,
            info,
            pivot_registry,
        }
    }

    fn insert_value(&mut self, key: CellKey, value: Value) {
        let existed = self.values.insert(key, value).is_some();
        if !existed {
            self.ordered_cells.insert(key);
        }
    }

    fn remove_value(&mut self, key: &CellKey) {
        if self.values.remove(key).is_some() {
            self.ordered_cells.remove(key);
        }
    }
}

impl crate::eval::ValueResolver for Snapshot {
    fn sheet_exists(&self, sheet_id: usize) -> bool {
        self.tab_index_by_sheet_id
            .get(sheet_id)
            .copied()
            .is_some_and(|idx| idx != usize::MAX)
    }

    fn sheet_order_index(&self, sheet_id: usize) -> Option<usize> {
        let idx = *self.tab_index_by_sheet_id.get(sheet_id)?;
        (idx != usize::MAX).then_some(idx)
    }

    fn expand_sheet_span(&self, start_sheet_id: usize, end_sheet_id: usize) -> Option<Vec<usize>> {
        let start_idx = *self.tab_index_by_sheet_id.get(start_sheet_id)?;
        if start_idx == usize::MAX {
            return None;
        }
        let end_idx = *self.tab_index_by_sheet_id.get(end_sheet_id)?;
        if end_idx == usize::MAX {
            return None;
        }
        let (start_idx, end_idx) = if start_idx <= end_idx {
            (start_idx, end_idx)
        } else {
            (end_idx, start_idx)
        };
        Some(self.sheet_order[start_idx..end_idx.saturating_add(1)].to_vec())
    }

    fn text_codepage(&self) -> u16 {
        self.text_codepage
    }

    fn sheet_count(&self) -> usize {
        self.sheet_order.len()
    }

    fn info_system(&self) -> Option<&str> {
        self.info.system.as_deref()
    }

    fn info_directory(&self) -> Option<&str> {
        self.info.directory.as_deref()
    }

    fn info_osversion(&self) -> Option<&str> {
        self.info.osversion.as_deref()
    }

    fn info_release(&self) -> Option<&str> {
        self.info.release.as_deref()
    }

    fn info_version(&self) -> Option<&str> {
        self.info.version.as_deref()
    }

    fn info_memavail(&self) -> Option<f64> {
        self.info.memavail
    }

    fn info_totmem(&self) -> Option<f64> {
        self.info.totmem
    }

    fn info_origin(&self, sheet_id: usize) -> Option<&str> {
        self.info
            .origin_by_sheet
            .get(&sheet_id)
            .map(|s| s.as_str())
            .or(self.info.origin.as_deref())
    }

    fn sheet_name(&self, sheet_id: usize) -> Option<&str> {
        self.sheet_display_names_by_id.get(sheet_id)?.as_deref()
    }

    fn sheet_dimensions(&self, sheet_id: usize) -> (u32, u32) {
        self.sheet_dimensions
            .get(sheet_id)
            .copied()
            .unwrap_or((0, 0))
    }

    fn sheet_default_col_width(&self, sheet_id: usize) -> Option<f32> {
        self.sheet_default_col_width
            .get(sheet_id)
            .copied()
            .flatten()
    }

    fn sheet_origin_cell(&self, sheet_id: usize) -> Option<CellAddr> {
        // Prefer explicit per-sheet view metadata configured via `Engine::set_sheet_origin`. When
        // unset, fall back to the host-provided `EngineInfo.origin` values (workbook-level or
        // per-sheet overrides) so `set_engine_info` / `set_info_origin_for_sheet` can drive
        // `INFO("origin")` without requiring sheet state mutations.
        self.sheet_origin_cells
            .get(sheet_id)
            .copied()
            .flatten()
            .or_else(|| {
                let origin = self
                    .info
                    .origin_by_sheet
                    .get(&sheet_id)
                    .map(|s| s.as_str())
                    .or(self.info.origin.as_deref())?;

                let addr = parse_a1(origin).ok()?;

                // Reject out-of-bounds origin coordinates to keep `INFO("origin")` deterministic
                // and consistent with `Engine::set_sheet_origin` validation.
                let (rows, cols) = self.sheet_dimensions(sheet_id);
                if addr.row < rows && addr.col < cols {
                    Some(addr)
                } else {
                    None
                }
            })
    }

    fn get_cell_formula(&self, sheet_id: usize, addr: CellAddr) -> Option<&str> {
        self.formulas
            .get(&CellKey {
                sheet: sheet_id,
                addr,
            })
            .map(|s| s.as_ref())
    }

    fn get_cell_phonetic(&self, sheet_id: usize, addr: CellAddr) -> Option<&str> {
        self.phonetics
            .get(&CellKey {
                sheet: sheet_id,
                addr,
            })
            .map(|s| s.as_str())
    }

    fn style_table(&self) -> Option<&StyleTable> {
        Some(&self.styles)
    }

    fn sheet_default_style_id(&self, sheet_id: usize) -> Option<u32> {
        if !self.sheet_exists(sheet_id) {
            return None;
        }
        self.sheet_default_style_ids
            .get(sheet_id)
            .copied()
            .flatten()
    }

    fn cell_style_id(&self, sheet_id: usize, addr: CellAddr) -> u32 {
        self.style_ids
            .get(&CellKey {
                sheet: sheet_id,
                addr,
            })
            .copied()
            .unwrap_or(0)
    }

    fn format_run_style_id(&self, sheet_id: usize, addr: CellAddr) -> u32 {
        self.format_runs_by_col
            .get(sheet_id)
            .and_then(|cols| cols.get(&addr.col))
            .map(|runs| {
                // Runs are expected to be sorted and non-overlapping, but we use a conservative
                // linear scan (last-match wins) to preserve deterministic behavior even if hosts
                // provide unexpected overlaps.
                let mut style_id = 0;
                for run in runs {
                    if addr.row < run.start_row {
                        break;
                    }
                    if addr.row >= run.end_row_exclusive {
                        continue;
                    }
                    style_id = run.style_id;
                }
                style_id
            })
            .unwrap_or(0)
    }

    fn row_style_id(&self, sheet_id: usize, row: u32) -> Option<u32> {
        self.row_properties
            .get(sheet_id)
            .and_then(|map| map.get(&row))
            .and_then(|props| props.style_id)
    }

    fn col_properties(&self, sheet_id: usize, col: u32) -> Option<ColProperties> {
        self.col_properties
            .get(sheet_id)
            .and_then(|map| map.get(&col))
            .cloned()
    }

    fn range_run_style_id(&self, sheet_id: usize, addr: CellAddr) -> u32 {
        let (rows, cols) = self.sheet_dimensions(sheet_id);
        if addr.row >= rows || addr.col >= cols {
            return 0;
        }

        style_id_for_row_in_runs(
            self.format_runs_by_col
                .get(sheet_id)
                .and_then(|map| map.get(&addr.col))
                .map(|runs| runs.as_slice()),
            addr.row,
        )
    }

    fn workbook_directory(&self) -> Option<&str> {
        self.workbook_directory.as_deref()
    }

    fn workbook_filename(&self) -> Option<&str> {
        self.workbook_filename.as_deref()
    }

    fn get_cell_number_format(&self, sheet_id: usize, addr: CellAddr) -> Option<&str> {
        if !self.sheet_exists(sheet_id) {
            return None;
        }

        let (rows, cols) = self.sheet_dimensions(sheet_id);
        if addr.row >= rows || addr.col >= cols {
            return None;
        }

        let key = CellKey {
            sheet: sheet_id,
            addr,
        };

        // Explicit per-cell override always wins.
        if let Some(fmt) = self.number_formats.get(&key) {
            return Some(fmt.as_str());
        }

        // Spilled outputs inherit formatting from the spill origin cell.
        if let Some(origin) = self.spill_origin_by_cell.get(&key).copied() {
            if origin != key {
                return self.get_cell_number_format(origin.sheet, origin.addr);
            }
        }

        let sheet_style_id = self.sheet_default_style_id(sheet_id).unwrap_or(0);
        let col_style_id = self
            .col_properties
            .get(sheet_id)
            .and_then(|cols| cols.get(&addr.col))
            .and_then(|props| props.style_id)
            .unwrap_or(0);
        let row_style_id = self
            .row_properties
            .get(sheet_id)
            .and_then(|rows| rows.get(&addr.row))
            .and_then(|props| props.style_id)
            .unwrap_or(0);
        let run_style_id = self
            .format_runs_by_col
            .get(sheet_id)
            .and_then(|cols| cols.get(&addr.col))
            .map(|runs| {
                // Runs are expected to be sorted and non-overlapping, but use a conservative
                // linear scan (last-match wins) to preserve deterministic behavior even if hosts
                // provide unexpected overlaps.
                let mut style_id = 0;
                for run in runs {
                    if addr.row < run.start_row {
                        break;
                    }
                    if addr.row >= run.end_row_exclusive {
                        continue;
                    }
                    style_id = run.style_id;
                }
                style_id
            })
            .unwrap_or(0);
        let cell_style_id = self.cell_style_id(sheet_id, addr);

        // Style precedence matches DocumentController layering:
        // sheet < col < row < range-run < cell.
        //
        // When a style does not specify a number format (`number_format=None`), it is treated as
        // "inherit" so lower-precedence layers can contribute the number format.
        for style_id in [
            cell_style_id,
            run_style_id,
            row_style_id,
            col_style_id,
            sheet_style_id,
        ] {
            if let Some(fmt) = self
                .styles
                .get(style_id)
                .and_then(|style| style.number_format.as_deref())
            {
                return Some(fmt);
            }
        }

        None
    }

    fn get_cell_value(&self, sheet_id: usize, addr: CellAddr) -> Value {
        let (rows, cols) = self.sheet_dimensions(sheet_id);
        if addr.row >= rows || addr.col >= cols {
            return Value::Error(ErrorKind::Ref);
        }

        if let Some(v) = self.values.get(&CellKey {
            sheet: sheet_id,
            addr,
        }) {
            return v.clone();
        }

        if let Some(provider) = &self.external_value_provider {
            if let Some(sheet_key) = self
                .sheet_keys_by_id
                .get(sheet_id)
                .and_then(|s| s.as_deref())
            {
                if let Some(v) = provider.get(sheet_key, addr) {
                    return v;
                }
            }
        }

        Value::Blank
    }

    fn cell_horizontal_alignment(
        &self,
        sheet_id: usize,
        addr: CellAddr,
    ) -> Option<HorizontalAlignment> {
        let (rows, cols) = self.sheet_dimensions(sheet_id);
        if addr.row >= rows || addr.col >= cols {
            return None;
        }

        let mut out: Option<HorizontalAlignment> = None;

        // Resolve style layers using document precedence:
        // sheet < col < row < range-run < cell.
        let sheet_style_id = self.sheet_default_style_id(sheet_id).unwrap_or(0);
        let col_style_id = self
            .col_properties
            .get(sheet_id)
            .and_then(|cols| cols.get(&addr.col))
            .and_then(|props| props.style_id)
            .unwrap_or(0);
        let row_style_id = self
            .row_properties
            .get(sheet_id)
            .and_then(|rows| rows.get(&addr.row))
            .and_then(|props| props.style_id)
            .unwrap_or(0);
        let run_style_id = self.format_run_style_id(sheet_id, addr);
        let cell_style_id = self.cell_style_id(sheet_id, addr);

        for style_id in [
            sheet_style_id,
            col_style_id,
            row_style_id,
            run_style_id,
            cell_style_id,
        ] {
            if style_id == 0 {
                continue;
            }
            let Some(style) = self.styles.get(style_id) else {
                continue;
            };
            let Some(horizontal) = style
                .alignment
                .as_ref()
                .and_then(|alignment| alignment.horizontal)
            else {
                continue;
            };
            out = Some(horizontal);
        }

        out
    }

    fn external_data_provider(&self) -> Option<&dyn ExternalDataProvider> {
        self.external_data_provider.as_deref()
    }

    fn pivot_registry(&self) -> Option<&crate::pivot_registry::PivotRegistry> {
        Some(&self.pivot_registry)
    }

    fn get_external_value(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
        self.external_value_provider
            .as_ref()
            .and_then(|provider| provider.get(sheet, addr))
    }

    fn external_sheet_order(&self, workbook: &str) -> Option<Vec<String>> {
        let provider = self.external_value_provider.as_ref()?;
        provider.sheet_order(workbook).or_else(|| {
            provider
                .workbook_sheet_names(workbook)
                .map(|names| names.as_ref().to_vec())
        })
    }

    fn workbook_sheet_names(&self, workbook: &str) -> Option<Arc<[String]>> {
        self.external_value_provider
            .as_ref()
            .and_then(|provider| provider.workbook_sheet_names(workbook))
    }

    fn external_workbook_table(&self, workbook: &str, table_name: &str) -> Option<(String, Table)> {
        self.external_value_provider
            .as_ref()
            .and_then(|provider| provider.workbook_table(workbook, table_name))
    }

    fn sheet_id(&self, name: &str) -> Option<usize> {
        // Excel resolves sheet names case-insensitively across Unicode using compatibility
        // normalization (NFKC). This ensures runtime lookups (e.g. INDIRECT, SHEET("name"))
        // agree with compile-time reference rewriting / workbook sheet-key semantics.
        //
        // Resolve by display name first (Excel tab name), with a stable-key fallback so existing
        // call sites that still pass `sheet_key` keep working.
        let key = Workbook::sheet_key(name);
        self.sheet_display_name_to_id
            .get(&key)
            .copied()
            .or_else(|| self.sheet_key_to_id.get(&key).copied())
    }

    fn iter_sheet_cells(&self, sheet_id: usize) -> Option<Box<dyn Iterator<Item = CellAddr> + '_>> {
        // When values are provided out-of-band, we cannot safely enumerate only the snapshot's
        // stored cells: provider-backed values may exist for addresses that are not present in
        // `ordered_cells`, and skipping them would produce incorrect results for range functions
        // (e.g. SUM/COUNT over provider-backed inputs).
        //
        // Fall back to dense range iteration in the evaluator by returning `None`.
        if self.external_value_provider.is_some() {
            return None;
        }
        if !self.sheet_exists(sheet_id) {
            return None;
        }
        let start = CellKey {
            sheet: sheet_id,
            addr: CellAddr { row: 0, col: 0 },
        };
        let end = CellKey {
            sheet: sheet_id,
            addr: CellAddr {
                row: u32::MAX,
                col: u32::MAX,
            },
        };
        Some(Box::new(
            self.ordered_cells.range(start..=end).map(|k| k.addr),
        ))
    }

    fn resolve_structured_ref(
        &self,
        ctx: crate::eval::EvalContext,
        sref: &crate::structured_refs::StructuredRef,
    ) -> Result<Vec<(usize, CellAddr, CellAddr)>, ErrorKind> {
        crate::structured_refs::resolve_structured_ref(
            &self.tables,
            ctx.current_sheet,
            ctx.current_cell,
            sref,
        )
    }

    fn resolve_name(&self, sheet_id: usize, name: &str) -> Option<crate::eval::ResolvedName> {
        let key = normalize_defined_name(name);
        if let Some(map) = self.sheet_names.get(sheet_id) {
            if let Some(def) = map.get(&key) {
                return Some(def.clone());
            }
        }
        if let Some(def) = self.workbook_names.get(&key) {
            return Some(def.clone());
        }

        // Excel allows referring to a table by name (e.g. `=Table1`) which resolves to the table's
        // default data body area.
        let name = name.trim();
        if name.is_empty() {
            return None;
        }
        for tables in &self.tables {
            for table in tables {
                if crate::value::cmp_case_insensitive(&table.name, name) == Ordering::Equal
                    || crate::value::cmp_case_insensitive(&table.display_name, name)
                        == Ordering::Equal
                {
                    return Some(crate::eval::ResolvedName::Expr(Expr::StructuredRef(
                        crate::eval::StructuredRefExpr {
                            sheet: crate::eval::SheetReference::Current,
                            sref: crate::structured_refs::StructuredRef {
                                table_name: Some(name.to_string()),
                                items: Vec::new(),
                                columns: crate::structured_refs::StructuredColumns::All,
                            },
                        },
                    )));
                }
            }
        }

        None
    }

    fn spill_origin(&self, sheet_id: usize, addr: CellAddr) -> Option<CellAddr> {
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        if self.spill_end_by_origin.contains_key(&key) {
            return Some(addr);
        }
        self.spill_origin_by_cell.get(&key).map(|k| k.addr)
    }

    fn spill_range(&self, sheet_id: usize, origin: CellAddr) -> Option<(CellAddr, CellAddr)> {
        let key = CellKey {
            sheet: sheet_id,
            addr: origin,
        };
        self.spill_end_by_origin.get(&key).map(|end| (origin, *end))
    }

    fn system_info(&self) -> Option<&str> {
        Some(self.info.system.as_deref().unwrap_or("pcdos"))
    }

    fn origin(&self) -> Option<&str> {
        self.info.origin.as_deref()
    }
}

fn resolve_defined_name<'a>(
    workbook: &'a Workbook,
    sheet_id: SheetId,
    name_key: &str,
) -> Option<&'a DefinedName> {
    workbook
        .sheets
        .get(sheet_id)
        .and_then(|s| s.names.get(name_key))
        .or_else(|| workbook.names.get(name_key))
}

/// Rewrite `expr` by inlining workbook/sheet-scoped constant defined names as literal values.
///
/// This is only intended for bytecode compilation. AST evaluation should continue to
/// resolve names dynamically at runtime so defined-name changes are observable.
///
/// LET/LAMBDA introduce lexical bindings that can shadow defined names. We track those bindings
/// while walking the expression so we only inline name constants when they cannot be shadowed by
/// a local identifier.
fn rewrite_defined_name_constants_for_bytecode(
    expr: &crate::Expr,
    current_sheet: SheetId,
    workbook: &Workbook,
) -> Option<crate::Expr> {
    fn name_is_local(scopes: &[HashSet<String>], name_key: &str) -> bool {
        scopes.iter().rev().any(|scope| scope.contains(name_key))
    }

    fn bare_identifier(expr: &crate::Expr) -> Option<String> {
        match expr {
            crate::Expr::NameRef(nref) if nref.workbook.is_none() && nref.sheet.is_none() => {
                let name_key = normalize_defined_name(&nref.name);
                (!name_key.is_empty()).then_some(name_key)
            }
            _ => None,
        }
    }

    fn value_to_bytecode_literal_expr(value: &Value) -> Option<crate::Expr> {
        fn scalar_value_to_bytecode_literal_expr(value: &Value) -> Option<crate::Expr> {
            match value {
                Value::Number(n) if n.is_finite() => Some(crate::Expr::Number(n.to_string())),
                Value::Number(_) => None,
                Value::Text(s) => Some(crate::Expr::String(s.clone())),
                Value::Bool(b) => Some(crate::Expr::Boolean(*b)),
                Value::Blank => Some(crate::Expr::Missing),
                Value::Error(e) => Some(crate::Expr::Error(e.as_code().to_string())),
                // Treat any other value (including rich types like Entity/Record) as non-literal for
                // bytecode inlining. This keeps name resolution conservative and avoids having to
                // serialize opaque payloads into canonical formula strings.
                Value::Entity(_)
                | Value::Record(_)
                | Value::Reference(_)
                | Value::ReferenceUnion(_)
                | Value::Array(_)
                | Value::Lambda(_)
                | Value::Spill { .. } => None,
            }
        }

        match value {
            Value::Array(arr) => {
                const MAX_INLINE_ARRAY_CELLS: usize = 256;
                let total = arr.rows.saturating_mul(arr.cols);
                if total == 0 || total > MAX_INLINE_ARRAY_CELLS {
                    return None;
                }

                let mut rows = Vec::with_capacity(arr.rows);
                for r in 0..arr.rows {
                    let mut row = Vec::with_capacity(arr.cols);
                    for c in 0..arr.cols {
                        let el = arr.get(r, c)?;
                        row.push(scalar_value_to_bytecode_literal_expr(el)?);
                    }
                    rows.push(row);
                }
                Some(crate::Expr::Array(crate::ArrayLiteral { rows }))
            }
            other => scalar_value_to_bytecode_literal_expr(other),
        }
    }

    fn inline_name_ref(
        nref: &crate::NameRef,
        current_sheet: SheetId,
        workbook: &Workbook,
        lexical_scopes: &[HashSet<String>],
    ) -> Option<crate::Expr> {
        // Bytecode can't interact with external workbooks and we don't maintain an external
        // defined-name map, so never inline external prefixes.
        if nref.workbook.is_some() {
            return None;
        }

        let sheet_id = match nref.sheet.as_ref() {
            None => current_sheet,
            Some(sheet_ref) => {
                let sheet_name = sheet_ref.as_single_sheet()?;
                workbook.sheet_id(sheet_name)?
            }
        };

        let name_key = normalize_defined_name(&nref.name);
        if name_key.is_empty() {
            return None;
        }

        // LET/LAMBDA lexical bindings are only visible for unqualified identifiers.
        // Explicit sheet-qualified names (e.g. `Sheet1!X`) should still resolve as defined names.
        if nref.sheet.is_none() && name_is_local(lexical_scopes, &name_key) {
            return None;
        }

        let def = resolve_defined_name(workbook, sheet_id, &name_key)?;
        match &def.definition {
            NameDefinition::Constant(v) => value_to_bytecode_literal_expr(v),
            NameDefinition::Reference(_) | NameDefinition::Formula(_) => None,
        }
    }

    fn rewrite_inner(
        expr: &crate::Expr,
        current_sheet: SheetId,
        workbook: &Workbook,
        lexical_scopes: &mut Vec<HashSet<String>>,
    ) -> Option<crate::Expr> {
        match expr {
            crate::Expr::NameRef(nref) => {
                inline_name_ref(nref, current_sheet, workbook, lexical_scopes)
            }
            crate::Expr::FieldAccess(access) => rewrite_inner(
                access.base.as_ref(),
                current_sheet,
                workbook,
                lexical_scopes,
            )
            .map(|inner| {
                crate::Expr::FieldAccess(crate::FieldAccessExpr {
                    base: Box::new(inner),
                    field: access.field.clone(),
                })
            }),
            crate::Expr::FunctionCall(call) if call.name.name_upper == "LET" => {
                if call.args.len() < 3 || call.args.len() % 2 == 0 {
                    return None;
                }

                lexical_scopes.push(HashSet::new());
                let mut changed = false;
                let mut args = Vec::with_capacity(call.args.len());

                for pair in call.args[..call.args.len() - 1].chunks_exact(2) {
                    let Some(name_key) = bare_identifier(&pair[0]) else {
                        lexical_scopes.pop();
                        return None;
                    };

                    // LET binding identifiers are not evaluated; keep them as written.
                    args.push(pair[0].clone());

                    if let Some(rewritten) =
                        rewrite_inner(&pair[1], current_sheet, workbook, lexical_scopes)
                    {
                        args.push(rewritten);
                        changed = true;
                    } else {
                        args.push(pair[1].clone());
                    }

                    lexical_scopes
                        .last_mut()
                        .expect("pushed scope")
                        .insert(name_key);
                }

                if let Some(rewritten) = rewrite_inner(
                    &call.args[call.args.len() - 1],
                    current_sheet,
                    workbook,
                    lexical_scopes,
                ) {
                    args.push(rewritten);
                    changed = true;
                } else {
                    args.push(call.args[call.args.len() - 1].clone());
                }

                lexical_scopes.pop();
                changed.then_some(crate::Expr::FunctionCall(crate::FunctionCall {
                    name: call.name.clone(),
                    args,
                }))
            }
            crate::Expr::FunctionCall(call) if call.name.name_upper == "LAMBDA" => {
                if call.args.is_empty() {
                    return None;
                }

                let mut scope = HashSet::new();
                for param in &call.args[..call.args.len() - 1] {
                    let Some(name_key) = bare_identifier(param) else {
                        return None;
                    };
                    if !scope.insert(name_key) {
                        return None;
                    }
                }

                lexical_scopes.push(scope);
                let mut changed = false;
                let mut args = Vec::with_capacity(call.args.len());
                args.extend(call.args[..call.args.len() - 1].iter().cloned());

                if let Some(rewritten) = rewrite_inner(
                    &call.args[call.args.len() - 1],
                    current_sheet,
                    workbook,
                    lexical_scopes,
                ) {
                    args.push(rewritten);
                    changed = true;
                } else {
                    args.push(call.args[call.args.len() - 1].clone());
                }

                lexical_scopes.pop();
                changed.then_some(crate::Expr::FunctionCall(crate::FunctionCall {
                    name: call.name.clone(),
                    args,
                }))
            }
            crate::Expr::FunctionCall(call) => {
                if matches!(
                    bytecode::ast::Function::from_name(&call.name.name_upper),
                    bytecode::ast::Function::Unknown(_)
                ) {
                    return None;
                }
                let mut args: Option<Vec<crate::Expr>> = None;
                for (idx, arg) in call.args.iter().enumerate() {
                    if let Some(rewritten) =
                        rewrite_inner(arg, current_sheet, workbook, lexical_scopes)
                    {
                        let vec = args.get_or_insert_with(|| {
                            let mut out = Vec::with_capacity(call.args.len());
                            out.extend(call.args[..idx].iter().cloned());
                            out
                        });
                        vec.push(rewritten);
                    } else if let Some(vec) = args.as_mut() {
                        vec.push(arg.clone());
                    }
                }
                args.map(|args| {
                    crate::Expr::FunctionCall(crate::FunctionCall {
                        name: call.name.clone(),
                        args,
                    })
                })
            }
            crate::Expr::Call(_) => None,
            crate::Expr::Unary(u) => {
                rewrite_inner(&u.expr, current_sheet, workbook, lexical_scopes).map(|inner| {
                    crate::Expr::Unary(crate::UnaryExpr {
                        op: u.op,
                        expr: Box::new(inner),
                    })
                })
            }
            crate::Expr::Postfix(p) => {
                rewrite_inner(&p.expr, current_sheet, workbook, lexical_scopes).map(|inner| {
                    crate::Expr::Postfix(crate::PostfixExpr {
                        op: p.op,
                        expr: Box::new(inner),
                    })
                })
            }
            crate::Expr::Binary(b) => {
                if matches!(b.op, crate::BinaryOp::Union | crate::BinaryOp::Intersect) {
                    return None;
                }
                let left = rewrite_inner(&b.left, current_sheet, workbook, lexical_scopes);
                let right = rewrite_inner(&b.right, current_sheet, workbook, lexical_scopes);
                (left.is_some() || right.is_some()).then_some(crate::Expr::Binary(
                    crate::BinaryExpr {
                        op: b.op,
                        left: Box::new(left.unwrap_or_else(|| (*b.left).clone())),
                        right: Box::new(right.unwrap_or_else(|| (*b.right).clone())),
                    },
                ))
            }
            crate::Expr::Array(arr) => {
                let mut changed = false;
                let mut rows: Vec<Vec<crate::Expr>> = Vec::with_capacity(arr.rows.len());
                for row in &arr.rows {
                    let mut out_row = Vec::with_capacity(row.len());
                    for el in row {
                        if let Some(rewritten) =
                            rewrite_inner(el, current_sheet, workbook, lexical_scopes)
                        {
                            out_row.push(rewritten);
                            changed = true;
                        } else {
                            out_row.push(el.clone());
                        }
                    }
                    rows.push(out_row);
                }
                changed.then_some(crate::Expr::Array(crate::ArrayLiteral { rows }))
            }
            crate::Expr::Number(_)
            | crate::Expr::String(_)
            | crate::Expr::Boolean(_)
            | crate::Expr::Error(_)
            | crate::Expr::CellRef(_)
            | crate::Expr::ColRef(_)
            | crate::Expr::RowRef(_)
            | crate::Expr::StructuredRef(_)
            | crate::Expr::Missing => None,
        }
    }

    let mut lexical_scopes: Vec<HashSet<String>> = Vec::new();
    rewrite_inner(expr, current_sheet, workbook, &mut lexical_scopes)
}

/// Host-provided cell values that are not stored in the engine's in-memory workbook.
///
/// This is used for:
/// - **External workbook references**, e.g. `=[Book.xlsx]Sheet1!A1`.
/// - **Out-of-band values** for the current workbook (e.g. streaming/virtualized grids). When a
///   cell is not present in the engine's internal sheet storage, the engine will query the
///   provider using the local sheet display name (e.g. `"Sheet1"`).
///
/// If no provider is configured, external workbook references evaluate to `#REF!`.
///
/// # External sheet key format
///
/// For external workbook references, the engine passes a canonical **external sheet key** as the
/// `sheet` argument:
///
/// * `"[workbook]sheet"`
///
/// Where:
/// - `workbook` is the workbook identifier inside `[...]` (e.g. `"Book.xlsx"`).
/// - `sheet` is the worksheet display name, with any formula quoting removed (e.g. `'Sheet 1'`
///   becomes `Sheet 1`).
///
/// Examples:
/// - `[Book.xlsx]Sheet1!A1` → `sheet = "[Book.xlsx]Sheet1"`
/// - `'C:\path\[Book.xlsx]Sheet1'!A1` → `sheet = "[C:\path\Book.xlsx]Sheet1"`
///   - In Rust string literals you will typically escape backslashes:
///     `'C:\\path\\[Book.xlsx]Sheet1'!A1` → `sheet = "[C:\\path\\Book.xlsx]Sheet1"`
/// - `'[Book.xlsx]Sheet 1'!A1` → `sheet = "[Book.xlsx]Sheet 1"`
/// - `'[Book.xlsx]Bob''s Sheet'!A1` → `sheet = "[Book.xlsx]Bob's Sheet"`
///
/// # Return value semantics
///
/// `get` returns an [`Option<Value>`] so providers can distinguish between a blank cell and a
/// missing/unresolvable reference:
///
/// * For **local** lookups (where `sheet` is a plain worksheet name like `"Sheet1"`), returning
///   `None` is treated as a blank cell (`Value::Blank`).
/// * For **external workbook** lookups (where `sheet` is a key like `"[Book.xlsx]Sheet1"`), returning
///   `None` is treated as an unresolved external link and evaluates to `#REF!`. Providers should
///   return `Some(Value::Blank)` to represent a blank cell in an external workbook.
/// * `addr` is 0-indexed (`A1` = `CellAddr { row: 0, col: 0 }`).
///
/// # Threading / performance
///
/// The engine may call [`ExternalValueProvider::get`] (and [`ExternalValueProvider::sheet_order`])
/// many times when evaluating range functions (e.g. `SUM([Book.xlsx]Sheet1!A:A)`).
///
/// The engine currently resolves ranges by calling `get(sheet, addr)` per-cell (there is no
/// bulk/range API). For external sheets, whole-row/whole-column references are resolved against
/// Excel’s default sheet bounds (1,048,576 rows × 16,384 columns), which can result in a very large
/// number of provider calls.
///
/// External workbook sheet dimensions are not currently exposed to the engine. Aside from resolving
/// whole-row/whole-column sentinels (`A:A`, `1:1`) against Excel’s default bounds, the engine does
/// not bounds-check external addresses—providers may see large row/col indices if formulas refer to
/// them.
///
/// Because provider-backed values may exist for addresses that are not present in the engine’s
/// internal cell storage, enabling an `ExternalValueProvider` can also force the evaluator to use
/// dense iteration for local range functions (e.g. `SUM(A:A)`), which may have performance
/// implications for large ranges.
///
/// The engine also caps materialization of rectangular references into in-memory arrays at
/// `MAX_MATERIALIZED_ARRAY_CELLS` (currently 5,000,000 cells). If a reference would exceed this
/// limit (e.g. `[Book.xlsx]Sheet1!A:XFD`), evaluation returns `#SPILL!` rather than attempting a
/// huge allocation.
///
/// When multi-threaded recalculation is enabled, provider methods may also be called concurrently
/// from multiple threads. Implementations should be thread-safe and keep lookups fast (e.g. by
/// caching results internally or minimizing lock contention).
///
/// # Volatility / invalidation
///
/// By default, formulas that reference external workbooks are treated as **volatile**: they are
/// reevaluated on every [`Engine::recalculate`] pass (Excel-compatible behavior).
///
/// Hosts can disable this behavior via [`Engine::set_external_refs_volatile(false)`] and instead
/// explicitly invalidate affected formulas via [`Engine::mark_external_sheet_dirty`] /
/// [`Engine::mark_external_workbook_dirty`].
///
/// The engine does not track dependencies to individual external cells; invalidation is coarse
/// (sheet key / workbook id).
///
/// For external-workbook 3D spans like `"[Book.xlsx]Sheet1:Sheet3"`, the engine will expand the
/// span to its component per-sheet keys for invalidation **when sheet order is available**. When
/// sheet order is unavailable, the raw span key is tracked as a single dependency (and callers can
/// fall back to workbook-level invalidation).
///
/// Note: Excel compares sheet names case-insensitively across Unicode and applies compatibility
/// normalization (NFKC). The engine preserves the formula's casing in the sheet key for single-sheet
/// external references, so providers that want Excel-compatible behavior should generally match the
/// **sheet name** portion using the same semantics (see
/// [`formula_model::sheet_name_eq_case_insensitive`]).
pub trait ExternalValueProvider: Send + Sync {
    fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value>;

    /// Return the sheet order for an external workbook as an `Arc` slice.
    ///
    /// This is equivalent to [`ExternalValueProvider::sheet_order`], but allows providers to cache
    /// and share sheet lists efficiently (cloning an `Arc` is cheaper than cloning a `Vec` of
    /// strings).
    ///
    /// The default implementation forwards to [`ExternalValueProvider::sheet_order`].
    fn workbook_sheet_names(&self, workbook: &str) -> Option<Arc<[String]>> {
        self.sheet_order(workbook).map(Arc::from)
    }

    /// Return the sheet order for an external workbook.
    ///
    /// This is required to expand external-workbook 3D spans like
    /// `"[Book.xlsx]Sheet1:Sheet3!A1"`.
    /// For example, `=SUM([Book.xlsx]'Sheet 1':'Sheet 3'!A1)` requires `sheet_order("Book.xlsx")`
    /// to expand the span.
    ///
    /// Implementations should return sheet names in workbook order (without the `[Book.xlsx]`
    /// prefix). Sheet names should be unquoted display names (e.g. return `Sheet 1`, not
    /// `'Sheet 1'`) and each sheet should appear exactly once. The order should reflect the
    /// workbook's tab order as Excel would use for 3D references (generally including hidden
    /// sheets).
    ///
    /// Endpoint matching (`Sheet1` / `Sheet3`) uses Excel’s Unicode-aware, NFKC + case-insensitive
    /// comparison semantics (see [`formula_model::sheet_name_eq_case_insensitive`]).
    ///
    /// Spans are resolved by workbook sheet order regardless of whether the user writes them
    /// “forward” or “reversed” in the formula (e.g. `Sheet3:Sheet1` is treated the same as
    /// `Sheet1:Sheet3`).
    ///
    /// The returned sheet names are used to form per-sheet keys passed to [`ExternalValueProvider::get`]
    /// (e.g. `"[Book.xlsx]{sheet_name}"`), so the casing/spelling in this list should correspond
    /// to the provider's `get` keying strategy.
    ///
    /// The input `workbook` is the raw name inside the bracketed prefix (e.g. `"Book.xlsx"` or
    /// `"C:\\path\\Book.xlsx"`). In some Excel contexts the workbook identifier may not be a
    /// filename at all (e.g. numeric workbook indices like `[1]Sheet1!A1` for other open
    /// workbooks); the engine treats it as an opaque string and passes it through as-is.
    ///
    /// For example, `=SUM('C:\path\[Book.xlsx]Sheet1:Sheet3'!A1)` calls
    /// `sheet_order("C:\path\Book.xlsx")`.
    ///
    /// Note: the engine parses the workbook identifier by splitting the `"[workbook]..."` key at
    /// the **last** `]`. This allows workbook identifiers to include bracket characters (e.g. a
    /// directory named `C:\[foo]\`). Sheet names are expected to follow Excel restrictions
    /// (notably: no `]`), so this split is unambiguous.
    ///
    /// The engine currently treats workbook identifiers as opaque strings and does not perform any
    /// additional normalization (case folding, path separator normalization, etc). Providers should
    /// normalize/match this identifier as needed.
    ///
    /// Returning `None` indicates that the sheet order is not available, in which case external
    /// 3D spans evaluate to `#REF!`.
    fn sheet_order(&self, _workbook: &str) -> Option<Vec<String>> {
        None
    }

    /// Return table metadata for an external workbook.
    ///
    /// This is used to evaluate external workbook structured references like
    /// `"[Book.xlsx]Sheet1!Table1[Col]"`.
    ///
    /// The input `workbook` is the raw name inside the bracketed prefix (e.g. `"Book.xlsx"` or
    /// `"C:\\path\\Book.xlsx"`), matching what [`crate::eval::split_external_sheet_key`] returns.
    ///
    /// Returning `None` indicates that table metadata is not available, in which case external
    /// structured references evaluate to `#REF!`.
    fn workbook_table(&self, _workbook: &str, _table_name: &str) -> Option<(String, Table)> {
        None
    }
}

pub trait ExternalDataProvider: Send + Sync {
    fn rtd(&self, prog_id: &str, server: &str, topics: &[String]) -> Value;
    fn cube_value(&self, connection: &str, tuples: &[String]) -> Value;
    fn cube_member(
        &self,
        connection: &str,
        member_expression: &str,
        caption: Option<&str>,
    ) -> Value;
    fn cube_member_property(
        &self,
        connection: &str,
        member_expression_or_handle: &str,
        property: &str,
    ) -> Value;
    fn cube_ranked_member(
        &self,
        connection: &str,
        set_expression_or_handle: &str,
        rank: i64,
        caption: Option<&str>,
    ) -> Value;
    fn cube_set(
        &self,
        connection: &str,
        set_expression: &str,
        caption: Option<&str>,
        sort_order: Option<i64>,
        sort_by: Option<&str>,
    ) -> Value;
    fn cube_set_count(&self, set_expression_or_handle: &str) -> Value;
    fn cube_kpi_member(
        &self,
        connection: &str,
        kpi_name: &str,
        kpi_property: &str,
        caption: Option<&str>,
    ) -> Value;
}

const EXCEL_MAX_COLS_I32: i32 = EXCEL_MAX_COLS as i32;
const EXCEL_MAX_ROWS_I32: i32 = EXCEL_MAX_ROWS as i32;
const BYTECODE_MAX_RANGE_CELLS: i64 = crate::eval::MAX_MATERIALIZED_ARRAY_CELLS as i64;

fn engine_error_to_bytecode(err: ErrorKind) -> bytecode::ErrorKind {
    err.into()
}

fn bytecode_error_to_engine(err: bytecode::ErrorKind) -> ErrorKind {
    err.into()
}

fn engine_value_to_bytecode(value: &Value) -> bytecode::Value {
    fn array_element_to_bytecode(value: &Value) -> bytecode::Value {
        match value {
            Value::Number(n) => bytecode::Value::Number(*n),
            Value::Bool(b) => bytecode::Value::Bool(*b),
            Value::Text(s) => bytecode::Value::Text(Arc::from(s.as_str())),
            Value::Entity(v) => bytecode::Value::Entity(Arc::new(v.clone())),
            Value::Record(v) => bytecode::Value::Record(Arc::new(v.clone())),
            Value::Blank => bytecode::Value::Empty,
            Value::Error(e) => bytecode::Value::Error(engine_error_to_bytecode(*e)),
            // Lambdas cannot appear as scalars in bytecode arrays; match spill materialization by
            // surfacing `#CALC!`.
            Value::Lambda(_) => bytecode::Value::Error(bytecode::ErrorKind::Calc),
            // References/unions cannot appear in arrays; degrade to a scalar type error.
            Value::Reference(_) | Value::ReferenceUnion(_) => {
                bytecode::Value::Error(bytecode::ErrorKind::Value)
            }
            // Nested arrays are not representable in the bytecode array model; treat them as
            // scalar type errors rather than panicking.
            Value::Array(_) => bytecode::Value::Error(bytecode::ErrorKind::Value),
            // Spill markers should not appear inside arrays, but degrade safely.
            Value::Spill { .. } => bytecode::Value::Error(bytecode::ErrorKind::Spill),
        }
    }
    match value {
        Value::Number(n) => bytecode::Value::Number(*n),
        Value::Bool(b) => bytecode::Value::Bool(*b),
        Value::Text(s) => bytecode::Value::Text(Arc::from(s.as_str())),
        Value::Entity(v) => bytecode::Value::Entity(Arc::new(v.clone())),
        Value::Record(v) => bytecode::Value::Record(Arc::new(v.clone())),
        Value::Blank => bytecode::Value::Empty,
        Value::Error(e) => bytecode::Value::Error(engine_error_to_bytecode(*e)),
        Value::Lambda(_) => bytecode::Value::Error(bytecode::ErrorKind::Calc),
        Value::Reference(_) | Value::ReferenceUnion(_) => {
            bytecode::Value::Error(bytecode::ErrorKind::Value)
        }
        Value::Array(arr) => {
            // The engine generally stores spilled array results in a dedicated spill table rather
            // than as `Value::Array` in the grid. However, callers can still populate cells with
            // `Value::Array` directly (e.g. via `set_cell_value`, rich-value fields, external value
            // providers, or tests). Bytecode evaluation should preserve these as materialized
            // arrays (bounded by `MAX_MATERIALIZED_ARRAY_CELLS`) rather than coercing them to
            // `#SPILL!`.
            let total = match arr.rows.checked_mul(arr.cols) {
                Some(v) => v,
                None => return bytecode::Value::Error(bytecode::ErrorKind::Spill),
            };
            if total != arr.values.len() {
                return bytecode::Value::Error(bytecode::ErrorKind::Num);
            }
            if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
                return bytecode::Value::Error(bytecode::ErrorKind::Spill);
            }
            let mut values = Vec::new();
            if values.try_reserve_exact(total).is_err() {
                return bytecode::Value::Error(bytecode::ErrorKind::Spill);
            }
            for v in arr.iter() {
                values.push(array_element_to_bytecode(v));
            }
            bytecode::Value::Array(bytecode::Array::new(arr.rows, arr.cols, values))
        }
        Value::Spill { .. } => bytecode::Value::Error(bytecode::ErrorKind::Spill),
    }
}

fn bytecode_value_to_engine(value: bytecode::Value) -> Value {
    match value {
        bytecode::Value::Number(n) => Value::Number(n),
        bytecode::Value::Bool(b) => Value::Bool(b),
        bytecode::Value::Text(s) => Value::Text(s.to_string()),
        bytecode::Value::Entity(v) => match Arc::try_unwrap(v) {
            Ok(entity) => Value::Entity(entity),
            Err(shared) => Value::Entity(shared.as_ref().clone()),
        },
        bytecode::Value::Record(v) => match Arc::try_unwrap(v) {
            Ok(record) => Value::Record(record),
            Err(shared) => Value::Record(shared.as_ref().clone()),
        },
        bytecode::Value::Empty => Value::Blank,
        bytecode::Value::Missing => Value::Blank,
        bytecode::Value::Error(e) => Value::Error(bytecode_error_to_engine(e)),
        bytecode::Value::Array(arr) => {
            let total = match arr.rows.checked_mul(arr.cols) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Spill),
            };
            if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut values = Vec::new();
            if values.try_reserve_exact(total).is_err() {
                return Value::Error(ErrorKind::Spill);
            }
            match Arc::try_unwrap(arr.values) {
                Ok(values_vec) => {
                    for v in values_vec {
                        // Arrays should only contain scalar values. If nested arrays/ranges appear,
                        // treat them as errors rather than attempting to spill recursively.
                        values.push(match v {
                            bytecode::Value::Array(_)
                            | bytecode::Value::Range(_)
                            | bytecode::Value::MultiRange(_) => Value::Error(ErrorKind::Value),
                            other => bytecode_value_to_engine(other),
                        });
                    }
                }
                Err(shared) => {
                    for v in shared.iter() {
                        values.push(match v {
                            bytecode::Value::Array(_)
                            | bytecode::Value::Range(_)
                            | bytecode::Value::MultiRange(_) => Value::Error(ErrorKind::Value),
                            other => bytecode_value_to_engine(other.clone()),
                        });
                    }
                }
            }
            Value::Array(Array::new(arr.rows, arr.cols, values))
        }
        // Lambdas cannot be returned from cells; match `apply_eval_result` by surfacing `#CALC!`.
        bytecode::Value::Lambda(_) => Value::Error(ErrorKind::Calc),
        // Discontiguous reference unions (e.g. `=A1,B1`) cannot be returned as a spillable array
        // result; treat them as a scalar #VALUE! like the AST evaluator.
        bytecode::Value::MultiRange(_) => Value::Error(ErrorKind::Value),
        // Bytecode arrays/refs are "spill" markers in the engine layer; other (future) rich values
        // should also degrade safely rather than panicking.
        _ => Value::Error(ErrorKind::Spill),
    }
}

#[derive(Debug, Clone, Copy)]
enum ColumnSliceMode {
    /// Use column slices only when the column range contains numbers/blanks.
    ///
    /// This is required for functions like `SUMPRODUCT` that coerce logical/text values.
    StrictNumeric,
    /// Allow column slices even when the column contains logical/text values.
    ///
    /// For SUM/AVERAGE/MIN/MAX/COUNT/COUNTIF range args, Excel ignores logical/text values, so
    /// representing them as NaN (ignored by the SIMD kernels) is correct and enables SIMD even
    /// for common "header + data" columns.
    IgnoreNonNumeric,
}

fn slice_mode_for_program(_program: &bytecode::Program) -> ColumnSliceMode {
    // Prefer allowing column slices even for "header + data" columns.
    //
    // Any bytecode runtime code paths that *require* strict numeric slices (e.g. SUMPRODUCT's
    // coercion semantics or COUNTIF/criteria comparisons that must distinguish blanks from text)
    // should call the `*_strict_numeric` slice APIs explicitly.
    ColumnSliceMode::IgnoreNonNumeric
}

#[derive(Debug, Clone)]
struct BytecodeColumnSegment {
    row_start: i32,
    values: Vec<f64>,
    blocked_rows_strict: Vec<i32>,
    blocked_rows_ignore_nonnumeric: Vec<i32>,
}

impl BytecodeColumnSegment {
    fn row_end(&self) -> i32 {
        self.row_start + self.values.len() as i32 - 1
    }
}

#[derive(Debug, Clone)]
struct BytecodeColumn {
    segments: Vec<BytecodeColumnSegment>,
}

#[derive(Debug)]
struct BytecodeColumnCache {
    by_sheet: Vec<HashMap<i32, BytecodeColumn>>,
}

impl BytecodeColumnCache {
    fn build(
        sheet_count: usize,
        snapshot: &Snapshot,
        tasks: &[(CellKey, CompiledFormula)],
    ) -> Self {
        #[derive(Debug, Clone)]
        enum StackValue {
            Range(usize),
            Ranges(Vec<usize>),
            Other,
        }

        fn range_refs_used_by_functions(program: &bytecode::Program) -> Vec<bool> {
            let mut used = vec![false; program.range_refs.len()];
            let mut stack: Vec<StackValue> = Vec::new();
            let mut locals: Vec<StackValue> = vec![StackValue::Other; program.locals.len()];

            for inst in program.instrs() {
                match inst.op() {
                    bytecode::OpCode::PushConst | bytecode::OpCode::LoadCell => {
                        stack.push(StackValue::Other);
                    }
                    bytecode::OpCode::LoadRange => {
                        stack.push(StackValue::Range(inst.a() as usize));
                    }
                    bytecode::OpCode::LoadMultiRange => {
                        // Multi-range references (e.g. 3D sheet spans) are tracked separately from
                        // `range_refs`, so treat them as opaque stack values here.
                        stack.push(StackValue::Other);
                    }
                    bytecode::OpCode::StoreLocal => {
                        let idx = inst.a() as usize;
                        let v = stack.pop().unwrap_or(StackValue::Other);
                        if idx < locals.len() {
                            locals[idx] = v;
                        }
                    }
                    bytecode::OpCode::LoadLocal => {
                        let idx = inst.a() as usize;
                        let v = locals.get(idx).cloned().unwrap_or(StackValue::Other);
                        stack.push(v);
                    }
                    bytecode::OpCode::Jump => {
                        // Control flow: conservatively ignore jumps when scanning for range refs
                        // passed to functions. This analysis is used only to decide whether
                        // building a column cache is necessary; over-approximating is acceptable.
                    }
                    bytecode::OpCode::JumpIfFalseOrError => {
                        let _ = stack.pop();
                    }
                    bytecode::OpCode::JumpIfNotError | bytecode::OpCode::JumpIfNotNaError => {}
                    bytecode::OpCode::UnaryPlus
                    | bytecode::OpCode::UnaryNeg
                    | bytecode::OpCode::ImplicitIntersection
                    | bytecode::OpCode::SpillRange => {
                        let _ = stack.pop();
                        stack.push(StackValue::Other);
                    }
                    bytecode::OpCode::Union | bytecode::OpCode::Intersect => {
                        let right = stack.pop().unwrap_or(StackValue::Other);
                        let left = stack.pop().unwrap_or(StackValue::Other);
                        let mut ranges: Vec<usize> = Vec::new();
                        match left {
                            StackValue::Range(idx) => ranges.push(idx),
                            StackValue::Ranges(mut idxs) => ranges.append(&mut idxs),
                            StackValue::Other => {}
                        }
                        match right {
                            StackValue::Range(idx) => ranges.push(idx),
                            StackValue::Ranges(mut idxs) => ranges.append(&mut idxs),
                            StackValue::Other => {}
                        }
                        stack.push(match ranges.len() {
                            0 => StackValue::Other,
                            1 => StackValue::Range(ranges[0]),
                            _ => StackValue::Ranges(ranges),
                        });
                    }
                    bytecode::OpCode::Add
                    | bytecode::OpCode::Sub
                    | bytecode::OpCode::Mul
                    | bytecode::OpCode::Div
                    | bytecode::OpCode::Pow
                    | bytecode::OpCode::Eq
                    | bytecode::OpCode::Ne
                    | bytecode::OpCode::Lt
                    | bytecode::OpCode::Le
                    | bytecode::OpCode::Gt
                    | bytecode::OpCode::Ge => {
                        let _ = stack.pop();
                        let _ = stack.pop();
                        stack.push(StackValue::Other);
                    }
                    bytecode::OpCode::CallFunc => {
                        let argc = inst.b() as usize;
                        for _ in 0..argc {
                            match stack.pop().unwrap_or(StackValue::Other) {
                                StackValue::Range(idx) => {
                                    if idx < used.len() {
                                        used[idx] = true;
                                    }
                                }
                                StackValue::Ranges(idxs) => {
                                    for idx in idxs {
                                        if idx < used.len() {
                                            used[idx] = true;
                                        }
                                    }
                                }
                                StackValue::Other => {}
                            }
                        }
                        stack.push(StackValue::Other);
                    }
                    bytecode::OpCode::MakeLambda => {
                        // Lambdas are opaque values for the purpose of range-cache analysis.
                        stack.push(StackValue::Other);
                    }
                    bytecode::OpCode::CallValue => {
                        // Pops args + callee and pushes the call result.
                        let argc = inst.b() as usize;
                        for _ in 0..argc {
                            let _ = stack.pop();
                        }
                        let _ = stack.pop(); // callee
                        stack.push(StackValue::Other);
                    }
                }
            }

            used
        }

        // Collect row windows for each referenced column so the cache can build compact
        // columnar buffers. This avoids allocating/scanning from row 0 (e.g. `A900000:A900010`),
        // and also avoids spanning huge gaps when formulas reference multiple disjoint windows.
        let mut row_ranges_by_col: Vec<HashMap<i32, Vec<(i32, i32)>>> =
            vec![HashMap::new(); sheet_count];

        let mut range_usage_cache: HashMap<usize, Vec<bool>> = HashMap::new();

        for (key, compiled) in tasks {
            let CompiledFormula::Bytecode(bc) = compiled else {
                continue;
            };

            if bc.program.range_refs.is_empty() && bc.program.multi_range_refs.is_empty() {
                continue;
            }

            let program_key = Arc::as_ptr(&bc.program) as usize;
            let has_multi_ranges = !bc.program.multi_range_refs.is_empty();
            let used = range_usage_cache
                .entry(program_key)
                .or_insert_with(|| range_refs_used_by_functions(&bc.program));
            if !has_multi_ranges && !used.iter().any(|v| *v) {
                // Ranges that are only used for implicit intersection don't require a columnar
                // cache because evaluation can fetch a single cell via `get_value`.
                continue;
            }

            let Ok(base_row) = i32::try_from(key.addr.row) else {
                continue;
            };
            let Ok(base_col) = i32::try_from(key.addr.col) else {
                continue;
            };
            let base = bytecode::CellCoord {
                row: base_row,
                col: base_col,
            };

            for (idx, range) in bc.program.range_refs.iter().enumerate() {
                if !used.get(idx).copied().unwrap_or(false) {
                    continue;
                }
                let resolved = range.resolve(base);
                let (sheet_rows, sheet_cols) = snapshot
                    .sheet_dimensions
                    .get(key.sheet)
                    .copied()
                    .unwrap_or((0, 0));
                let sheet_rows = i32::try_from(sheet_rows).unwrap_or(i32::MAX);
                let sheet_cols = i32::try_from(sheet_cols).unwrap_or(i32::MAX);
                if resolved.row_start < 0
                    || resolved.col_start < 0
                    || resolved.row_end >= sheet_rows
                    || resolved.col_end >= sheet_cols
                {
                    // Out-of-bounds ranges must evaluate via per-cell access so `#REF!` can be
                    // surfaced. Don't build a cache that would otherwise treat them as empty/NaN.
                    continue;
                }
                if resolved.rows() > crate::bytecode::runtime::BYTECODE_SPARSE_RANGE_ROW_THRESHOLD {
                    // Avoid allocating huge columnar buffers for sparse ranges like `A:A`. The
                    // bytecode runtime can compute aggregates over these ranges by iterating the
                    // stored (non-implicit-blank) cells instead.
                    continue;
                }
                let cells = i64::from(resolved.rows())
                    .checked_mul(i64::from(resolved.cols()))
                    .unwrap_or(i64::MAX);
                if cells > BYTECODE_MAX_RANGE_CELLS {
                    // Avoid allocating enormous columnar buffers for wide ranges where the total
                    // cell count would exceed the runtime materialization limit. Aggregate
                    // functions can still evaluate these ranges efficiently via sparse iteration
                    // (`Grid::iter_cells_on_sheet`) without allocating a dense cache.
                    continue;
                }
                for col in resolved.col_start..=resolved.col_end {
                    row_ranges_by_col[key.sheet]
                        .entry(col)
                        .or_default()
                        .push((resolved.row_start, resolved.row_end));
                }
            }

            for multi in &bc.program.multi_range_refs {
                for area in multi.areas.iter() {
                    let sheet_id = match &area.sheet {
                        bytecode::SheetId::Local(sheet_id) => *sheet_id,
                        // External workbook ranges cannot be cached in the columnar local-sheet
                        // buffers, and we don't know their true dimensions anyway.
                        bytecode::SheetId::External(_) => continue,
                    };
                    if sheet_id >= sheet_count {
                        continue;
                    }
                    let resolved = area.range.resolve(base);
                    let (sheet_rows, sheet_cols) = snapshot
                        .sheet_dimensions
                        .get(sheet_id)
                        .copied()
                        .unwrap_or((0, 0));
                    let sheet_rows = i32::try_from(sheet_rows).unwrap_or(i32::MAX);
                    let sheet_cols = i32::try_from(sheet_cols).unwrap_or(i32::MAX);
                    if resolved.row_start < 0
                        || resolved.col_start < 0
                        || resolved.row_end >= sheet_rows
                        || resolved.col_end >= sheet_cols
                    {
                        // Out-of-bounds ranges must evaluate via per-cell access so `#REF!` can be
                        // surfaced. Don't build a cache that would otherwise treat them as empty/NaN.
                        continue;
                    }
                    if resolved.rows()
                        > crate::bytecode::runtime::BYTECODE_SPARSE_RANGE_ROW_THRESHOLD
                    {
                        // Avoid allocating huge columnar buffers for sparse ranges like `A:A`. The
                        // bytecode runtime can compute aggregates over these ranges by iterating the
                        // stored (non-implicit-blank) cells instead.
                        continue;
                    }
                    let cells = i64::from(resolved.rows())
                        .checked_mul(i64::from(resolved.cols()))
                        .unwrap_or(i64::MAX);
                    if cells > BYTECODE_MAX_RANGE_CELLS {
                        // Avoid allocating enormous columnar buffers for wide ranges where the
                        // total cell count would exceed the runtime materialization limit.
                        continue;
                    }
                    for col in resolved.col_start..=resolved.col_end {
                        row_ranges_by_col[sheet_id]
                            .entry(col)
                            .or_default()
                            .push((resolved.row_start, resolved.row_end));
                    }
                }
            }
        }

        // If no formulas in this batch require columnar caching, avoid scanning the entire sheet's
        // stored values. This is a critical fast path for deep dependency chains where we may
        // evaluate many single-cell levels (each with no range args).
        if row_ranges_by_col.iter().all(|cols| cols.is_empty()) {
            return Self {
                by_sheet: vec![HashMap::new(); sheet_count],
            };
        }

        fn apply_value(seg: &mut BytecodeColumnSegment, value: &Value, row: i32) {
            match value {
                Value::Number(n) => seg.values[(row - seg.row_start) as usize] = *n,
                Value::Blank => {}
                Value::Error(_)
                | Value::Reference(_)
                | Value::ReferenceUnion(_)
                | Value::Array(_)
                | Value::Lambda(_)
                | Value::Spill { .. } => {
                    seg.blocked_rows_strict.push(row);
                    seg.blocked_rows_ignore_nonnumeric.push(row);
                }
                // Rich values (Entity/Record) should behave like text: ignored by most aggregates,
                // but they disqualify strict-numeric column slices.
                Value::Entity(_) | Value::Record(_) => seg.blocked_rows_strict.push(row),
                // Excel ignores logical/text values in references for most aggregates, so allow
                // IgnoreNonNumeric SIMD slices.
                _ => seg.blocked_rows_strict.push(row),
            }
        }

        let has_provider = snapshot.external_value_provider.is_some();
        let mut by_sheet: Vec<HashMap<i32, BytecodeColumn>> = Vec::with_capacity(sheet_count);
        for sheet_id in 0..sheet_count {
            let mut cols: HashMap<i32, BytecodeColumn> = HashMap::new();
            for (col, ranges) in row_ranges_by_col[sheet_id].iter() {
                let mut merged = ranges.clone();
                merged.sort_unstable();

                let mut segments: Vec<(i32, i32)> = Vec::new();
                let mut cur_start = merged[0].0;
                let mut cur_end = merged[0].1;
                for &(row_start, row_end) in &merged[1..] {
                    if row_start <= cur_end.saturating_add(1) {
                        cur_end = cur_end.max(row_end);
                    } else {
                        segments.push((cur_start, cur_end));
                        cur_start = row_start;
                        cur_end = row_end;
                    }
                }
                segments.push((cur_start, cur_end));

                let mut col_segments: Vec<BytecodeColumnSegment> =
                    Vec::with_capacity(segments.len());

                let sheet_name = snapshot
                    .sheet_keys_by_id
                    .get(sheet_id)
                    .and_then(|s| s.as_deref());
                let provider = snapshot.external_value_provider.as_ref();

                for (row_start, row_end) in segments {
                    debug_assert!(row_start >= 0);
                    debug_assert!(row_end >= row_start);
                    let len = (row_end - row_start + 1) as usize;
                    let mut segment = BytecodeColumnSegment {
                        row_start,
                        values: vec![f64::NAN; len],
                        blocked_rows_strict: Vec::new(),
                        blocked_rows_ignore_nonnumeric: Vec::new(),
                    };

                    // When there is no external value provider, treat missing cells as blank and
                    // avoid scanning the entire row window just to do HashMap lookups. We'll fill
                    // from `snapshot.values` in a second pass after all segments are allocated.
                    if has_provider {
                        for row in row_start..=row_end {
                            let addr = CellAddr {
                                row: row as u32,
                                col: (*col) as u32,
                            };
                            if let Some(v) = snapshot.values.get(&CellKey {
                                sheet: sheet_id,
                                addr,
                            }) {
                                apply_value(&mut segment, v, row);
                                continue;
                            }

                            if let (Some(provider), Some(sheet_name)) = (provider, sheet_name) {
                                if let Some(v) = provider.get(sheet_name, addr) {
                                    apply_value(&mut segment, &v, row);
                                }
                            }
                        }
                    }

                    segment.blocked_rows_strict.sort_unstable();
                    segment.blocked_rows_strict.dedup();
                    segment.blocked_rows_ignore_nonnumeric.sort_unstable();
                    segment.blocked_rows_ignore_nonnumeric.dedup();

                    col_segments.push(segment);
                }

                cols.insert(
                    *col,
                    BytecodeColumn {
                        segments: col_segments,
                    },
                );
            }
            by_sheet.push(cols);
        }

        if !has_provider {
            for (key, value) in &snapshot.values {
                if matches!(value, Value::Blank) {
                    continue;
                }
                let Some(sheet_cols) = by_sheet.get_mut(key.sheet) else {
                    continue;
                };
                let Ok(col) = i32::try_from(key.addr.col) else {
                    continue;
                };
                let Some(column) = sheet_cols.get_mut(&col) else {
                    continue;
                };
                let Ok(row) = i32::try_from(key.addr.row) else {
                    continue;
                };
                let idx = column.segments.partition_point(|seg| seg.row_end() < row);
                if idx >= column.segments.len() {
                    continue;
                }
                let seg = &mut column.segments[idx];
                if row < seg.row_start || row > seg.row_end() {
                    continue;
                }
                apply_value(seg, value, row);
            }

            for sheet_cols in &mut by_sheet {
                for col in sheet_cols.values_mut() {
                    for seg in &mut col.segments {
                        seg.blocked_rows_strict.sort_unstable();
                        seg.blocked_rows_strict.dedup();
                        seg.blocked_rows_ignore_nonnumeric.sort_unstable();
                        seg.blocked_rows_ignore_nonnumeric.dedup();
                    }
                }
            }
        }

        Self { by_sheet }
    }
}

fn has_blocked_row(blocked_rows: &[i32], row_start: i32, row_end: i32) -> bool {
    if blocked_rows.is_empty() {
        return false;
    }
    let idx = blocked_rows.partition_point(|r| *r < row_start);
    blocked_rows.get(idx).is_some_and(|r| *r <= row_end)
}

struct EngineBytecodeGrid<'a> {
    snapshot: &'a Snapshot,
    sheet_id: SheetId,
    cols: &'a HashMap<i32, BytecodeColumn>,
    cols_by_sheet: &'a [HashMap<i32, BytecodeColumn>],
    slice_mode: ColumnSliceMode,
    trace: Option<&'a Mutex<crate::eval::DependencyTrace>>,
}

impl<'a> EngineBytecodeGrid<'a> {
    fn column_slice_impl(
        &self,
        cols: &'a HashMap<i32, BytecodeColumn>,
        col: i32,
        row_start: i32,
        row_end: i32,
        slice_mode: ColumnSliceMode,
    ) -> Option<&'a [f64]> {
        // Bounds checking (sheet dimensions, col range) must be done by the caller since this
        // helper is used for both the current sheet and sheet-qualified references.
        if col < 0 || row_start < 0 || row_end < 0 || row_start > row_end {
            return None;
        }
        let data = cols.get(&col)?;
        let idx = data
            .segments
            .partition_point(|seg| seg.row_end() < row_start);
        let seg = data.segments.get(idx)?;
        if row_start < seg.row_start || row_end > seg.row_end() {
            return None;
        }

        let blocked_rows = match slice_mode {
            ColumnSliceMode::StrictNumeric => &seg.blocked_rows_strict,
            ColumnSliceMode::IgnoreNonNumeric => &seg.blocked_rows_ignore_nonnumeric,
        };
        if has_blocked_row(blocked_rows, row_start, row_end) {
            return None;
        }

        let start = (row_start - seg.row_start) as usize;
        let end = (row_end - seg.row_start) as usize;
        if end >= seg.values.len() {
            return None;
        }
        Some(&seg.values[start..=end])
    }
}

impl bytecode::grid::Grid for EngineBytecodeGrid<'_> {
    fn get_value(&self, coord: bytecode::CellCoord) -> bytecode::Value {
        self.get_value_on_sheet(&bytecode::SheetId::Local(self.sheet_id), coord)
    }

    fn sheet_order_index(&self, sheet_id: usize) -> Option<usize> {
        let idx = *self.snapshot.tab_index_by_sheet_id.get(sheet_id)?;
        (idx != usize::MAX).then_some(idx)
    }

    fn external_sheet_order(&self, workbook: &str) -> Option<Vec<String>> {
        let provider = self.snapshot.external_value_provider.as_ref()?;
        provider.sheet_order(workbook).or_else(|| {
            provider
                .workbook_sheet_names(workbook)
                .map(|names| names.as_ref().to_vec())
        })
    }

    fn get_value_on_sheet(
        &self,
        sheet: &bytecode::SheetId,
        coord: bytecode::CellCoord,
    ) -> bytecode::Value {
        match sheet {
            bytecode::SheetId::Local(sheet_id) => {
                let sheet_id = *sheet_id;
                if !self.snapshot.sheet_exists(sheet_id) {
                    return bytecode::Value::Error(bytecode::ErrorKind::Ref);
                }

                let (rows, cols) = self.bounds_on_sheet(sheet);
                if coord.row < 0 || coord.col < 0 || coord.row >= rows || coord.col >= cols {
                    return bytecode::Value::Error(bytecode::ErrorKind::Ref);
                }
                let addr = CellAddr {
                    row: coord.row as u32,
                    col: coord.col as u32,
                };

                self.snapshot
                    .values
                    .get(&CellKey {
                        sheet: sheet_id,
                        addr,
                    })
                    .map(engine_value_to_bytecode)
                    .or_else(|| {
                        let provider = self.snapshot.external_value_provider.as_ref()?;
                        let sheet_name =
                            self.snapshot.sheet_keys_by_id.get(sheet_id)?.as_deref()?;
                        provider
                            .get(sheet_name, addr)
                            .as_ref()
                            .map(engine_value_to_bytecode)
                    })
                    .unwrap_or(bytecode::Value::Empty)
            }
            bytecode::SheetId::External(sheet_key) => {
                // External workbooks do not expose dimensions via the provider interface.
                // Treat bounds as unknown/valid (within Excel's fixed max grid) and rely on the
                // provider returning `None` to surface `#REF!`.
                if coord.row < 0
                    || coord.col < 0
                    || coord.row >= EXCEL_MAX_ROWS_I32
                    || coord.col >= EXCEL_MAX_COLS_I32
                {
                    return bytecode::Value::Error(bytecode::ErrorKind::Ref);
                }
                let addr = CellAddr {
                    row: coord.row as u32,
                    col: coord.col as u32,
                };

                // Bytecode dependency tracing historically ignored external workbook references
                // because the internal dependency graph cannot represent them. The engine now uses
                // these traces for auditing (and optional explicit invalidation), so record the
                // external dereference here.
                if let Some(trace) = self.trace {
                    let reference = crate::functions::Reference {
                        sheet_id: crate::functions::SheetId::External(sheet_key.to_string()),
                        start: addr,
                        end: addr,
                    };
                    let mut guard = match trace.lock() {
                        Ok(g) => g,
                        Err(poisoned) => poisoned.into_inner(),
                    };
                    guard.record_reference(reference);
                }

                let Some(provider) = self.snapshot.external_value_provider.as_ref() else {
                    return bytecode::Value::Error(bytecode::ErrorKind::Ref);
                };
                provider
                    .get(sheet_key, addr)
                    .as_ref()
                    .map(engine_value_to_bytecode)
                    .unwrap_or(bytecode::Value::Error(bytecode::ErrorKind::Ref))
            }
        }
    }

    fn record_reference(&self, sheet: usize, start: bytecode::CellCoord, end: bytecode::CellCoord) {
        // Dependency tracing is only used for the engine's internal dependency graph, which does
        // not represent external workbooks yet. Ignore references into synthetic external sheets.
        if !self.snapshot.sheet_exists(sheet) {
            return;
        }
        let Some(trace) = self.trace else {
            return;
        };
        let (Ok(start_row), Ok(start_col), Ok(end_row), Ok(end_col)) = (
            u32::try_from(start.row),
            u32::try_from(start.col),
            u32::try_from(end.row),
            u32::try_from(end.col),
        ) else {
            return;
        };

        // Dependency tracing is only used for the engine's internal dependency graph, which does
        // not represent external workbooks yet. Bytecode evaluation currently only supports local
        // sheet ids, so record as local precedents.
        let reference = crate::functions::Reference {
            sheet_id: crate::functions::SheetId::Local(sheet),
            start: CellAddr {
                row: start_row,
                col: start_col,
            },
            end: CellAddr {
                row: end_row,
                col: end_col,
            },
        };

        let mut guard = match trace.lock() {
            Ok(g) => g,
            Err(poisoned) => poisoned.into_inner(),
        };
        guard.record_reference(reference);
    }

    fn record_reference_on_sheet(
        &self,
        sheet: &bytecode::SheetId,
        start: bytecode::CellCoord,
        end: bytecode::CellCoord,
    ) {
        match sheet {
            bytecode::SheetId::Local(sheet_id) => self.record_reference(*sheet_id, start, end),
            bytecode::SheetId::External(sheet_key) => {
                let Some(trace) = self.trace else {
                    return;
                };
                let (Ok(start_row), Ok(start_col), Ok(end_row), Ok(end_col)) = (
                    u32::try_from(start.row),
                    u32::try_from(start.col),
                    u32::try_from(end.row),
                    u32::try_from(end.col),
                ) else {
                    return;
                };
                let reference = crate::functions::Reference {
                    sheet_id: crate::functions::SheetId::External(sheet_key.to_string()),
                    start: CellAddr {
                        row: start_row,
                        col: start_col,
                    },
                    end: CellAddr {
                        row: end_row,
                        col: end_col,
                    },
                };
                let mut guard = match trace.lock() {
                    Ok(g) => g,
                    Err(poisoned) => poisoned.into_inner(),
                };
                guard.record_reference(reference);
            }
        }
    }

    fn sheet_id(&self) -> usize {
        self.sheet_id
    }

    fn column_slice(&self, col: i32, row_start: i32, row_end: i32) -> Option<&[f64]> {
        let (rows, cols) = self.bounds();
        if col < 0
            || col >= cols
            || row_start < 0
            || row_end < 0
            || row_start > row_end
            || row_end >= rows
        {
            return None;
        }
        self.column_slice_impl(self.cols, col, row_start, row_end, self.slice_mode)
    }

    fn column_slice_strict_numeric(
        &self,
        col: i32,
        row_start: i32,
        row_end: i32,
    ) -> Option<&[f64]> {
        let (rows, cols) = self.bounds();
        if col < 0
            || col >= cols
            || row_start < 0
            || row_end < 0
            || row_start > row_end
            || row_end >= rows
        {
            return None;
        }
        self.column_slice_impl(
            self.cols,
            col,
            row_start,
            row_end,
            ColumnSliceMode::StrictNumeric,
        )
    }

    fn iter_cells(
        &self,
    ) -> Option<Box<dyn Iterator<Item = (bytecode::CellCoord, bytecode::Value)> + '_>> {
        self.iter_cells_on_sheet(&bytecode::SheetId::Local(self.sheet_id))
    }

    fn iter_cells_on_sheet(
        &self,
        sheet: &bytecode::SheetId,
    ) -> Option<Box<dyn Iterator<Item = (bytecode::CellCoord, bytecode::Value)> + '_>> {
        let sheet_id = match sheet {
            bytecode::SheetId::Local(id) => *id,
            bytecode::SheetId::External(_) => return None,
        };
        // When external values are provided out-of-band, we cannot safely iterate just the
        // snapshot's stored cells because we'd miss provider-backed cells that should contribute
        // to aggregates.
        if self.snapshot.external_value_provider.is_some() {
            return None;
        }
        if !self.snapshot.sheet_exists(sheet_id) {
            return None;
        }

        let start = CellKey {
            sheet: sheet_id,
            addr: CellAddr { row: 0, col: 0 },
        };
        let end = CellKey {
            sheet: sheet_id,
            addr: CellAddr {
                row: u32::MAX,
                col: u32::MAX,
            },
        };

        Some(Box::new(
            self.snapshot
                .ordered_cells
                .range(start..=end)
                .filter_map(|k| {
                    let value = self.snapshot.values.get(k)?;
                    Some((
                        bytecode::CellCoord {
                            row: k.addr.row as i32,
                            col: k.addr.col as i32,
                        },
                        engine_value_to_bytecode(value),
                    ))
                }),
        ))
    }

    fn column_slice_on_sheet(
        &self,
        sheet: &bytecode::SheetId,
        col: i32,
        row_start: i32,
        row_end: i32,
    ) -> Option<&[f64]> {
        let sheet_id = match sheet {
            bytecode::SheetId::Local(id) => *id,
            bytecode::SheetId::External(_) => return None,
        };
        if !self.snapshot.sheet_exists(sheet_id) {
            return None;
        }
        let (rows, cols) = self.bounds_on_sheet(sheet);
        if col < 0
            || col >= cols
            || row_start < 0
            || row_end < 0
            || row_start > row_end
            || row_end >= rows
        {
            return None;
        }
        let sheet_cols = self.cols_by_sheet.get(sheet_id)?;
        self.column_slice_impl(sheet_cols, col, row_start, row_end, self.slice_mode)
    }

    fn column_slice_on_sheet_strict_numeric(
        &self,
        sheet: &bytecode::SheetId,
        col: i32,
        row_start: i32,
        row_end: i32,
    ) -> Option<&[f64]> {
        let sheet_id = match sheet {
            bytecode::SheetId::Local(id) => *id,
            bytecode::SheetId::External(_) => return None,
        };
        if !self.snapshot.sheet_exists(sheet_id) {
            return None;
        }
        let (rows, cols) = self.bounds_on_sheet(sheet);
        if col < 0
            || col >= cols
            || row_start < 0
            || row_end < 0
            || row_start > row_end
            || row_end >= rows
        {
            return None;
        }
        let sheet_cols = self.cols_by_sheet.get(sheet_id)?;
        self.column_slice_impl(
            sheet_cols,
            col,
            row_start,
            row_end,
            ColumnSliceMode::StrictNumeric,
        )
    }

    fn bounds(&self) -> (i32, i32) {
        let (rows, cols) = self
            .snapshot
            .sheet_dimensions
            .get(self.sheet_id)
            .copied()
            .unwrap_or((0, 0));
        let rows = i32::try_from(rows).unwrap_or(i32::MAX);
        let cols = i32::try_from(cols).unwrap_or(i32::MAX);
        (rows, cols)
    }

    fn bounds_on_sheet(&self, sheet: &bytecode::SheetId) -> (i32, i32) {
        match sheet {
            bytecode::SheetId::Local(sheet_id) => {
                let sheet_id = *sheet_id;
                if !self.snapshot.sheet_exists(sheet_id) {
                    return (0, 0);
                }
                let (rows, cols) = self
                    .snapshot
                    .sheet_dimensions
                    .get(sheet_id)
                    .copied()
                    .unwrap_or((0, 0));
                let rows = i32::try_from(rows).unwrap_or(i32::MAX);
                let cols = i32::try_from(cols).unwrap_or(i32::MAX);
                (rows, cols)
            }
            bytecode::SheetId::External(_) => (EXCEL_MAX_ROWS_I32, EXCEL_MAX_COLS_I32),
        }
    }

    fn resolve_sheet_name(&self, name: &str) -> Option<usize> {
        // Excel resolves sheet names case-insensitively across Unicode using compatibility
        // normalization (NFKC). Ensure bytecode runtime sheet-name lookups (e.g. INDIRECT, SHEET)
        // match the engine's canonical sheet-key semantics.
        let key = Workbook::sheet_key(name);
        let local = self
            .snapshot
            .sheet_display_name_to_id
            .get(&key)
            .copied()
            .or_else(|| self.snapshot.sheet_key_to_id.get(&key).copied());
        if local.is_some() {
            return local;
        }
        None
    }

    fn spill_origin(&self, sheet_id: &bytecode::SheetId, addr: CellAddr) -> Option<CellAddr> {
        match sheet_id {
            bytecode::SheetId::Local(sheet_id) => self.snapshot.spill_origin(*sheet_id, addr),
            bytecode::SheetId::External(_) => None,
        }
    }

    fn spill_range(
        &self,
        sheet_id: &bytecode::SheetId,
        origin: CellAddr,
    ) -> Option<(CellAddr, CellAddr)> {
        match sheet_id {
            bytecode::SheetId::Local(sheet_id) => self.snapshot.spill_range(*sheet_id, origin),
            bytecode::SheetId::External(_) => None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BytecodeLocalBindingKind {
    /// Scalar value (number/text/bool/empty/error).
    Scalar,
    /// Single-cell reference value (e.g. `A1` / `A1:A1`).
    ///
    /// This is treated as scalar-safe in contexts that do not allow spilling ranges (the bytecode
    /// compiler will apply implicit intersection when such locals are consumed in scalar context),
    /// but can also participate in "reference-like" positions (e.g. XLOOKUP/XMATCH lookup vectors).
    RefSingle,
    /// Range reference value.
    Range,
    /// Array literal constant value.
    ArrayLiteral,
}

fn bytecode_expr_is_eligible(expr: &bytecode::Expr) -> bool {
    let mut lexical_scopes: Vec<HashMap<Arc<str>, BytecodeLocalBindingKind>> = Vec::new();
    // Top-level formulas use dynamic reference dereference, so range references are eligible and
    // may spill (e.g. `=A1:A3` / `=A1:A3+1`).
    //
    // Array literals are also eligible in top-level contexts now that the bytecode backend can
    // represent and spill mixed-type arrays (e.g. `={1,2;3,4}`).
    bytecode_expr_is_eligible_inner(expr, true, true, &mut lexical_scopes)
}

fn bytecode_expr_first_unsupported_function(expr: &bytecode::Expr) -> Option<Arc<str>> {
    match expr {
        bytecode::Expr::FuncCall {
            func: bytecode::ast::Function::Unknown(name),
            ..
        } => Some(name.clone()),
        bytecode::Expr::FuncCall { args, .. } => args
            .iter()
            .find_map(|arg| bytecode_expr_first_unsupported_function(arg)),
        bytecode::Expr::SpillRange(inner) => bytecode_expr_first_unsupported_function(inner),
        bytecode::Expr::Unary { expr, .. } => bytecode_expr_first_unsupported_function(expr),
        bytecode::Expr::Binary { left, right, .. } => {
            bytecode_expr_first_unsupported_function(left)
                .or_else(|| bytecode_expr_first_unsupported_function(right))
        }
        bytecode::Expr::Lambda { body, .. } => bytecode_expr_first_unsupported_function(body),
        bytecode::Expr::Call { callee, args } => bytecode_expr_first_unsupported_function(callee)
            .or_else(|| {
                args.iter()
                    .find_map(|arg| bytecode_expr_first_unsupported_function(arg))
            }),
        bytecode::Expr::Literal(_)
        | bytecode::Expr::CellRef(_)
        | bytecode::Expr::RangeRef(_)
        | bytecode::Expr::MultiRangeRef(_)
        | bytecode::Expr::NameRef(_) => None,
    }
}

fn bytecode_expr_within_grid_limits(
    expr: &bytecode::Expr,
    origin: bytecode::CellCoord,
    origin_sheet_bounds: (i32, i32),
    sheet_bounds: &mut impl FnMut(&bytecode::SheetId) -> (i32, i32),
) -> Result<(), BytecodeCompileReason> {
    bytecode_expr_within_grid_limits_inner(
        expr,
        origin,
        origin_sheet_bounds,
        sheet_bounds,
        BYTECODE_MAX_RANGE_CELLS,
    )
}

fn bytecode_expr_within_grid_limits_inner(
    expr: &bytecode::Expr,
    origin: bytecode::CellCoord,
    origin_sheet_bounds: (i32, i32),
    sheet_bounds: &mut impl FnMut(&bytecode::SheetId) -> (i32, i32),
    max_range_cells: i64,
) -> Result<(), BytecodeCompileReason> {
    // Bytecode coordinate math uses signed 32-bit indices. For now we enforce the engine's
    // hard bounds:
    // - rows are limited to `i32::MAX` (so max row index is `i32::MAX - 1`)
    // - cols are limited to Excel's 16,384-column maximum (`EXCEL_MAX_COLS`)
    //
    // Note: We intentionally do *not* validate against per-sheet configured dimensions here.
    // References that are out-of-bounds for a particular sheet are still valid formulas that
    // evaluate to `#REF!`; the bytecode VM handles those cases at runtime.
    let max_rows = i32::MAX;
    let max_cols = EXCEL_MAX_COLS_I32;
    match expr {
        bytecode::Expr::Literal(_) => Ok(()),
        bytecode::Expr::CellRef(r) => {
            let coord = r.resolve(origin);
            if coord.row >= 0 && coord.col >= 0 && coord.row < max_rows && coord.col < max_cols {
                Ok(())
            } else {
                Err(BytecodeCompileReason::ExceedsGridLimits)
            }
        }
        bytecode::Expr::RangeRef(r) => {
            let resolved = r.resolve(origin);
            if resolved.row_start < 0
                || resolved.col_start < 0
                || resolved.row_end >= max_rows
                || resolved.col_end >= max_cols
            {
                return Err(BytecodeCompileReason::ExceedsGridLimits);
            }
            let cells = (resolved.rows() as i64) * (resolved.cols() as i64);
            if cells <= max_range_cells {
                Ok(())
            } else {
                Err(BytecodeCompileReason::ExceedsRangeCellLimit)
            }
        }
        bytecode::Expr::MultiRangeRef(r) => {
            let mut total: i64 = 0;
            for area in r.areas.iter() {
                let resolved = area.range.resolve(origin);
                if resolved.row_start < 0
                    || resolved.col_start < 0
                    || resolved.row_end >= max_rows
                    || resolved.col_end >= max_cols
                {
                    return Err(BytecodeCompileReason::ExceedsGridLimits);
                }
                let cells = (resolved.rows() as i64) * (resolved.cols() as i64);
                total = total.saturating_add(cells);
                if total > max_range_cells {
                    return Err(BytecodeCompileReason::ExceedsRangeCellLimit);
                }
            }
            Ok(())
        }
        bytecode::Expr::SpillRange(inner) => bytecode_expr_within_grid_limits_inner(
            inner,
            origin,
            origin_sheet_bounds,
            sheet_bounds,
            max_range_cells,
        ),
        bytecode::Expr::NameRef(_) => Ok(()),
        bytecode::Expr::Unary { op, expr } => match op {
            bytecode::ast::UnaryOp::Plus | bytecode::ast::UnaryOp::Neg => {
                bytecode_expr_within_grid_limits_inner(
                    expr,
                    origin,
                    origin_sheet_bounds,
                    sheet_bounds,
                    max_range_cells,
                )
            }
            bytecode::ast::UnaryOp::ImplicitIntersection => match expr.as_ref() {
                // Implicit intersection only dereferences at most one cell from a range, so it
                // doesn't require allocating columnar buffers proportional to the range size.
                // Skip the range-cell-count limit here while still validating grid bounds.
                bytecode::Expr::RangeRef(r) => {
                    let resolved = r.resolve(origin);
                    if resolved.row_start < 0
                        || resolved.col_start < 0
                        || resolved.row_end >= max_rows
                        || resolved.col_end >= max_cols
                    {
                        Err(BytecodeCompileReason::ExceedsGridLimits)
                    } else {
                        Ok(())
                    }
                }
                bytecode::Expr::MultiRangeRef(r) => {
                    for area in r.areas.iter() {
                        let resolved = area.range.resolve(origin);
                        if resolved.row_start < 0
                            || resolved.col_start < 0
                            || resolved.row_end >= max_rows
                            || resolved.col_end >= max_cols
                        {
                            return Err(BytecodeCompileReason::ExceedsGridLimits);
                        }
                    }
                    Ok(())
                }
                _ => bytecode_expr_within_grid_limits_inner(
                    expr,
                    origin,
                    origin_sheet_bounds,
                    sheet_bounds,
                    max_range_cells,
                ),
            },
        },
        bytecode::Expr::Binary { left, right, .. } => {
            bytecode_expr_within_grid_limits_inner(
                left,
                origin,
                origin_sheet_bounds,
                sheet_bounds,
                max_range_cells,
            )?;
            bytecode_expr_within_grid_limits_inner(
                right,
                origin,
                origin_sheet_bounds,
                sheet_bounds,
                max_range_cells,
            )?;
            Ok(())
        }
        bytecode::Expr::FuncCall { func, args } => {
            use bytecode::ast::Function;
            match func {
                // ROW/COLUMN treat whole-row/whole-column references as 1-D arrays. Use the real
                // sheet bounds to apply an accurate cell-count limit for these special cases.
                Function::Row | Function::Column => {
                    for (idx, arg) in args.iter().enumerate() {
                        if idx == 0 {
                            match arg {
                                bytecode::Expr::RangeRef(r) => {
                                    let resolved = r.resolve(origin);
                                    if resolved.row_start < 0
                                        || resolved.col_start < 0
                                        || resolved.row_end >= max_rows
                                        || resolved.col_end >= max_cols
                                    {
                                        return Err(BytecodeCompileReason::ExceedsGridLimits);
                                    }

                                    let (sheet_rows, sheet_cols) = origin_sheet_bounds;
                                    let spans_all_cols = resolved.col_start == 0
                                        && resolved.col_end == sheet_cols.saturating_sub(1);
                                    let spans_all_rows = resolved.row_start == 0
                                        && resolved.row_end == sheet_rows.saturating_sub(1);

                                    let cells = if spans_all_cols || spans_all_rows {
                                        match func {
                                            Function::Row => i64::from(resolved.rows()),
                                            Function::Column => i64::from(resolved.cols()),
                                            _ => unreachable!("matched above"),
                                        }
                                    } else {
                                        (i64::from(resolved.rows())) * (i64::from(resolved.cols()))
                                    };
                                    if cells > BYTECODE_MAX_RANGE_CELLS {
                                        return Err(BytecodeCompileReason::ExceedsRangeCellLimit);
                                    }
                                    continue;
                                }
                                bytecode::Expr::MultiRangeRef(r) => {
                                    match r.areas.as_ref() {
                                        // `ROW()`/`COLUMN()` treat empty unions as `#REF!`; the VM
                                        // will short-circuit without allocating.
                                        [] => continue,
                                        [only] => {
                                            let resolved = only.range.resolve(origin);
                                            if resolved.row_start < 0
                                                || resolved.col_start < 0
                                                || resolved.row_end >= max_rows
                                                || resolved.col_end >= max_cols
                                            {
                                                return Err(
                                                    BytecodeCompileReason::ExceedsGridLimits,
                                                );
                                            }

                                            let (sheet_rows, sheet_cols) =
                                                sheet_bounds(&only.sheet);
                                            // If the range is out-of-bounds for the referenced
                                            // sheet, evaluation returns `#REF!` without attempting
                                            // to allocate. Do not reject these formulas due to
                                            // cell-count limits.
                                            if resolved.row_end >= sheet_rows
                                                || resolved.col_end >= sheet_cols
                                                || sheet_rows <= 0
                                                || sheet_cols <= 0
                                            {
                                                continue;
                                            }

                                            let spans_all_cols = resolved.col_start == 0
                                                && resolved.col_end == sheet_cols.saturating_sub(1);
                                            let spans_all_rows = resolved.row_start == 0
                                                && resolved.row_end == sheet_rows.saturating_sub(1);

                                            let cells = if spans_all_cols || spans_all_rows {
                                                match func {
                                                    Function::Row => i64::from(resolved.rows()),
                                                    Function::Column => i64::from(resolved.cols()),
                                                    _ => unreachable!("matched above"),
                                                }
                                            } else {
                                                (i64::from(resolved.rows()))
                                                    * (i64::from(resolved.cols()))
                                            };
                                            if cells > BYTECODE_MAX_RANGE_CELLS {
                                                return Err(
                                                    BytecodeCompileReason::ExceedsRangeCellLimit,
                                                );
                                            }
                                            continue;
                                        }
                                        // Multi-area references return `#VALUE!` for ROW/COLUMN,
                                        // so they cannot allocate a large output array. Still
                                        // validate that each component range lies within the VM's
                                        // representable grid.
                                        areas => {
                                            for area in areas {
                                                let resolved = area.range.resolve(origin);
                                                if resolved.row_start < 0
                                                    || resolved.col_start < 0
                                                    || resolved.row_end >= max_rows
                                                    || resolved.col_end >= max_cols
                                                {
                                                    return Err(
                                                        BytecodeCompileReason::ExceedsGridLimits,
                                                    );
                                                }
                                            }
                                            continue;
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }

                        bytecode_expr_within_grid_limits_inner(
                            arg,
                            origin,
                            origin_sheet_bounds,
                            sheet_bounds,
                            max_range_cells,
                        )?;
                    }
                    Ok(())
                }
                // ROWS/COLUMNS only need reference bounds; they do not allocate buffers
                // proportional to the range size.
                Function::Rows | Function::Columns => {
                    for arg in args {
                        match arg {
                            bytecode::Expr::RangeRef(r) => {
                                let resolved = r.resolve(origin);
                                if resolved.row_start < 0
                                    || resolved.col_start < 0
                                    || resolved.row_end >= max_rows
                                    || resolved.col_end >= max_cols
                                {
                                    return Err(BytecodeCompileReason::ExceedsGridLimits);
                                }
                            }
                            bytecode::Expr::MultiRangeRef(r) => {
                                for area in r.areas.iter() {
                                    let resolved = area.range.resolve(origin);
                                    if resolved.row_start < 0
                                        || resolved.col_start < 0
                                        || resolved.row_end >= max_rows
                                        || resolved.col_end >= max_cols
                                    {
                                        return Err(BytecodeCompileReason::ExceedsGridLimits);
                                    }
                                }
                            }
                            _ => {
                                bytecode_expr_within_grid_limits_inner(
                                    arg,
                                    origin,
                                    origin_sheet_bounds,
                                    sheet_bounds,
                                    max_range_cells,
                                )?;
                            }
                        }
                    }
                    Ok(())
                }
                _ => {
                    for (arg_idx, arg) in args.iter().enumerate() {
                        let arg_limit = match func {
                            // Aggregates can iterate ranges without forcing dense materialization,
                            // so skip the range cell-count limit for reference-like arguments.
                            Function::Sum
                            | Function::SumIf
                            | Function::SumIfs
                            | Function::Average
                            | Function::AverageIf
                            | Function::AverageIfs
                            | Function::Min
                            | Function::MinIfs
                            | Function::Max
                            | Function::MaxIfs
                            | Function::Count
                            | Function::CountA
                            | Function::CountBlank
                            | Function::CountIf
                            | Function::CountIfs
                            | Function::SumProduct
                            | Function::And
                            | Function::Or
                            | Function::Xor => {
                                // Criteria arguments are scalar, so keep the default limit there.
                                match (func, arg_idx) {
                                    (Function::SumIf | Function::AverageIf, 1) => max_range_cells,
                                    (Function::CountIf, 1) => max_range_cells,
                                    (
                                        Function::SumIfs
                                        | Function::AverageIfs
                                        | Function::MinIfs
                                        | Function::MaxIfs,
                                        idx,
                                    ) if idx > 0 && idx % 2 == 0 => max_range_cells,
                                    (Function::CountIfs, idx) if idx % 2 == 1 => max_range_cells,
                                    _ => i64::MAX,
                                }
                            }
                            _ => max_range_cells,
                        };

                        bytecode_expr_within_grid_limits_inner(
                            arg,
                            origin,
                            origin_sheet_bounds,
                            sheet_bounds,
                            arg_limit,
                        )?;
                    }
                    Ok(())
                }
            }
        }
        bytecode::Expr::Lambda { body, .. } => bytecode_expr_within_grid_limits_inner(
            body,
            origin,
            origin_sheet_bounds,
            sheet_bounds,
            max_range_cells,
        ),
        bytecode::Expr::Call { callee, args } => {
            bytecode_expr_within_grid_limits_inner(
                callee,
                origin,
                origin_sheet_bounds,
                sheet_bounds,
                max_range_cells,
            )?;
            for arg in args {
                bytecode_expr_within_grid_limits_inner(
                    arg,
                    origin,
                    origin_sheet_bounds,
                    sheet_bounds,
                    max_range_cells,
                )?;
            }
            Ok(())
        }
    }
}

fn bytecode_expr_is_eligible_inner(
    expr: &bytecode::Expr,
    allow_range: bool,
    allow_array_literals: bool,
    lexical_scopes: &mut Vec<HashMap<Arc<str>, BytecodeLocalBindingKind>>,
) -> bool {
    fn local_binding_kind(
        scopes: &[HashMap<Arc<str>, BytecodeLocalBindingKind>],
        name: &Arc<str>,
    ) -> Option<BytecodeLocalBindingKind> {
        scopes
            .iter()
            .rev()
            .find_map(|scope| scope.get(name).copied())
    }

    fn infer_binding_kind(
        expr: &bytecode::Expr,
        scopes: &mut Vec<HashMap<Arc<str>, BytecodeLocalBindingKind>>,
    ) -> BytecodeLocalBindingKind {
        use bytecode::ast::{Function, UnaryOp};

        match expr {
            bytecode::Expr::Literal(v) => match v {
                bytecode::Value::Range(r) => {
                    if r.start == r.end {
                        BytecodeLocalBindingKind::RefSingle
                    } else {
                        BytecodeLocalBindingKind::Range
                    }
                }
                bytecode::Value::MultiRange(_) => BytecodeLocalBindingKind::Range,
                bytecode::Value::Array(_) => BytecodeLocalBindingKind::ArrayLiteral,
                _ => BytecodeLocalBindingKind::Scalar,
            },
            bytecode::Expr::CellRef(_) => BytecodeLocalBindingKind::RefSingle,
            bytecode::Expr::RangeRef(r) => {
                if r.start == r.end {
                    BytecodeLocalBindingKind::RefSingle
                } else {
                    BytecodeLocalBindingKind::Range
                }
            }
            bytecode::Expr::MultiRangeRef(_) | bytecode::Expr::SpillRange(_) => {
                BytecodeLocalBindingKind::Range
            }
            bytecode::Expr::NameRef(name) => {
                local_binding_kind(scopes, name).unwrap_or(BytecodeLocalBindingKind::Scalar)
            }
            bytecode::Expr::Unary { op, expr } => match op {
                UnaryOp::ImplicitIntersection => BytecodeLocalBindingKind::Scalar,
                UnaryOp::Plus | UnaryOp::Neg => match infer_binding_kind(expr, scopes) {
                    BytecodeLocalBindingKind::Scalar | BytecodeLocalBindingKind::RefSingle => {
                        BytecodeLocalBindingKind::Scalar
                    }
                    BytecodeLocalBindingKind::Range | BytecodeLocalBindingKind::ArrayLiteral => {
                        BytecodeLocalBindingKind::ArrayLiteral
                    }
                },
            },
            bytecode::Expr::Binary { left, right, .. } => {
                let left_kind = infer_binding_kind(left, scopes);
                let right_kind = infer_binding_kind(right, scopes);
                match (left_kind, right_kind) {
                    (
                        BytecodeLocalBindingKind::Scalar | BytecodeLocalBindingKind::RefSingle,
                        BytecodeLocalBindingKind::Scalar | BytecodeLocalBindingKind::RefSingle,
                    ) => BytecodeLocalBindingKind::Scalar,
                    _ => BytecodeLocalBindingKind::ArrayLiteral,
                }
            }
            bytecode::Expr::Lambda { .. } => BytecodeLocalBindingKind::Scalar,
            bytecode::Expr::Call { callee, .. } => match callee.as_ref() {
                bytecode::Expr::Lambda { body, .. } => infer_binding_kind(body, scopes),
                _ => BytecodeLocalBindingKind::Scalar,
            },
            bytecode::Expr::FuncCall {
                func: Function::Let,
                args,
            } => {
                if args.len() < 3 || args.len() % 2 == 0 {
                    return BytecodeLocalBindingKind::Scalar;
                }
                scopes.push(HashMap::new());
                for pair in args[..args.len() - 1].chunks_exact(2) {
                    let bytecode::Expr::NameRef(name) = &pair[0] else {
                        scopes.pop();
                        return BytecodeLocalBindingKind::Scalar;
                    };
                    let kind = infer_binding_kind(&pair[1], scopes);
                    scopes
                        .last_mut()
                        .expect("pushed scope")
                        .insert(name.clone(), kind);
                }
                let kind = infer_binding_kind(&args[args.len() - 1], scopes);
                scopes.pop();
                kind
            }
            // Most supported functions return scalars in the bytecode backend. A handful lift over
            // array/range inputs (e.g. ISBLANK) or can return dynamic arrays (e.g. ROW/ COLUMN).
            //
            // This kind inference is used to prevent LET locals from "smuggling" array results into
            // scalar-only bytecode contexts (e.g. ABS / CONCAT_OP), which would otherwise compile to
            // bytecode but produce incorrect `#SPILL!` errors at runtime.
            bytecode::Expr::FuncCall { func, args } => match func {
                // Field access (`A1.Price`) lifts over array/range bases and returns an array in
                // those cases. Treat it as array-producing when its base is range/array-like so LET
                // locals cannot smuggle dynamic arrays into scalar-only contexts.
                Function::FieldAccess => {
                    if args.len() != 2 {
                        return BytecodeLocalBindingKind::Scalar;
                    }
                    match infer_binding_kind(&args[0], scopes) {
                        BytecodeLocalBindingKind::Scalar | BytecodeLocalBindingKind::RefSingle => {
                            BytecodeLocalBindingKind::Scalar
                        }
                        BytecodeLocalBindingKind::Range
                        | BytecodeLocalBindingKind::ArrayLiteral => {
                            BytecodeLocalBindingKind::ArrayLiteral
                        }
                    }
                }
                // CHOOSE can return either scalars, ranges, or array literals depending on the
                // selected value argument. Use the "widest" kind across its value args so LET
                // locals cannot smuggle range/array results into scalar-only bytecode contexts.
                Function::Choose => {
                    let mut kind = BytecodeLocalBindingKind::Scalar;
                    // Skip the index argument: it is always a scalar in bytecode-eligible CHOOSE.
                    for arg in args.iter().skip(1) {
                        match infer_binding_kind(arg, scopes) {
                            BytecodeLocalBindingKind::Scalar => {}
                            BytecodeLocalBindingKind::RefSingle => {
                                if kind == BytecodeLocalBindingKind::Scalar {
                                    kind = BytecodeLocalBindingKind::RefSingle;
                                }
                            }
                            BytecodeLocalBindingKind::Range => {
                                kind = BytecodeLocalBindingKind::Range;
                            }
                            BytecodeLocalBindingKind::ArrayLiteral => {
                                kind = BytecodeLocalBindingKind::ArrayLiteral;
                                break;
                            }
                        }
                    }
                    kind
                }
                Function::XLookup => {
                    // XLOOKUP can spill in a few situations:
                    // - array/range `lookup_value` (vectorized lookup)
                    // - 2D `return_array` (row/column slice)
                    // - array/range `if_not_found` fallback (e.g. `{100;200}`)
                    //
                    // LET bindings are always validated in a "range + array literal" context, so
                    // we must conservatively tag potentially-spilling XLOOKUP expressions as array
                    // values. This prevents scalar-only bytecode contexts (e.g. ABS / CONCAT_OP)
                    // from incorrectly accepting LET locals that may spill at runtime.
                    let lookup_value_is_array = matches!(
                        args.get(0).map(|arg| infer_binding_kind(arg, scopes)),
                        Some(
                            BytecodeLocalBindingKind::Range
                                | BytecodeLocalBindingKind::ArrayLiteral
                        )
                    );

                    let return_array_is_2d_literal = matches!(
                        args.get(2),
                        Some(bytecode::Expr::Literal(bytecode::Value::Array(arr)))
                            if arr.rows > 1 && arr.cols > 1
                    );

                    let if_not_found_is_array = matches!(
                        args.get(3).map(|arg| infer_binding_kind(arg, scopes)),
                        Some(
                            BytecodeLocalBindingKind::ArrayLiteral
                                | BytecodeLocalBindingKind::Range
                        )
                    );

                    if lookup_value_is_array || return_array_is_2d_literal || if_not_found_is_array
                    {
                        BytecodeLocalBindingKind::ArrayLiteral
                    } else {
                        BytecodeLocalBindingKind::Scalar
                    }
                }
                // OFFSET/INDIRECT produce reference values (ranges). Treat them conservatively as
                // range-like so LET bindings cannot smuggle them into scalar-only bytecode
                // contexts.
                Function::Offset | Function::Indirect => BytecodeLocalBindingKind::Range,
                Function::Row
                | Function::Column
                | Function::IsError
                | Function::IsNa
                | Function::IsBlank
                | Function::IsNumber
                | Function::IsText
                | Function::IsLogical
                | Function::IsErr
                | Function::ErrorType
                | Function::N
                | Function::T
                | Function::Abs
                | Function::Int
                | Function::Round
                | Function::RoundUp
                | Function::RoundDown
                | Function::Mod
                | Function::Sign
                | Function::ConcatOp
                | Function::Not => {
                    let mut all_scalar = true;
                    for arg in args {
                        if !matches!(
                            infer_binding_kind(arg, scopes),
                            BytecodeLocalBindingKind::Scalar | BytecodeLocalBindingKind::RefSingle
                        ) {
                            all_scalar = false;
                            break;
                        }
                    }
                    if all_scalar {
                        BytecodeLocalBindingKind::Scalar
                    } else {
                        BytecodeLocalBindingKind::ArrayLiteral
                    }
                }
                _ => BytecodeLocalBindingKind::Scalar,
            },
        }
    }

    fn choose_index_is_guaranteed_scalar(
        expr: &bytecode::Expr,
        lexical_scopes: &mut Vec<HashMap<Arc<str>, BytecodeLocalBindingKind>>,
    ) -> bool {
        use bytecode::ast::{Function, UnaryOp};

        match expr {
            bytecode::Expr::Literal(v) => !matches!(
                v,
                bytecode::Value::Array(_)
                    | bytecode::Value::Range(_)
                    | bytecode::Value::MultiRange(_)
            ),
            bytecode::Expr::CellRef(_) => true,
            // Bare range values are not scalar indices (even if they resolve to a single cell).
            bytecode::Expr::RangeRef(_)
            | bytecode::Expr::MultiRangeRef(_)
            | bytecode::Expr::SpillRange(_) => false,
            bytecode::Expr::NameRef(name) => {
                matches!(
                    local_binding_kind(lexical_scopes, name),
                    Some(BytecodeLocalBindingKind::Scalar | BytecodeLocalBindingKind::RefSingle)
                )
            }
            bytecode::Expr::Unary { op, expr } => match op {
                UnaryOp::ImplicitIntersection => true,
                UnaryOp::Plus | UnaryOp::Neg => {
                    choose_index_is_guaranteed_scalar(expr, lexical_scopes)
                }
            },
            bytecode::Expr::Binary { left, right, .. } => {
                choose_index_is_guaranteed_scalar(left, lexical_scopes)
                    && choose_index_is_guaranteed_scalar(right, lexical_scopes)
            }
            // Lambda values are scalars (even if they will later coerce to `#VALUE!`).
            bytecode::Expr::Lambda { .. } => true,
            bytecode::Expr::Call { callee, args } => {
                // Conservatively assume the result is scalar only for direct lambda calls where
                // both the arguments and the lambda body are known to stay scalar.
                let bytecode::Expr::Lambda { params, body } = callee.as_ref() else {
                    return false;
                };

                if !args
                    .iter()
                    .all(|arg| choose_index_is_guaranteed_scalar(arg, lexical_scopes))
                {
                    return false;
                }

                lexical_scopes.push(HashMap::new());
                for p in params.iter() {
                    lexical_scopes
                        .last_mut()
                        .expect("pushed scope")
                        .insert(p.clone(), BytecodeLocalBindingKind::Scalar);
                }
                let ok = choose_index_is_guaranteed_scalar(body, lexical_scopes);
                lexical_scopes.pop();
                ok
            }
            bytecode::Expr::FuncCall { func, args } => match func {
                // These control-flow functions can return different expressions depending on runtime
                // values; treat the result as scalar only if all possible result branches are scalar.
                Function::If => match args.as_slice() {
                    [_, t, f] => {
                        choose_index_is_guaranteed_scalar(t, lexical_scopes)
                            && choose_index_is_guaranteed_scalar(f, lexical_scopes)
                    }
                    [_, t] => choose_index_is_guaranteed_scalar(t, lexical_scopes),
                    _ => true, // invalid IF => #VALUE! (scalar error)
                },
                Function::IfError | Function::IfNa => match args.as_slice() {
                    [a, b] => {
                        choose_index_is_guaranteed_scalar(a, lexical_scopes)
                            && choose_index_is_guaranteed_scalar(b, lexical_scopes)
                    }
                    _ => true,
                },
                Function::Ifs => {
                    if args.len() < 2 || args.len() % 2 != 0 {
                        return true;
                    }
                    // Only the value expressions affect the output kind (conditions always coerce to bool).
                    args.chunks_exact(2)
                        .all(|pair| choose_index_is_guaranteed_scalar(&pair[1], lexical_scopes))
                }
                Function::Switch => {
                    if args.len() < 3 {
                        return true;
                    }
                    let has_default = (args.len() - 1) % 2 != 0;
                    let pairs_end = if has_default {
                        args.len() - 1
                    } else {
                        args.len()
                    };
                    let pairs = &args[1..pairs_end];
                    if pairs.len() < 2 || pairs.len() % 2 != 0 {
                        return true;
                    }

                    // Result expressions must be scalar. Case values don't affect the output kind.
                    for pair in pairs.chunks_exact(2) {
                        if !choose_index_is_guaranteed_scalar(&pair[1], lexical_scopes) {
                            return false;
                        }
                    }
                    if has_default {
                        choose_index_is_guaranteed_scalar(&args[args.len() - 1], lexical_scopes)
                    } else {
                        true
                    }
                }
                Function::Choose => {
                    if args.len() < 2 || args.len() > 255 {
                        return true;
                    }
                    // Scalar output requires both a scalar index and scalar choices.
                    if !choose_index_is_guaranteed_scalar(&args[0], lexical_scopes) {
                        return false;
                    }
                    args[1..]
                        .iter()
                        .all(|arg| choose_index_is_guaranteed_scalar(arg, lexical_scopes))
                }
                Function::Let => {
                    if args.len() < 3 || args.len() % 2 == 0 {
                        return true;
                    }
                    let last = args.len() - 1;
                    lexical_scopes.push(HashMap::new());
                    for pair in args[..last].chunks_exact(2) {
                        let bytecode::Expr::NameRef(name) = &pair[0] else {
                            lexical_scopes.pop();
                            return true;
                        };
                        let kind = infer_binding_kind(&pair[1], lexical_scopes);
                        lexical_scopes
                            .last_mut()
                            .expect("pushed scope")
                            .insert(name.clone(), kind);
                    }
                    let ok = choose_index_is_guaranteed_scalar(&args[last], lexical_scopes);
                    lexical_scopes.pop();
                    ok
                }
                // ROW/COLUMN can return arrays when passed multi-cell references, so only treat them
                // as scalar indices when the argument is clearly a single-cell reference.
                Function::Row | Function::Column => match args.as_slice() {
                    [] => true,
                    [bytecode::Expr::CellRef(_)] => true,
                    [bytecode::Expr::RangeRef(r)] => r.start == r.end,
                    // Multi-range arguments produce #VALUE! (scalar error) in the bytecode runtime.
                    [bytecode::Expr::MultiRangeRef(_)] => true,
                    _ => true,
                },
                // These functions can return arrays when passed ranges/arrays (they map over the input),
                // so only allow them as scalar indices when the argument is scalar.
                Function::IsBlank
                | Function::IsNumber
                | Function::IsText
                | Function::IsLogical
                | Function::IsErr
                | Function::ErrorType
                | Function::N
                | Function::T => match args.as_slice() {
                    [arg] => choose_index_is_guaranteed_scalar(arg, lexical_scopes),
                    _ => true,
                },
                // OFFSET/INDIRECT are reference-valued and can spill; they are not scalar-safe as
                // CHOOSE/SWITCH indices.
                Function::Offset | Function::Indirect => false,
                // All other supported functions in the bytecode backend return scalars.
                _ => true,
            },
        }
    }

    match expr {
        bytecode::Expr::Literal(v) => match v {
            bytecode::Value::Number(_) | bytecode::Value::Bool(_) => true,
            bytecode::Value::Text(_) => true,
            bytecode::Value::Entity(_) | bytecode::Value::Record(_) => true,
            bytecode::Value::Empty => true,
            bytecode::Value::Missing => true,
            bytecode::Value::Error(_) => true,
            // Array literals are supported by the bytecode runtime as full typed arrays, but not
            // all bytecode function implementations support Excel's array-lifting semantics yet.
            // Gate array literals by context using the `allow_array_literals` flag.
            bytecode::Value::Array(_) => allow_array_literals,
            bytecode::Value::Range(_)
            | bytecode::Value::MultiRange(_)
            | bytecode::Value::Lambda(_) => false,
        },
        bytecode::Expr::CellRef(_) => true,
        bytecode::Expr::RangeRef(_) => allow_range,
        bytecode::Expr::SpillRange(inner) => {
            allow_range && bytecode_expr_is_eligible_inner(inner, true, false, lexical_scopes)
        }
        bytecode::Expr::NameRef(name) => match local_binding_kind(lexical_scopes, name) {
            Some(BytecodeLocalBindingKind::Scalar) => true,
            Some(BytecodeLocalBindingKind::RefSingle) => true,
            Some(BytecodeLocalBindingKind::Range) => allow_range,
            Some(BytecodeLocalBindingKind::ArrayLiteral) => allow_array_literals,
            None => false,
        },
        bytecode::Expr::MultiRangeRef(_) => allow_range,
        bytecode::Expr::Unary { op, expr } => match op {
            bytecode::ast::UnaryOp::Plus | bytecode::ast::UnaryOp::Neg => {
                bytecode_expr_is_eligible_inner(
                    expr,
                    allow_range,
                    allow_array_literals,
                    lexical_scopes,
                )
            }
            // Implicit intersection is only defined for references (ranges) in the bytecode VM.
            // Keep array literals ineligible so `@{...}` falls back to the AST evaluator.
            bytecode::ast::UnaryOp::ImplicitIntersection => {
                bytecode_expr_is_eligible_inner(expr, true, false, lexical_scopes)
            }
        },
        bytecode::Expr::Binary { op, left, right } => match op {
            bytecode::ast::BinaryOp::Union | bytecode::ast::BinaryOp::Intersect => {
                allow_range
                    && bytecode_expr_is_eligible_inner(left, true, false, lexical_scopes)
                    && bytecode_expr_is_eligible_inner(right, true, false, lexical_scopes)
            }
            _ => {
                matches!(
                    op,
                    bytecode::ast::BinaryOp::Add
                        | bytecode::ast::BinaryOp::Sub
                        | bytecode::ast::BinaryOp::Mul
                        | bytecode::ast::BinaryOp::Div
                        | bytecode::ast::BinaryOp::Pow
                        | bytecode::ast::BinaryOp::Eq
                        | bytecode::ast::BinaryOp::Ne
                        | bytecode::ast::BinaryOp::Lt
                        | bytecode::ast::BinaryOp::Le
                        | bytecode::ast::BinaryOp::Gt
                        | bytecode::ast::BinaryOp::Ge
                ) && bytecode_expr_is_eligible_inner(
                    left,
                    allow_range,
                    allow_array_literals,
                    lexical_scopes,
                ) && bytecode_expr_is_eligible_inner(
                    right,
                    allow_range,
                    allow_array_literals,
                    lexical_scopes,
                )
            }
        },
        bytecode::Expr::FuncCall { func, args } => match func {
            bytecode::ast::Function::FieldAccess => {
                if args.len() != 2 {
                    return false;
                }
                // Field access evaluates its base with dynamic dereference semantics (and may
                // therefore spill when the base is a multi-cell reference). Gate the base by the
                // caller's allow_range/allow_array_literals flags.
                //
                // The field name is a scalar argument, but callers can provide references via
                // direct `_FIELDACCESS` calls, so allow range values here and let the runtime apply
                // implicit intersection.
                bytecode_expr_is_eligible_inner(
                    &args[0],
                    allow_range,
                    allow_array_literals,
                    lexical_scopes,
                ) && bytecode_expr_is_eligible_inner(&args[1], true, false, lexical_scopes)
            }
            bytecode::ast::Function::If => {
                if args.len() < 2 || args.len() > 3 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::Choose => {
                if args.len() < 2 || args.len() > 255 {
                    return false;
                }

                // The bytecode backend only supports CHOOSE for scalar indices. If the index can
                // evaluate to an array, fall back to the AST evaluator.
                let index_ok =
                    bytecode_expr_is_eligible_inner(&args[0], false, false, lexical_scopes)
                        && choose_index_is_guaranteed_scalar(&args[0], lexical_scopes);
                if !index_ok {
                    return false;
                }

                args[1..].iter().all(|arg| {
                    bytecode_expr_is_eligible_inner(
                        arg,
                        allow_range,
                        allow_array_literals,
                        lexical_scopes,
                    )
                })
            }
            bytecode::ast::Function::Ifs => {
                if args.len() < 2 || args.len() % 2 != 0 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::And
            | bytecode::ast::Function::Or
            | bytecode::ast::Function::Xor => {
                // Excel limits AND/OR/XOR to 255 arguments.
                if args.is_empty() || args.len() > 255 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, true, true, lexical_scopes))
            }
            bytecode::ast::Function::IfError | bytecode::ast::Function::IfNa => {
                if args.len() != 2 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::IsError | bytecode::ast::Function::IsNa => {
                if args.len() != 1 {
                    return false;
                }
                bytecode_expr_is_eligible_inner(&args[0], true, true, lexical_scopes)
            }
            bytecode::ast::Function::True | bytecode::ast::Function::False => args.is_empty(),
            bytecode::ast::Function::Na => args.is_empty(),
            bytecode::ast::Function::Switch => {
                if args.len() < 3 {
                    return false;
                }
                let has_default = (args.len() - 1) % 2 != 0;
                let pairs_end = if has_default {
                    args.len() - 1
                } else {
                    args.len()
                };
                let pairs = &args[1..pairs_end];
                if pairs.len() < 2 || pairs.len() % 2 != 0 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::Let => {
                if args.len() < 3 || args.len() % 2 == 0 {
                    return false;
                }

                lexical_scopes.push(HashMap::new());

                for pair in args[..args.len() - 1].chunks_exact(2) {
                    let bytecode::Expr::NameRef(name) = &pair[0] else {
                        lexical_scopes.pop();
                        return false;
                    };
                    if name.is_empty() {
                        lexical_scopes.pop();
                        return false;
                    }

                    // Allow recursive lambdas of the form:
                    // `LET(f, LAMBDA(x, f(x)), f(1))`
                    // by treating the binding name as visible while checking the RHS.
                    lexical_scopes
                        .last_mut()
                        .expect("pushed scope")
                        .insert(name.clone(), BytecodeLocalBindingKind::Scalar);
                    // LET bindings can hold scalars, ranges, and array literals. We only need to
                    // ensure the overall LET expression remains non-spilling in scalar contexts;
                    // that is enforced by the `NameRef` eligibility check, which gates range/array
                    // locals based on the `allow_range` / `allow_array_literals` flags.
                    if !bytecode_expr_is_eligible_inner(&pair[1], true, true, lexical_scopes) {
                        lexical_scopes.pop();
                        return false;
                    }

                    let kind = infer_binding_kind(&pair[1], lexical_scopes);

                    lexical_scopes
                        .last_mut()
                        .expect("pushed scope")
                        .insert(name.clone(), kind);
                }

                // LET can return scalars, ranges, or array literals depending on context:
                // - Range references are gated by `allow_range`.
                // - Array literals are gated by `allow_array_literals`.
                let ok = bytecode_expr_is_eligible_inner(
                    &args[args.len() - 1],
                    allow_range,
                    allow_array_literals,
                    lexical_scopes,
                );
                lexical_scopes.pop();
                ok
            }
            bytecode::ast::Function::IsOmitted => {
                if args.len() != 1 {
                    return false;
                }
                // ISOMITTED is a special form: it requires a bare identifier argument, and should
                // not evaluate it as a value. It returns TRUE when the corresponding LAMBDA
                // parameter was omitted at the call site.
                matches!(args[0], bytecode::Expr::NameRef(_))
            }
            bytecode::ast::Function::Sum
            | bytecode::ast::Function::Average
            | bytecode::ast::Function::Min
            | bytecode::ast::Function::Max
            | bytecode::ast::Function::Count => args
                .iter()
                .all(|arg| bytecode_expr_is_eligible_inner(arg, true, true, lexical_scopes)),
            bytecode::ast::Function::CountA | bytecode::ast::Function::CountBlank => args
                .iter()
                .all(|arg| bytecode_expr_is_eligible_inner(arg, true, true, lexical_scopes)),
            bytecode::ast::Function::SumIf | bytecode::ast::Function::AverageIf => {
                if args.len() != 2 && args.len() != 3 {
                    return false;
                }
                let range_ok =
                    bytecode_expr_is_eligible_inner(&args[0], true, true, lexical_scopes);
                let criteria_ok =
                    bytecode_expr_is_eligible_inner(&args[1], false, false, lexical_scopes);
                let sum_range_ok = match args.get(2) {
                    None => true,
                    // Excel treats an explicitly missing optional range arg as "omitted".
                    Some(bytecode::Expr::Literal(bytecode::Value::Missing)) => true,
                    Some(arg) => bytecode_expr_is_eligible_inner(arg, true, true, lexical_scopes),
                };

                range_ok && criteria_ok && sum_range_ok
            }
            bytecode::ast::Function::SumIfs
            | bytecode::ast::Function::AverageIfs
            | bytecode::ast::Function::MinIfs
            | bytecode::ast::Function::MaxIfs => {
                if args.len() < 3 || (args.len() - 1) % 2 != 0 {
                    return false;
                }
                let value_range_ok =
                    bytecode_expr_is_eligible_inner(&args[0], true, true, lexical_scopes);
                if !value_range_ok {
                    return false;
                }
                for pair in args[1..].chunks_exact(2) {
                    let range_ok =
                        bytecode_expr_is_eligible_inner(&pair[0], true, true, lexical_scopes);
                    let criteria_ok =
                        bytecode_expr_is_eligible_inner(&pair[1], false, false, lexical_scopes);
                    if !range_ok || !criteria_ok {
                        return false;
                    }
                }
                true
            }
            bytecode::ast::Function::Row | bytecode::ast::Function::Column => match args.as_slice()
            {
                [] => true,
                // ROW/COLUMN are reference-only, but the bytecode runtime will still produce the
                // correct `#VALUE!` error for non-reference arguments. Allow any eligible argument
                // expression so higher-order constructs (e.g. `ROW(LAMBDA(r,r)(A1))`) can be
                // evaluated by the bytecode backend without forcing an AST fallback.
                [arg] => bytecode_expr_is_eligible_inner(arg, true, true, lexical_scopes),
                _ => false,
            },
            bytecode::ast::Function::Rows | bytecode::ast::Function::Columns => {
                if args.len() != 1 {
                    return false;
                }
                bytecode_expr_is_eligible_inner(&args[0], true, true, lexical_scopes)
            }
            bytecode::ast::Function::Address => {
                if !(2..=5).contains(&args.len()) {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::Offset => {
                // OFFSET returns a reference/range value. Only allow it in contexts that allow
                // reference values to propagate (spills or range-taking functions).
                if !allow_range {
                    return false;
                }
                if !(3..=5).contains(&args.len()) {
                    return false;
                }
                // base is reference-valued.
                let base_ok =
                    bytecode_expr_is_eligible_inner(&args[0], true, false, lexical_scopes);
                // rows/cols/height/width are scalar arguments, but Excel applies implicit
                // intersection when they are provided as references. Allow reference values but
                // reject array literals in these scalar positions.
                let mut scalar_ok = true;
                for arg in &args[1..] {
                    if !bytecode_expr_is_eligible_inner(arg, true, false, lexical_scopes) {
                        scalar_ok = false;
                        break;
                    }
                }
                base_ok && scalar_ok
            }
            bytecode::ast::Function::Indirect => {
                // INDIRECT returns a reference/range value, so only allow it in range contexts.
                if !allow_range {
                    return false;
                }
                if args.is_empty() || args.len() > 2 {
                    return false;
                }
                // Both arguments are scalar (text + optional bool), but accept references via
                // implicit intersection.
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, true, false, lexical_scopes))
            }
            bytecode::ast::Function::Rand => args.is_empty(),
            bytecode::ast::Function::RandBetween => {
                if args.len() != 2 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::CountIf => {
                if args.len() != 2 {
                    return false;
                }
                let range_ok =
                    bytecode_expr_is_eligible_inner(&args[0], true, true, lexical_scopes);
                let criteria_ok =
                    bytecode_expr_is_eligible_inner(&args[1], false, false, lexical_scopes);

                range_ok && criteria_ok
            }
            bytecode::ast::Function::CountIfs => {
                if args.len() < 2 || args.len() % 2 != 0 {
                    return false;
                }
                for pair in args.chunks_exact(2) {
                    let range_ok =
                        bytecode_expr_is_eligible_inner(&pair[0], true, true, lexical_scopes);
                    let criteria_ok =
                        bytecode_expr_is_eligible_inner(&pair[1], false, false, lexical_scopes);
                    if !range_ok || !criteria_ok {
                        return false;
                    }
                }
                true
            }
            bytecode::ast::Function::SumProduct => {
                if args.len() != 2 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, true, true, lexical_scopes))
            }
            bytecode::ast::Function::VLookup | bytecode::ast::Function::HLookup => {
                if args.len() < 3 || args.len() > 4 {
                    return false;
                }

                // table_array supports both range references and array values (e.g. array literals,
                // LET-bound arrays, and computed arrays like `A1:A3*10`). Require it to be eligible
                // in a range/array context; runtime still enforces the actual table semantics.
                let table_ok =
                    bytecode_expr_is_eligible_inner(&args[1], true, true, lexical_scopes);
                // `lookup_value` is a scalar argument, so Excel applies implicit intersection when
                // it is provided as a range reference. Allow range values here and let the runtime
                // perform implicit intersection (matching the AST evaluator's `eval_scalar_arg`).
                let lookup_ok =
                    bytecode_expr_is_eligible_inner(&args[0], true, false, lexical_scopes);
                let index_ok =
                    bytecode_expr_is_eligible_inner(&args[2], false, false, lexical_scopes);
                let range_lookup_ok = if args.len() == 4 {
                    bytecode_expr_is_eligible_inner(&args[3], false, false, lexical_scopes)
                } else {
                    true
                };

                table_ok && lookup_ok && index_ok && range_lookup_ok
            }
            bytecode::ast::Function::Match => {
                if args.len() < 2 || args.len() > 3 {
                    return false;
                }

                // MATCH accepts either reference-like lookup arrays or array values. Allow both
                // ranges (including spill ranges) and array literals/expressions here.
                let array_ok =
                    bytecode_expr_is_eligible_inner(&args[1], true, true, lexical_scopes);
                // `lookup_value` is scalar and uses implicit intersection when passed a range.
                let lookup_ok =
                    bytecode_expr_is_eligible_inner(&args[0], true, false, lexical_scopes);
                let match_type_ok = if args.len() == 3 {
                    bytecode_expr_is_eligible_inner(&args[2], false, false, lexical_scopes)
                } else {
                    true
                };

                array_ok && lookup_ok && match_type_ok
            }
            bytecode::ast::Function::CoupDayBs
            | bytecode::ast::Function::CoupDays
            | bytecode::ast::Function::CoupDaysNc
            | bytecode::ast::Function::CoupNcd
            | bytecode::ast::Function::CoupNum
            | bytecode::ast::Function::CoupPcd => {
                if args.len() != 3 && args.len() != 4 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::Price | bytecode::ast::Function::Yield => {
                if args.len() != 6 && args.len() != 7 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::Duration | bytecode::ast::Function::MDuration => {
                if args.len() != 5 && args.len() != 6 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::Accrintm => {
                if args.len() != 4 && args.len() != 5 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::Accrint => {
                if !(6..=8).contains(&args.len()) {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::Disc
            | bytecode::ast::Function::PriceDisc
            | bytecode::ast::Function::YieldDisc
            | bytecode::ast::Function::Intrate
            | bytecode::ast::Function::Received => {
                if args.len() != 4 && args.len() != 5 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::PriceMat | bytecode::ast::Function::YieldMat => {
                if args.len() != 5 && args.len() != 6 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::TbillEq
            | bytecode::ast::Function::TbillPrice
            | bytecode::ast::Function::TbillYield => {
                if args.len() != 3 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::OddFPrice | bytecode::ast::Function::OddFYield => {
                if args.len() != 8 && args.len() != 9 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::OddLPrice | bytecode::ast::Function::OddLYield => {
                if args.len() != 7 && args.len() != 8 {
                    return false;
                }
                args.iter()
                    .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes))
            }
            bytecode::ast::Function::Abs
            | bytecode::ast::Function::Int
            | bytecode::ast::Function::Round
            | bytecode::ast::Function::RoundUp
            | bytecode::ast::Function::RoundDown
            | bytecode::ast::Function::Mod
            | bytecode::ast::Function::Sign
            | bytecode::ast::Function::ConcatOp
            | bytecode::ast::Function::Not => args.iter().all(|arg| {
                if allow_array_literals {
                    bytecode_expr_is_eligible_inner(arg, true, true, lexical_scopes)
                } else {
                    bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes)
                }
            }),
            bytecode::ast::Function::Concat | bytecode::ast::Function::Concatenate => args
                .iter()
                .all(|arg| bytecode_expr_is_eligible_inner(arg, true, true, lexical_scopes)),
            bytecode::ast::Function::Now
            | bytecode::ast::Function::Today
            | bytecode::ast::Function::Db
            | bytecode::ast::Function::Vdb => args
                .iter()
                .all(|arg| bytecode_expr_is_eligible_inner(arg, false, false, lexical_scopes)),
            bytecode::ast::Function::IsBlank
            | bytecode::ast::Function::IsNumber
            | bytecode::ast::Function::IsText
            | bytecode::ast::Function::IsLogical
            | bytecode::ast::Function::IsErr
            | bytecode::ast::Function::ErrorType
            | bytecode::ast::Function::N
            | bytecode::ast::Function::T => {
                if args.len() != 1 {
                    return false;
                }
                bytecode_expr_is_eligible_inner(&args[0], true, true, lexical_scopes)
            }
            bytecode::ast::Function::Type => {
                if args.len() != 1 {
                    return false;
                }
                // TYPE is scalar even for multi-cell ranges/arrays (returns 64), so allow them.
                bytecode_expr_is_eligible_inner(&args[0], true, true, lexical_scopes)
            }
            bytecode::ast::Function::XMatch => {
                if args.len() < 2 || args.len() > 4 {
                    return false;
                }
                // `lookup_value` is scalar and uses implicit intersection when passed a range.
                let lookup_ok =
                    bytecode_expr_is_eligible_inner(&args[0], true, false, lexical_scopes);
                // Restrict lookup_array to "vector-like" arguments (ranges/spills/LET refs or 1D
                // array literals). This keeps scalar-only implicit intersection out of the bytecode
                // path while still supporting common XLOOKUP/XMATCH patterns.
                //
                // In addition to direct references (A1:A3), allow spill ranges (A1#) and LET-bound
                // range locals (LET(a, A1:A3, XMATCH(..., a))).
                let lookup_array_is_range_like = matches!(args[1], bytecode::Expr::CellRef(_))
                    || matches!(
                        infer_binding_kind(&args[1], lexical_scopes),
                        BytecodeLocalBindingKind::Range
                            | BytecodeLocalBindingKind::RefSingle
                            | BytecodeLocalBindingKind::ArrayLiteral
                    );
                let lookup_array_ok = lookup_array_is_range_like
                    && bytecode_expr_is_eligible_inner(&args[1], true, true, lexical_scopes);
                let match_mode_ok = args.get(2).map_or(true, |arg| {
                    // match_mode is a scalar argument, but Excel applies implicit intersection when
                    // it is provided as a range reference.
                    bytecode_expr_is_eligible_inner(arg, true, false, lexical_scopes)
                });
                let search_mode_ok = args.get(3).map_or(true, |arg| {
                    // search_mode is a scalar argument, but Excel applies implicit intersection when
                    // it is provided as a range reference.
                    bytecode_expr_is_eligible_inner(arg, true, false, lexical_scopes)
                });
                lookup_ok && lookup_array_ok && match_mode_ok && search_mode_ok
            }
            bytecode::ast::Function::XLookup => {
                if args.len() < 3 || args.len() > 6 {
                    return false;
                }
                // `lookup_value` is scalar and uses implicit intersection when passed a range.
                let lookup_ok =
                    bytecode_expr_is_eligible_inner(&args[0], true, false, lexical_scopes);
                let lookup_array_is_range_like = matches!(args[1], bytecode::Expr::CellRef(_))
                    || matches!(
                        infer_binding_kind(&args[1], lexical_scopes),
                        BytecodeLocalBindingKind::Range
                            | BytecodeLocalBindingKind::RefSingle
                            | BytecodeLocalBindingKind::ArrayLiteral
                    );
                let lookup_array_ok = lookup_array_is_range_like
                    && bytecode_expr_is_eligible_inner(&args[1], true, true, lexical_scopes);
                // Bytecode supports XLOOKUP's vector spill semantics (row/column slice) as long as
                // lookup_array/return_array are vector-like arguments (ranges or arrays).
                let return_array_is_range_like = matches!(args[2], bytecode::Expr::CellRef(_))
                    || matches!(
                        infer_binding_kind(&args[2], lexical_scopes),
                        BytecodeLocalBindingKind::Range
                            | BytecodeLocalBindingKind::RefSingle
                            | BytecodeLocalBindingKind::ArrayLiteral
                    );
                let return_array_ok = return_array_is_range_like
                    && bytecode_expr_is_eligible_inner(&args[2], true, true, lexical_scopes);
                let if_not_found_ok = args.get(3).map_or(true, |arg| {
                    // `if_not_found` is a scalar argument. Like other scalar arguments, Excel
                    // applies implicit intersection when it is passed as a reference/range, and it
                    // can also be an array literal/expression (allowing the fallback to spill).
                    //
                    // Examples:
                    // - `XLOOKUP(99,{1;2},{10;20},B1:B2)` -> implicitly intersects `B1:B2`.
                    // - `XLOOKUP(99,{1;2},{10;20},{100;200})` -> spills `{100;200}`.
                    bytecode_expr_is_eligible_inner(arg, true, true, lexical_scopes)
                });
                let match_mode_ok = args.get(4).map_or(true, |arg| {
                    // match_mode is a scalar argument, but Excel applies implicit intersection when
                    // it is provided as a range reference.
                    bytecode_expr_is_eligible_inner(arg, true, false, lexical_scopes)
                });
                let search_mode_ok = args.get(5).map_or(true, |arg| {
                    // search_mode is a scalar argument, but Excel applies implicit intersection when
                    // it is provided as a range reference.
                    bytecode_expr_is_eligible_inner(arg, true, false, lexical_scopes)
                });
                lookup_ok
                    && lookup_array_ok
                    && return_array_ok
                    && if_not_found_ok
                    && match_mode_ok
                    && search_mode_ok
            }
            bytecode::ast::Function::Unknown(_) => false,
        },
        bytecode::Expr::Lambda { params, body } => {
            lexical_scopes.push(HashMap::new());
            for p in params.iter() {
                if p.is_empty() {
                    lexical_scopes.pop();
                    return false;
                }
                lexical_scopes
                    .last_mut()
                    .expect("pushed scope")
                    .insert(p.clone(), BytecodeLocalBindingKind::Scalar);
            }
            let ok = bytecode_expr_is_eligible_inner(body, false, false, lexical_scopes);
            lexical_scopes.pop();
            ok
        }
        bytecode::Expr::Call { callee, args } => {
            if !bytecode_expr_is_eligible_inner(callee, false, false, lexical_scopes) {
                return false;
            }
            args.iter()
                .all(|arg| bytecode_expr_is_eligible_inner(arg, true, true, lexical_scopes))
        }
    }
}

fn analyze_expr_flags(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
    workbook: &Workbook,
    external_refs_volatile: bool,
) -> (HashSet<String>, bool, bool, bool, bool) {
    let mut names = HashSet::new();
    let mut volatile = false;
    let mut thread_safe = true;
    let mut dynamic_deps = false;
    let mut origin_deps = false;
    let mut visiting_names = HashSet::new();
    let mut lexical_scopes: Vec<HashSet<String>> = Vec::new();
    walk_expr_flags(
        expr,
        current_cell,
        tables_by_sheet,
        workbook,
        &mut names,
        &mut volatile,
        &mut thread_safe,
        &mut dynamic_deps,
        &mut origin_deps,
        &mut visiting_names,
        &mut lexical_scopes,
        external_refs_volatile,
    );
    (names, volatile, thread_safe, dynamic_deps, origin_deps)
}

fn walk_expr_flags(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
    workbook: &Workbook,
    names: &mut HashSet<String>,
    volatile: &mut bool,
    thread_safe: &mut bool,
    dynamic_deps: &mut bool,
    origin_deps: &mut bool,
    visiting_names: &mut HashSet<(SheetId, String)>,
    lexical_scopes: &mut Vec<HashSet<String>>,
    external_refs_volatile: bool,
) {
    fn name_is_local(scopes: &[HashSet<String>], name_key: &str) -> bool {
        scopes.iter().rev().any(|scope| scope.contains(name_key))
    }

    fn bare_identifier(expr: &CompiledExpr) -> Option<String> {
        match expr {
            Expr::NameRef(nref) if matches!(nref.sheet, SheetReference::Current) => {
                Some(normalize_defined_name(&nref.name))
            }
            _ => None,
        }
    }

    match expr {
        Expr::NameRef(nref) => {
            let Some(sheet) = resolve_single_sheet(&nref.sheet, current_cell.sheet) else {
                return;
            };
            let name_key = normalize_defined_name(&nref.name);
            if name_key.is_empty() {
                return;
            }

            // LET/LAMBDA lexical bindings are only visible for unqualified identifiers.
            // If a name reference is explicitly sheet-qualified (e.g. `Sheet1!X`), it should
            // bypass the local LET/LAMBDA scope and resolve as a defined name.
            if matches!(nref.sheet, SheetReference::Current)
                && name_is_local(lexical_scopes, &name_key)
            {
                return;
            }

            // Bare table names (e.g. `=Table1`) are treated as table references when no defined
            // name exists. Do not register them as defined-name dependencies.
            if resolve_defined_name(workbook, sheet, &name_key).is_none() {
                let candidate = nref.name.trim();
                if !candidate.is_empty()
                    && tables_by_sheet.iter().flatten().any(|t| {
                        crate::value::cmp_case_insensitive(&t.name, candidate) == Ordering::Equal
                            || crate::value::cmp_case_insensitive(&t.display_name, candidate)
                                == Ordering::Equal
                    })
                {
                    return;
                }
            }

            names.insert(name_key.clone());

            let visit_key = (sheet, name_key.clone());
            if !visiting_names.insert(visit_key.clone()) {
                // Cycle in the name definition graph. Stop expanding to avoid infinite recursion;
                // evaluation will surface `#NAME?` via the runtime recursion guard.
                return;
            }

            if let Some(def) = resolve_defined_name(workbook, sheet, &name_key) {
                if let Some(expr) = def.compiled.as_ref() {
                    walk_expr_flags(
                        expr,
                        CellKey {
                            sheet,
                            addr: current_cell.addr,
                        },
                        tables_by_sheet,
                        workbook,
                        names,
                        volatile,
                        thread_safe,
                        dynamic_deps,
                        origin_deps,
                        visiting_names,
                        lexical_scopes,
                        external_refs_volatile,
                    );
                }
            }

            visiting_names.remove(&visit_key);
        }
        Expr::Unary { expr, .. } | Expr::Postfix { expr, .. } => {
            walk_expr_flags(
                expr,
                current_cell,
                tables_by_sheet,
                workbook,
                names,
                volatile,
                thread_safe,
                dynamic_deps,
                origin_deps,
                visiting_names,
                lexical_scopes,
                external_refs_volatile,
            );
        }
        Expr::FieldAccess { base, .. } => {
            walk_expr_flags(
                base,
                current_cell,
                tables_by_sheet,
                workbook,
                names,
                volatile,
                thread_safe,
                dynamic_deps,
                origin_deps,
                visiting_names,
                lexical_scopes,
                external_refs_volatile,
            );
        }
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            walk_expr_flags(
                left,
                current_cell,
                tables_by_sheet,
                workbook,
                names,
                volatile,
                thread_safe,
                dynamic_deps,
                origin_deps,
                visiting_names,
                lexical_scopes,
                external_refs_volatile,
            );
            walk_expr_flags(
                right,
                current_cell,
                tables_by_sheet,
                workbook,
                names,
                volatile,
                thread_safe,
                dynamic_deps,
                origin_deps,
                visiting_names,
                lexical_scopes,
                external_refs_volatile,
            );
        }
        Expr::FunctionCall { name, args, .. } => {
            if let Some(spec) = crate::functions::lookup_function(name) {
                if spec.volatility == crate::functions::Volatility::Volatile {
                    *volatile = true;
                }
                if spec.thread_safety == crate::functions::ThreadSafety::NotThreadSafe {
                    *thread_safe = false;
                }
                // Some functions produce runtime-determined precedents that must be traced during
                // evaluation so the dependency graph stays accurate across recalcs.
                //
                // - OFFSET / INDIRECT: reference-returning functions.
                // - GETPIVOTDATA: depends on the (runtime-registered) pivot output range; the
                //   function records a dynamic reference to the full pivot destination so pivot
                //   refreshes trigger dependent formulas even when the `pivot_table` argument is a
                //   single (unchanged) cell.
                if matches!(spec.name, "OFFSET" | "INDIRECT" | "GETPIVOTDATA") {
                    *dynamic_deps = true;
                }

                if spec.name == "INFO" {
                    // `INFO("origin")` depends on host-provided worksheet view state (scroll position +
                    // frozen panes). Detect formulas that could depend on that value so
                    // `set_sheet_origin` can mark them dirty.
                    //
                    // If the key is a constant string, we can be precise; otherwise conservatively
                    // assume it could evaluate to `"origin"` at runtime.
                    match args.first() {
                        Some(Expr::Text(s)) => {
                            if s.trim().eq_ignore_ascii_case("origin") {
                                *origin_deps = true;
                            }
                        }
                        Some(_) | None => {
                            *origin_deps = true;
                        }
                    }
                }

                match spec.name {
                    "CELL" => {
                        // Excel worksheet metadata can affect certain CELL info_types even when no
                        // referenced cell values change (notably `CELL("width")` which depends on
                        // column width/hidden state). Excel treats these formulas as volatile; the
                        // engine approximates that behavior by marking `CELL("width", ...)` calls
                        // volatile at compile time.
                        //
                        // Keep the volatility narrow (info_type-dependent) so other CELL keys that
                        // are purely value/address based do not force unrelated recalculation.
                        if let Some(info_type) = args.first() {
                            match info_type {
                                Expr::Text(s) => {
                                    if s.trim().eq_ignore_ascii_case("width") {
                                        *volatile = true;
                                    }
                                }
                                // If the info_type is non-constant, conservatively treat the
                                // formula as volatile since it could evaluate to "width" at
                                // runtime.
                                _ => {
                                    *volatile = true;
                                }
                            }
                        }
                    }
                    "LET" => {
                        if args.len() < 3 || args.len() % 2 == 0 {
                            return;
                        }

                        lexical_scopes.push(HashSet::new());
                        for pair in args[..args.len() - 1].chunks_exact(2) {
                            let Some(name_key) = bare_identifier(&pair[0]) else {
                                lexical_scopes.pop();
                                return;
                            };

                            // Allow recursive lambdas of the form:
                            //   LET(f, LAMBDA(x, f(x)), f(1))
                            //
                            // The LET binding name isn't in scope while evaluating the value expression, but
                            // Excel's lambda invocation semantics inject the call name into the call scope
                            // at runtime, enabling recursion. Treat the binding name as local while walking
                            // the lambda body so we don't incorrectly mark it as an unresolved UDF / defined
                            // name reference (which would disable the bytecode backend).
                            if matches!(&pair[1], Expr::FunctionCall { name, .. } if name == "LAMBDA")
                            {
                                lexical_scopes
                                    .last_mut()
                                    .expect("pushed scope")
                                    .insert(name_key.clone());
                            }

                            walk_expr_flags(
                                &pair[1],
                                current_cell,
                                tables_by_sheet,
                                workbook,
                                names,
                                volatile,
                                thread_safe,
                                dynamic_deps,
                                origin_deps,
                                visiting_names,
                                lexical_scopes,
                                external_refs_volatile,
                            );
                            lexical_scopes
                                .last_mut()
                                .expect("pushed scope")
                                .insert(name_key);
                        }

                        walk_expr_flags(
                            &args[args.len() - 1],
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            names,
                            volatile,
                            thread_safe,
                            dynamic_deps,
                            origin_deps,
                            visiting_names,
                            lexical_scopes,
                            external_refs_volatile,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    "LAMBDA" => {
                        if args.is_empty() {
                            return;
                        }

                        let mut scope = HashSet::new();
                        for param in &args[..args.len() - 1] {
                            let Some(name_key) = bare_identifier(param) else {
                                return;
                            };
                            if !scope.insert(name_key) {
                                return;
                            }
                        }

                        lexical_scopes.push(scope);
                        walk_expr_flags(
                            &args[args.len() - 1],
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            names,
                            volatile,
                            thread_safe,
                            dynamic_deps,
                            origin_deps,
                            visiting_names,
                            lexical_scopes,
                            external_refs_volatile,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    _ => {}
                }
            } else {
                let name_key = normalize_defined_name(name);
                let is_local = !name_key.is_empty() && name_is_local(lexical_scopes, &name_key);
                let mut resolved_defined_name = None;
                if !name_key.is_empty() && !is_local {
                    names.insert(name_key.clone());

                    let sheet = current_cell.sheet;
                    let visit_key = (sheet, name_key.clone());
                    if visiting_names.insert(visit_key.clone()) {
                        resolved_defined_name = resolve_defined_name(workbook, sheet, &name_key);
                        if let Some(def) = resolved_defined_name.as_ref() {
                            if let Some(expr) = def.compiled.as_ref() {
                                walk_expr_flags(
                                    expr,
                                    CellKey {
                                        sheet,
                                        addr: current_cell.addr,
                                    },
                                    tables_by_sheet,
                                    workbook,
                                    names,
                                    volatile,
                                    thread_safe,
                                    dynamic_deps,
                                    origin_deps,
                                    visiting_names,
                                    lexical_scopes,
                                    external_refs_volatile,
                                );
                            }
                        }
                        visiting_names.remove(&visit_key);
                    }
                }

                // Placeholder: treat unresolved UDFs as non-thread-safe.
                if !is_local && resolved_defined_name.is_none() {
                    *thread_safe = false;
                }
            }
            for a in args {
                walk_expr_flags(
                    a,
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    names,
                    volatile,
                    thread_safe,
                    dynamic_deps,
                    origin_deps,
                    visiting_names,
                    lexical_scopes,
                    external_refs_volatile,
                );
            }
        }
        Expr::Call { callee, args } => {
            walk_expr_flags(
                callee,
                current_cell,
                tables_by_sheet,
                workbook,
                names,
                volatile,
                thread_safe,
                dynamic_deps,
                origin_deps,
                visiting_names,
                lexical_scopes,
                external_refs_volatile,
            );
            for a in args {
                walk_expr_flags(
                    a,
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    names,
                    volatile,
                    thread_safe,
                    dynamic_deps,
                    origin_deps,
                    visiting_names,
                    lexical_scopes,
                    external_refs_volatile,
                );
            }
        }
        Expr::ArrayLiteral { values, .. } => {
            for el in values.iter() {
                walk_expr_flags(
                    el,
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    names,
                    volatile,
                    thread_safe,
                    dynamic_deps,
                    origin_deps,
                    visiting_names,
                    lexical_scopes,
                    external_refs_volatile,
                );
            }
        }
        Expr::ImplicitIntersection(inner) | Expr::SpillRange(inner) => {
            walk_expr_flags(
                inner,
                current_cell,
                tables_by_sheet,
                workbook,
                names,
                volatile,
                thread_safe,
                dynamic_deps,
                origin_deps,
                visiting_names,
                lexical_scopes,
                external_refs_volatile,
            );
        }
        Expr::CellRef(r) => {
            if external_refs_volatile {
                if let SheetReference::External(key) = &r.sheet {
                    if crate::eval::split_external_sheet_key(key).is_some() {
                        *volatile = true;
                    }
                }
            }
        }
        Expr::RangeRef(r) => {
            if external_refs_volatile {
                if let SheetReference::External(key) = &r.sheet {
                    if crate::eval::split_external_sheet_key(key).is_some() {
                        *volatile = true;
                    }
                }
            }
        }
        Expr::StructuredRef(r) => {
            if external_refs_volatile {
                if let SheetReference::External(key) = &r.sheet {
                    // Workbook-only external structured refs compile as `SheetReference::External("[Book.xlsx]")`,
                    // so treat any bracket-prefixed key as an external workbook dependency.
                    if key.starts_with('[') {
                        *volatile = true;
                    }
                }
            }
        }
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Blank | Expr::Error(_) => {}
    }
}

fn analyze_calc_precedents(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
    workbook: &Workbook,
    spills: &SpillState,
) -> HashSet<Precedent> {
    let mut out = HashSet::new();
    let mut visiting_names = HashSet::new();
    let mut lexical_scopes: Vec<HashSet<String>> = Vec::new();
    walk_calc_expr(
        expr,
        current_cell,
        tables_by_sheet,
        workbook,
        spills,
        &mut out,
        &mut visiting_names,
        &mut lexical_scopes,
    );
    out
}

fn analyze_external_precedents(
    expr: &CompiledExpr,
    current_cell: CellKey,
    workbook: &Workbook,
    external_value_provider: Option<&dyn ExternalValueProvider>,
) -> Vec<PrecedentNode> {
    let mut out: HashSet<PrecedentNode> = HashSet::new();
    let mut visiting_names = HashSet::new();
    let mut lexical_scopes: Vec<HashSet<String>> = Vec::new();
    walk_external_expr(
        expr,
        current_cell,
        workbook,
        external_value_provider,
        &mut out,
        &mut visiting_names,
        &mut lexical_scopes,
    );
    out.into_iter().collect()
}

/// Expands an external workbook 3D sheet span key like `[Book.xlsx]Sheet1:Sheet3`
/// into per-sheet external keys like `[Book.xlsx]Sheet2`.
///
/// Returns `None` when expansion is not possible (e.g. missing provider, missing sheet order,
/// or unknown boundary sheets).
fn expand_external_sheet_span_key(
    key: &str,
    provider: Option<&dyn ExternalValueProvider>,
) -> Option<Vec<String>> {
    let provider = provider?;
    let (workbook, start, end) = crate::eval::split_external_sheet_span_key(key)?;
    let sheet_names = provider.workbook_sheet_names(workbook)?;

    let mut start_idx: Option<usize> = None;
    let mut end_idx: Option<usize> = None;
    for (idx, name) in sheet_names.iter().enumerate() {
        if start_idx.is_none() && formula_model::sheet_name_eq_case_insensitive(name, start) {
            start_idx = Some(idx);
        }
        if end_idx.is_none() && formula_model::sheet_name_eq_case_insensitive(name, end) {
            end_idx = Some(idx);
        }
        if start_idx.is_some() && end_idx.is_some() {
            break;
        }
    }

    let start_idx = start_idx?;
    let end_idx = end_idx?;
    let (lo, hi) = if start_idx <= end_idx {
        (start_idx, end_idx)
    } else {
        (end_idx, start_idx)
    };

    Some(
        sheet_names[lo..=hi]
            .iter()
            .map(|name| format!("[{workbook}]{name}"))
            .collect(),
    )
}

fn analyze_external_dependencies(
    expr: &CompiledExpr,
    current_cell: CellKey,
    workbook: &Workbook,
    external_value_provider: Option<&dyn ExternalValueProvider>,
) -> (HashSet<String>, HashSet<String>) {
    let mut external_sheets: HashSet<String> = HashSet::new();
    let mut external_workbooks: HashSet<String> = HashSet::new();
    let mut visiting_names = HashSet::new();
    let mut lexical_scopes: Vec<HashSet<String>> = Vec::new();
    walk_external_dependencies(
        expr,
        current_cell,
        workbook,
        external_value_provider,
        &mut external_sheets,
        &mut external_workbooks,
        &mut visiting_names,
        &mut lexical_scopes,
    );
    (external_sheets, external_workbooks)
}

fn walk_external_dependencies(
    expr: &CompiledExpr,
    current_cell: CellKey,
    workbook: &Workbook,
    external_value_provider: Option<&dyn ExternalValueProvider>,
    external_sheets: &mut HashSet<String>,
    external_workbooks: &mut HashSet<String>,
    visiting_names: &mut HashSet<(SheetId, String)>,
    lexical_scopes: &mut Vec<HashSet<String>>,
) {
    fn name_is_local(scopes: &[HashSet<String>], name_key: &str) -> bool {
        scopes.iter().rev().any(|scope| scope.contains(name_key))
    }

    fn bare_identifier(expr: &CompiledExpr) -> Option<String> {
        match expr {
            Expr::NameRef(nref) if matches!(nref.sheet, SheetReference::Current) => {
                Some(normalize_defined_name(&nref.name))
            }
            _ => None,
        }
    }

    match expr {
        Expr::CellRef(r) => {
            if let SheetReference::External(key) = &r.sheet {
                if let Some((workbook, _sheet)) = crate::eval::split_external_sheet_key(key) {
                    external_workbooks.insert(workbook.to_string());
                }
                if crate::eval::is_valid_external_sheet_key(key) {
                    external_sheets.insert(key.clone());
                } else if crate::eval::split_external_sheet_span_key(key).is_some() {
                    if let Some(expanded) =
                        expand_external_sheet_span_key(key, external_value_provider)
                    {
                        for sheet_key in expanded {
                            external_sheets.insert(sheet_key);
                        }
                    } else {
                        external_sheets.insert(key.clone());
                    }
                }
            }
        }
        Expr::RangeRef(r) => {
            if let SheetReference::External(key) = &r.sheet {
                if let Some((workbook, _sheet)) = crate::eval::split_external_sheet_key(key) {
                    external_workbooks.insert(workbook.to_string());
                }
                if crate::eval::is_valid_external_sheet_key(key) {
                    external_sheets.insert(key.clone());
                } else if crate::eval::split_external_sheet_span_key(key).is_some() {
                    if let Some(expanded) =
                        expand_external_sheet_span_key(key, external_value_provider)
                    {
                        for sheet_key in expanded {
                            external_sheets.insert(sheet_key);
                        }
                    } else {
                        external_sheets.insert(key.clone());
                    }
                }
            }
        }
        Expr::StructuredRef(r) => {
            if let SheetReference::External(key) = &r.sheet {
                // Structured refs can be workbook-only (e.g. `[Book.xlsx]Table1[Col]`), so handle
                // both `[workbook]sheet` and `[workbook]` forms.
                if let Some((workbook, _sheet)) = crate::eval::split_external_sheet_key(key) {
                    external_workbooks.insert(workbook.to_string());
                    if crate::eval::is_valid_external_sheet_key(key) {
                        external_sheets.insert(key.clone());
                    }
                } else if key.starts_with('[') {
                    // Workbook-only external ref key like `[Book.xlsx]`.
                    let Some(end) = key.rfind(']') else {
                        return;
                    };
                    if end <= 1 {
                        return;
                    }
                    let workbook = &key[1..end];
                    if workbook.is_empty() {
                        return;
                    }
                    external_workbooks.insert(workbook.to_string());

                    // Attempt to refine workbook-level invalidation down to a sheet key when table
                    // metadata is available (e.g. `[Book.xlsx]Table1[Col]`).
                    if let Some(provider) = external_value_provider {
                        if let Some(table_name) = r.sref.table_name.as_deref() {
                            if let Some((table_sheet, _table)) =
                                provider.workbook_table(workbook, table_name)
                            {
                                external_sheets.insert(format!("[{workbook}]{table_sheet}"));
                            }
                        }
                    }
                }
            }
        }
        Expr::NameRef(nref) => {
            let Some(sheet) = resolve_sheet(&nref.sheet, current_cell.sheet) else {
                return;
            };
            let name_key = normalize_defined_name(&nref.name);
            if name_key.is_empty() {
                return;
            }

            // LET/LAMBDA lexical bindings are only visible for unqualified identifiers.
            // Explicit sheet-qualified names should still resolve as defined names.
            if matches!(nref.sheet, SheetReference::Current)
                && name_is_local(lexical_scopes, &name_key)
            {
                return;
            }

            let visit_key = (sheet, name_key.clone());
            if !visiting_names.insert(visit_key.clone()) {
                return;
            }
            if let Some(def) = resolve_defined_name(workbook, sheet, &name_key) {
                if let Some(expr) = def.compiled.as_ref() {
                    walk_external_dependencies(
                        expr,
                        CellKey {
                            sheet,
                            addr: current_cell.addr,
                        },
                        workbook,
                        external_value_provider,
                        external_sheets,
                        external_workbooks,
                        visiting_names,
                        lexical_scopes,
                    );
                }
            }
            visiting_names.remove(&visit_key);
        }
        Expr::FieldAccess { base, .. } => walk_external_dependencies(
            base,
            current_cell,
            workbook,
            external_value_provider,
            external_sheets,
            external_workbooks,
            visiting_names,
            lexical_scopes,
        ),
        Expr::Unary { expr, .. }
        | Expr::Postfix { expr, .. }
        | Expr::ImplicitIntersection(expr)
        | Expr::SpillRange(expr) => walk_external_dependencies(
            expr,
            current_cell,
            workbook,
            external_value_provider,
            external_sheets,
            external_workbooks,
            visiting_names,
            lexical_scopes,
        ),
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            walk_external_dependencies(
                left,
                current_cell,
                workbook,
                external_value_provider,
                external_sheets,
                external_workbooks,
                visiting_names,
                lexical_scopes,
            );
            walk_external_dependencies(
                right,
                current_cell,
                workbook,
                external_value_provider,
                external_sheets,
                external_workbooks,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::FunctionCall { name, args, .. } => {
            if let Some(spec) = crate::functions::lookup_function(name) {
                match spec.name {
                    "INDIRECT" => {
                        // `INDIRECT` can dynamically produce references (including external workbook
                        // references) from text. Most uses are not statically analyzable, but if the
                        // ref_text and A1 flag are constant we can extract external dependencies up
                        // front so external invalidation works even before the formula has been
                        // evaluated.
                        if let Some(Expr::Text(text)) = args.first() {
                            // Only attempt static extraction when the optional A1 flag is either
                            // omitted or a literal boolean; otherwise the reference style is runtime
                            // dependent.
                            let a1 = match args.get(1) {
                                None => Some(true),
                                Some(Expr::Bool(v)) => Some(*v),
                                Some(_) => None,
                            };
                            if let Some(a1) = a1 {
                                let ref_text = text.trim();
                                if !ref_text.is_empty() {
                                    // Mirror `functions::builtins_reference::indirect_fn` parsing behavior:
                                    // parse the text as a standalone reference expression and only accept
                                    // simple cell/range references.
                                    if let Ok(parsed) = crate::parse_formula(
                                        ref_text,
                                        crate::ParseOptions {
                                            locale: crate::LocaleConfig::en_us(),
                                            reference_style: if a1 {
                                                crate::ReferenceStyle::A1
                                            } else {
                                                crate::ReferenceStyle::R1C1
                                            },
                                            normalize_relative_to: None,
                                        },
                                    ) {
                                        let origin_ast = crate::CellAddr::new(
                                            current_cell.addr.row,
                                            current_cell.addr.col,
                                        );
                                        let lowered = crate::eval::lower_ast(
                                            &parsed,
                                            if a1 { None } else { Some(origin_ast) },
                                        );
                                        let sheet_ref = match lowered {
                                            crate::eval::Expr::CellRef(r) => Some(r.sheet),
                                            crate::eval::Expr::RangeRef(r) => Some(r.sheet),
                                            _ => None,
                                        };
                                        if let Some(crate::eval::SheetReference::External(key)) =
                                            sheet_ref
                                        {
                                            // Match the runtime behavior: allow single-sheet external workbook
                                            // references, but reject external 3D spans.
                                            if crate::eval::is_valid_external_sheet_key(&key) {
                                                if let Some((workbook_id, _sheet)) =
                                                    crate::eval::split_external_sheet_key(&key)
                                                {
                                                    external_workbooks
                                                        .insert(workbook_id.to_string());
                                                }
                                                external_sheets.insert(key);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    "LET" => {
                        if args.len() < 3 || args.len() % 2 == 0 {
                            return;
                        }

                        lexical_scopes.push(HashSet::new());
                        for pair in args[..args.len() - 1].chunks_exact(2) {
                            let Some(name_key) = bare_identifier(&pair[0]) else {
                                lexical_scopes.pop();
                                return;
                            };
                            walk_external_dependencies(
                                &pair[1],
                                current_cell,
                                workbook,
                                external_value_provider,
                                external_sheets,
                                external_workbooks,
                                visiting_names,
                                lexical_scopes,
                            );
                            lexical_scopes
                                .last_mut()
                                .expect("pushed scope")
                                .insert(name_key);
                        }

                        walk_external_dependencies(
                            &args[args.len() - 1],
                            current_cell,
                            workbook,
                            external_value_provider,
                            external_sheets,
                            external_workbooks,
                            visiting_names,
                            lexical_scopes,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    "LAMBDA" => {
                        if args.is_empty() {
                            return;
                        }

                        let mut scope = HashSet::new();
                        for param in &args[..args.len() - 1] {
                            let Some(name_key) = bare_identifier(param) else {
                                return;
                            };
                            if !scope.insert(name_key) {
                                return;
                            }
                        }

                        lexical_scopes.push(scope);
                        walk_external_dependencies(
                            &args[args.len() - 1],
                            current_cell,
                            workbook,
                            external_value_provider,
                            external_sheets,
                            external_workbooks,
                            visiting_names,
                            lexical_scopes,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    _ => {}
                }
            } else {
                // Unknown function name: treat it as a potential defined name (UDF) and expand the
                // definition if present.
                let name_key = normalize_defined_name(name);
                let is_local = !name_key.is_empty() && name_is_local(lexical_scopes, &name_key);
                if !name_key.is_empty() && !is_local {
                    let sheet = current_cell.sheet;
                    let visit_key = (sheet, name_key.clone());
                    if visiting_names.insert(visit_key.clone()) {
                        if let Some(def) = resolve_defined_name(workbook, sheet, &name_key) {
                            if let Some(expr) = def.compiled.as_ref() {
                                walk_external_dependencies(
                                    expr,
                                    CellKey {
                                        sheet,
                                        addr: current_cell.addr,
                                    },
                                    workbook,
                                    external_value_provider,
                                    external_sheets,
                                    external_workbooks,
                                    visiting_names,
                                    lexical_scopes,
                                );
                            }
                        }
                        visiting_names.remove(&visit_key);
                    }
                }
            }

            for a in args {
                walk_external_dependencies(
                    a,
                    current_cell,
                    workbook,
                    external_value_provider,
                    external_sheets,
                    external_workbooks,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::Call { callee, args } => {
            walk_external_dependencies(
                callee,
                current_cell,
                workbook,
                external_value_provider,
                external_sheets,
                external_workbooks,
                visiting_names,
                lexical_scopes,
            );
            for a in args {
                walk_external_dependencies(
                    a,
                    current_cell,
                    workbook,
                    external_value_provider,
                    external_sheets,
                    external_workbooks,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::ArrayLiteral { values, .. } => {
            for el in values.iter() {
                walk_external_dependencies(
                    el,
                    current_cell,
                    workbook,
                    external_value_provider,
                    external_sheets,
                    external_workbooks,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Blank | Expr::Error(_) => {}
    }
}

fn walk_external_expr(
    expr: &CompiledExpr,
    current_cell: CellKey,
    workbook: &Workbook,
    external_value_provider: Option<&dyn ExternalValueProvider>,
    precedents: &mut HashSet<PrecedentNode>,
    visiting_names: &mut HashSet<(SheetId, String)>,
    lexical_scopes: &mut Vec<HashSet<String>>,
) {
    fn name_is_local(scopes: &[HashSet<String>], name_key: &str) -> bool {
        scopes.iter().rev().any(|scope| scope.contains(name_key))
    }

    fn bare_identifier(expr: &CompiledExpr) -> Option<String> {
        match expr {
            Expr::NameRef(nref) if matches!(nref.sheet, SheetReference::Current) => {
                Some(normalize_defined_name(&nref.name))
            }
            _ => None,
        }
    }

    match expr {
        Expr::CellRef(r) => {
            if let SheetReference::External(key) = &r.sheet {
                if crate::eval::is_valid_external_sheet_key(key) {
                    let Some(addr) = r.addr.resolve(current_cell.addr) else {
                        return;
                    };
                    precedents.insert(PrecedentNode::ExternalCell {
                        sheet: key.clone(),
                        addr,
                    });
                } else if crate::eval::split_external_sheet_span_key(key).is_some() {
                    let Some(addr) = r.addr.resolve(current_cell.addr) else {
                        return;
                    };
                    if let Some(expanded) =
                        expand_external_sheet_span_key(key, external_value_provider)
                    {
                        for sheet_key in expanded {
                            precedents.insert(PrecedentNode::ExternalCell {
                                sheet: sheet_key,
                                addr,
                            });
                        }
                    }
                }
            }
        }
        Expr::RangeRef(r) => {
            if let SheetReference::External(key) = &r.sheet {
                if crate::eval::is_valid_external_sheet_key(key) {
                    let Some(start) = r.start.resolve(current_cell.addr) else {
                        return;
                    };
                    let Some(end) = r.end.resolve(current_cell.addr) else {
                        return;
                    };
                    precedents.insert(PrecedentNode::ExternalRange {
                        sheet: key.clone(),
                        start: clamp_addr_to_excel_dimensions(start),
                        end: clamp_addr_to_excel_dimensions(end),
                    });
                } else if crate::eval::split_external_sheet_span_key(key).is_some() {
                    let Some(start) = r.start.resolve(current_cell.addr) else {
                        return;
                    };
                    let Some(end) = r.end.resolve(current_cell.addr) else {
                        return;
                    };
                    let start = clamp_addr_to_excel_dimensions(start);
                    let end = clamp_addr_to_excel_dimensions(end);
                    if let Some(expanded) =
                        expand_external_sheet_span_key(key, external_value_provider)
                    {
                        for sheet_key in expanded {
                            precedents.insert(PrecedentNode::ExternalRange {
                                sheet: sheet_key,
                                start,
                                end,
                            });
                        }
                    }
                }
            }
        }
        Expr::StructuredRef(sref_expr) => {
            let SheetReference::External(key) = &sref_expr.sheet else {
                return;
            };
            if !key.starts_with('[') {
                // `SheetReference::External` without a bracketed workbook prefix represents an
                // invalid/missing sheet at compile time; preserve `#REF!` semantics.
                return;
            }

            // External workbook structured references (e.g. `[Book.xlsx]Sheet1!Table1[Col]`) are
            // resolved dynamically using provider-supplied table metadata.
            let provider = match external_value_provider {
                Some(p) => p,
                None => return,
            };

            let (workbook, explicit_sheet_key) = match crate::eval::split_external_sheet_key(key) {
                Some((workbook, sheet)) if !sheet.contains(':') => (workbook, Some(key.as_str())),
                Some((_workbook, _sheet)) => {
                    // External 3D sheet spans are not valid structured-ref prefixes.
                    return;
                }
                None => {
                    // Workbook-only external reference (`[Book.xlsx]...`); parse the bracketed
                    // workbook prefix.
                    let Some(end) = key.rfind(']') else {
                        return;
                    };
                    let workbook = key.get(1..end).unwrap_or_default();
                    if workbook.is_empty() {
                        return;
                    }
                    (workbook, None)
                }
            };

            let Some(table_name) = sref_expr.sref.table_name.as_deref() else {
                return;
            };

            // Excel's `[@ThisRow]` semantics depend on the formula being inside the table. For
            // external workbooks we do not currently model the row context, so preserve `#REF!`
            // behavior by skipping precedent expansion.
            if sref_expr
                .sref
                .items
                .iter()
                .any(|item| matches!(item, crate::structured_refs::StructuredRefItem::ThisRow))
            {
                return;
            }

            let Some((table_sheet, table)) = provider.workbook_table(workbook, table_name) else {
                return;
            };

            let sheet_key = explicit_sheet_key
                .map(|s| s.to_string())
                .unwrap_or_else(|| format!("[{workbook}]{table_sheet}"));

            let ranges = match crate::structured_refs::resolve_structured_ref_in_table(
                &table,
                current_cell.addr,
                &sref_expr.sref,
            ) {
                Ok(ranges) => ranges,
                Err(_) => return,
            };

            for (start, end) in ranges {
                let start = clamp_addr_to_excel_dimensions(start);
                let end = clamp_addr_to_excel_dimensions(end);
                if start == end {
                    precedents.insert(PrecedentNode::ExternalCell {
                        sheet: sheet_key.clone(),
                        addr: start,
                    });
                } else {
                    precedents.insert(PrecedentNode::ExternalRange {
                        sheet: sheet_key.clone(),
                        start,
                        end,
                    });
                }
            }
        }
        Expr::NameRef(nref) => {
            let Some(sheet) = resolve_sheet(&nref.sheet, current_cell.sheet) else {
                return;
            };
            let name_key = normalize_defined_name(&nref.name);
            if name_key.is_empty() {
                return;
            }

            // LET/LAMBDA lexical bindings are only visible for unqualified identifiers.
            // Explicit sheet-qualified names (e.g. `Sheet1!X`) should still resolve as defined
            // names and surface any external precedents.
            if matches!(nref.sheet, SheetReference::Current)
                && name_is_local(lexical_scopes, &name_key)
            {
                return;
            }

            let visit_key = (sheet, name_key.clone());
            if !visiting_names.insert(visit_key.clone()) {
                return;
            }
            if let Some(def) = resolve_defined_name(workbook, sheet, &name_key) {
                if let Some(expr) = def.compiled.as_ref() {
                    walk_external_expr(
                        expr,
                        CellKey {
                            sheet,
                            addr: current_cell.addr,
                        },
                        workbook,
                        external_value_provider,
                        precedents,
                        visiting_names,
                        lexical_scopes,
                    );
                }
            }
            visiting_names.remove(&visit_key);
        }
        Expr::FieldAccess { base, .. } => walk_external_expr(
            base,
            current_cell,
            workbook,
            external_value_provider,
            precedents,
            visiting_names,
            lexical_scopes,
        ),
        Expr::Unary { expr, .. }
        | Expr::Postfix { expr, .. }
        | Expr::ImplicitIntersection(expr)
        | Expr::SpillRange(expr) => walk_external_expr(
            expr,
            current_cell,
            workbook,
            external_value_provider,
            precedents,
            visiting_names,
            lexical_scopes,
        ),
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            walk_external_expr(
                left,
                current_cell,
                workbook,
                external_value_provider,
                precedents,
                visiting_names,
                lexical_scopes,
            );
            walk_external_expr(
                right,
                current_cell,
                workbook,
                external_value_provider,
                precedents,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::FunctionCall { name, args, .. } => {
            if let Some(spec) = crate::functions::lookup_function(name) {
                match spec.name {
                    "LET" => {
                        if args.len() < 3 || args.len() % 2 == 0 {
                            return;
                        }

                        lexical_scopes.push(HashSet::new());
                        for pair in args[..args.len() - 1].chunks_exact(2) {
                            let Some(name_key) = bare_identifier(&pair[0]) else {
                                lexical_scopes.pop();
                                return;
                            };
                            walk_external_expr(
                                &pair[1],
                                current_cell,
                                workbook,
                                external_value_provider,
                                precedents,
                                visiting_names,
                                lexical_scopes,
                            );
                            lexical_scopes
                                .last_mut()
                                .expect("pushed scope")
                                .insert(name_key);
                        }

                        walk_external_expr(
                            &args[args.len() - 1],
                            current_cell,
                            workbook,
                            external_value_provider,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    "LAMBDA" => {
                        if args.is_empty() {
                            return;
                        }

                        let mut scope = HashSet::new();
                        for param in &args[..args.len() - 1] {
                            let Some(name_key) = bare_identifier(param) else {
                                return;
                            };
                            if !scope.insert(name_key) {
                                return;
                            }
                        }

                        lexical_scopes.push(scope);
                        walk_external_expr(
                            &args[args.len() - 1],
                            current_cell,
                            workbook,
                            external_value_provider,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    _ => {}
                }
            } else {
                let name_key = normalize_defined_name(name);
                let is_local = !name_key.is_empty() && name_is_local(lexical_scopes, &name_key);
                if !name_key.is_empty() && !is_local {
                    let sheet = current_cell.sheet;
                    let visit_key = (sheet, name_key.clone());
                    if visiting_names.insert(visit_key.clone()) {
                        if let Some(def) = resolve_defined_name(workbook, sheet, &name_key) {
                            if let Some(expr) = def.compiled.as_ref() {
                                walk_external_expr(
                                    expr,
                                    CellKey {
                                        sheet,
                                        addr: current_cell.addr,
                                    },
                                    workbook,
                                    external_value_provider,
                                    precedents,
                                    visiting_names,
                                    lexical_scopes,
                                );
                            }
                        }
                        visiting_names.remove(&visit_key);
                    }
                }
            }

            for a in args {
                walk_external_expr(
                    a,
                    current_cell,
                    workbook,
                    external_value_provider,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::Call { callee, args } => {
            walk_external_expr(
                callee,
                current_cell,
                workbook,
                external_value_provider,
                precedents,
                visiting_names,
                lexical_scopes,
            );
            for a in args {
                walk_external_expr(
                    a,
                    current_cell,
                    workbook,
                    external_value_provider,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::ArrayLiteral { values, .. } => {
            for el in values.iter() {
                walk_external_expr(
                    el,
                    current_cell,
                    workbook,
                    external_value_provider,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Blank | Expr::Error(_) => {}
    }
}

fn spill_range_target_cell(expr: &CompiledExpr, current_cell: CellKey) -> Option<CellKey> {
    match expr {
        Expr::CellRef(r) => {
            let sheet = resolve_single_sheet(&r.sheet, current_cell.sheet)?;
            let addr = r.addr.resolve(current_cell.addr)?;
            Some(CellKey { sheet, addr })
        }
        Expr::FieldAccess { base, .. } => spill_range_target_cell(base, current_cell),
        Expr::ImplicitIntersection(inner) | Expr::SpillRange(inner) => {
            spill_range_target_cell(inner, current_cell)
        }
        _ => None,
    }
}

fn spill_range_bounds(cell: CellKey, spills: &SpillState) -> (CellKey, CellAddr) {
    let origin = if spills.by_origin.contains_key(&cell) {
        cell
    } else {
        spills.origin_by_cell.get(&cell).copied().unwrap_or(cell)
    };
    let end = spills
        .by_origin
        .get(&origin)
        .map(|spill| spill.end)
        .unwrap_or(origin.addr);
    (origin, end)
}

fn walk_calc_expr(
    expr: &CompiledExpr,
    current_cell: CellKey,
    tables_by_sheet: &[Vec<Table>],
    workbook: &Workbook,
    spills: &SpillState,
    precedents: &mut HashSet<Precedent>,
    visiting_names: &mut HashSet<(SheetId, String)>,
    lexical_scopes: &mut Vec<HashSet<String>>,
) {
    fn name_is_local(scopes: &[HashSet<String>], name_key: &str) -> bool {
        scopes.iter().rev().any(|scope| scope.contains(name_key))
    }

    fn bare_identifier(expr: &CompiledExpr) -> Option<String> {
        match expr {
            Expr::NameRef(nref) if matches!(nref.sheet, SheetReference::Current) => {
                Some(normalize_defined_name(&nref.name))
            }
            _ => None,
        }
    }

    fn is_direct_self_reference(expr: &CompiledExpr, current_cell: CellKey) -> bool {
        match expr {
            Expr::CellRef(r) => {
                let Some(sheet) = resolve_single_sheet(&r.sheet, current_cell.sheet) else {
                    return false;
                };
                if sheet != current_cell.sheet {
                    return false;
                }
                let Some(addr) = r.addr.resolve(current_cell.addr) else {
                    return false;
                };
                addr == current_cell.addr
            }
            Expr::RangeRef(RangeRef { sheet, start, end }) => {
                let Some(sheet_id) = resolve_single_sheet(sheet, current_cell.sheet) else {
                    return false;
                };
                if sheet_id != current_cell.sheet {
                    return false;
                }
                let Some(start) = start.resolve(current_cell.addr) else {
                    return false;
                };
                let Some(end) = end.resolve(current_cell.addr) else {
                    return false;
                };
                start == current_cell.addr && end == current_cell.addr
            }
            _ => false,
        }
    }

    fn walk_calc_expr_reference_context(
        expr: &CompiledExpr,
        current_cell: CellKey,
        tables_by_sheet: &[Vec<Table>],
        workbook: &Workbook,
        spills: &SpillState,
        precedents: &mut HashSet<Precedent>,
        visiting_names: &mut HashSet<(SheetId, String)>,
        lexical_scopes: &mut Vec<HashSet<String>>,
    ) {
        match expr {
            // These functions only care about the reference's address/shape, not the referenced
            // cell values. Avoid introducing calc-graph precedents for direct references so
            // self-referential formulas like `=ROW(A1)` in `A1` aren't treated as circular.
            Expr::CellRef(_) | Expr::RangeRef(_) | Expr::StructuredRef(_) => {}
            Expr::NameRef(nref) => {
                let Some(sheet) = resolve_single_sheet(&nref.sheet, current_cell.sheet) else {
                    return;
                };
                let name_key = normalize_defined_name(&nref.name);
                if name_key.is_empty() {
                    return;
                }
                // LET/LAMBDA lexical bindings are only visible for unqualified identifiers.
                // Explicit sheet-qualified names (e.g. `Sheet1!X`) should still resolve as defined
                // names for dependency analysis and dirty propagation.
                if matches!(nref.sheet, SheetReference::Current)
                    && name_is_local(lexical_scopes, &name_key)
                {
                    return;
                }
                let visit_key = (sheet, name_key.clone());
                if !visiting_names.insert(visit_key.clone()) {
                    return;
                }
                if let Some(def) = resolve_defined_name(workbook, sheet, &name_key) {
                    if let Some(expr) = def.compiled.as_ref() {
                        walk_calc_expr_reference_context(
                            expr,
                            CellKey {
                                sheet,
                                addr: current_cell.addr,
                            },
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                    }
                }
                visiting_names.remove(&visit_key);
            }
            Expr::ImplicitIntersection(inner) => {
                // Treat implicit-intersection wrappers over direct references like the references
                // themselves so `ROW(@A1)` (or similar compiler-inserted wrappers) doesn't create
                // a spurious calc-graph cycle.
                match inner.as_ref() {
                    Expr::CellRef(_) | Expr::RangeRef(_) | Expr::StructuredRef(_) => {}
                    other => walk_calc_expr_reference_context(
                        other,
                        current_cell,
                        tables_by_sheet,
                        workbook,
                        spills,
                        precedents,
                        visiting_names,
                        lexical_scopes,
                    ),
                }
            }
            Expr::FunctionCall { name, args, .. } if name == "OFFSET" => {
                if args.is_empty() {
                    return;
                }

                // OFFSET's base reference is used for its address; the row/col/size arguments
                // determine the returned reference and should participate in dependency analysis.
                walk_calc_expr_reference_context(
                    &args[0],
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    spills,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
                for a in args.iter().skip(1) {
                    walk_calc_expr(
                        a,
                        current_cell,
                        tables_by_sheet,
                        workbook,
                        spills,
                        precedents,
                        visiting_names,
                        lexical_scopes,
                    );
                }
            }
            Expr::FunctionCall { name, args, .. } if name == "CHOOSE" => {
                if args.is_empty() {
                    return;
                }

                // CHOOSE can return references depending on surrounding context (see bytecode VM /
                // AST evaluator semantics). In reference context (e.g. ROW(CHOOSE(...))), the
                // selected choice expression should be treated as producing a reference rather than
                // dereferencing cell values. Otherwise, self-referential formulas like
                // `=ROW(CHOOSE(1, A1, B1))` entered into `A1` become spurious calc-graph cycles.
                //
                // The index argument itself is value-driven and should participate in dependency
                // analysis.
                walk_calc_expr(
                    &args[0],
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    spills,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
                for choice in args.iter().skip(1) {
                    walk_calc_expr_reference_context(
                        choice,
                        current_cell,
                        tables_by_sheet,
                        workbook,
                        spills,
                        precedents,
                        visiting_names,
                        lexical_scopes,
                    );
                }
            }
            Expr::FunctionCall { name, args, .. } if name == "INDEX" => {
                if args.is_empty() {
                    return;
                }

                // When INDEX is used in reference context (e.g. ROW(INDEX(...))), the input range
                // is used only for its bounds/shape; the row/col/area arguments determine the
                // returned reference and should participate in dependency analysis.
                walk_calc_expr_reference_context(
                    &args[0],
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    spills,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
                for a in args.iter().skip(1) {
                    walk_calc_expr(
                        a,
                        current_cell,
                        tables_by_sheet,
                        workbook,
                        spills,
                        precedents,
                        visiting_names,
                        lexical_scopes,
                    );
                }
            }
            // Spilled ranges are dynamic; consumers like ROW/COLUMN depend on the spill bounds.
            Expr::SpillRange(_) => walk_calc_expr(
                expr,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            ),
            other => walk_calc_expr(
                other,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            ),
        }
    }

    match expr {
        Expr::CellRef(r) => {
            if let Some(sheets) = resolve_sheet_span(&r.sheet, current_cell.sheet, workbook) {
                let Some(addr) = r.addr.resolve(current_cell.addr) else {
                    return;
                };
                for sheet in sheets {
                    precedents.insert(Precedent::Cell(CellId::new(
                        sheet_id_for_graph(sheet),
                        addr.row,
                        addr.col,
                    )));
                }
            }
        }
        Expr::RangeRef(RangeRef { sheet, start, end }) => {
            if let Some(sheets) = resolve_sheet_span(sheet, current_cell.sheet, workbook) {
                let Some(start) = start.resolve(current_cell.addr) else {
                    return;
                };
                let Some(end) = end.resolve(current_cell.addr) else {
                    return;
                };
                let range = Range::new(
                    CellRef::new(start.row, start.col),
                    CellRef::new(end.row, end.col),
                );
                for sheet_id in sheets {
                    precedents.insert(Precedent::Range(SheetRange::new(
                        sheet_id_for_graph(sheet_id),
                        range,
                    )));
                }
            }
        }
        Expr::StructuredRef(sref_expr) => {
            // Only local structured refs participate in the dependency graph. External workbook
            // structured refs are resolved dynamically through the external value provider and are
            // treated as volatile rather than producing calc precedents.
            if matches!(&sref_expr.sheet, SheetReference::External(_)) {
                return;
            }

            if let Ok(ranges) = crate::structured_refs::resolve_structured_ref(
                tables_by_sheet,
                current_cell.sheet,
                current_cell.addr,
                &sref_expr.sref,
            ) {
                for (sheet_id, start, end) in ranges {
                    let range = Range::new(
                        CellRef::new(start.row, start.col),
                        CellRef::new(end.row, end.col),
                    );
                    precedents.insert(Precedent::Range(SheetRange::new(
                        sheet_id_for_graph(sheet_id),
                        range,
                    )));
                }
            }
        }
        Expr::NameRef(nref) => {
            let Some(sheet) = resolve_single_sheet(&nref.sheet, current_cell.sheet) else {
                return;
            };
            let name_key = normalize_defined_name(&nref.name);
            if name_key.is_empty() {
                return;
            }
            // LET/LAMBDA lexical bindings are only visible for unqualified identifiers.
            // Explicit sheet-qualified names (e.g. `Sheet1!X`) should still resolve as defined
            // names for dependency analysis and dirty propagation.
            if matches!(nref.sheet, SheetReference::Current)
                && name_is_local(lexical_scopes, &name_key)
            {
                return;
            }
            let visit_key = (sheet, name_key.clone());
            if !visiting_names.insert(visit_key.clone()) {
                return;
            }
            if let Some(def) = resolve_defined_name(workbook, sheet, &name_key) {
                if let Some(expr) = def.compiled.as_ref() {
                    walk_calc_expr(
                        expr,
                        CellKey {
                            sheet,
                            addr: current_cell.addr,
                        },
                        tables_by_sheet,
                        workbook,
                        spills,
                        precedents,
                        visiting_names,
                        lexical_scopes,
                    );
                }
            } else {
                // If there's no defined name, Excel allows a bare table name (e.g. `=Table1`) which
                // resolves to the table's default data area. Treat those as structured references
                // for dependency analysis so table edits mark dependent formulas dirty.
                let candidate = nref.name.trim();
                if !candidate.is_empty() {
                    let sref = crate::structured_refs::StructuredRef {
                        table_name: Some(candidate.to_string()),
                        items: Vec::new(),
                        columns: crate::structured_refs::StructuredColumns::All,
                    };
                    if let Ok(ranges) = crate::structured_refs::resolve_structured_ref(
                        tables_by_sheet,
                        current_cell.sheet,
                        current_cell.addr,
                        &sref,
                    ) {
                        for (sheet_id, start, end) in ranges {
                            let range = Range::new(
                                CellRef::new(start.row, start.col),
                                CellRef::new(end.row, end.col),
                            );
                            precedents.insert(Precedent::Range(SheetRange::new(
                                sheet_id_for_graph(sheet_id),
                                range,
                            )));
                        }
                    }
                }
            }
            visiting_names.remove(&visit_key);
        }
        Expr::SpillRange(inner) => {
            if let Some(target) = spill_range_target_cell(inner, current_cell) {
                let (origin, end) = spill_range_bounds(target, spills);
                if origin.addr == end {
                    precedents.insert(Precedent::Cell(CellId::new(
                        sheet_id_for_graph(origin.sheet),
                        origin.addr.row,
                        origin.addr.col,
                    )));
                } else {
                    let range = Range::new(
                        CellRef::new(origin.addr.row, origin.addr.col),
                        CellRef::new(end.row, end.col),
                    );
                    precedents.insert(Precedent::Range(SheetRange::new(
                        sheet_id_for_graph(origin.sheet),
                        range,
                    )));
                }
            }
            walk_calc_expr(
                inner,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::FieldAccess { base, .. } => {
            walk_calc_expr(
                base,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::Unary { expr, .. } | Expr::Postfix { expr, .. } => {
            walk_calc_expr(
                expr,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::Binary { left, right, .. } | Expr::Compare { left, right, .. } => {
            walk_calc_expr(
                left,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
            walk_calc_expr(
                right,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
        }
        Expr::FunctionCall { name, args, .. } => {
            if let Some(spec) = crate::functions::lookup_function(name) {
                match spec.name {
                    "CELL" => {
                        if args.is_empty() {
                            return;
                        }

                        // `CELL(info_type, reference)` is unusual: for many `info_type` values,
                        // the `reference` argument is used only for its *address* (sheet/row/col)
                        // and should not create calculation dependencies on the referenced cells'
                        // values.
                        //
                        // Without this special-casing, a formula like `=CELL("width", A1)` entered
                        // into `A1` becomes a self-edge in the calc graph and is treated as a
                        // circular reference, even though the column width does not depend on the
                        // cell's value.
                        //
                        // When `info_type` is not a compile-time literal, fall back to the generic
                        // dependency walker.
                        // Treat formulas as untrusted input: avoid allocating a full lowercased
                        // copy of `info_type` (a string literal) just to do a couple
                        // case-insensitive comparisons.
                        let info_type_literal = match &args[0] {
                            Expr::Text(s) => Some(s.trim()),
                            _ => None,
                        };

                        // Always walk `info_type`; it can itself contain references (e.g. `CELL(A1, ...)`).
                        walk_calc_expr(
                            &args[0],
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );

                        if args.len() < 2 {
                            return;
                        }

                        if let Some(info_type) = info_type_literal {
                            let derefs_reference_value = info_type.eq_ignore_ascii_case("contents")
                                || info_type.eq_ignore_ascii_case("type");

                            if derefs_reference_value {
                                // `CELL("contents" | "type", reference)` consults the value/type
                                // (and stored formula text for "contents") of the *upper-left*
                                // cell in `reference` (Excel behavior). Represent direct
                                // references as single-cell precedents to avoid range-node cycles
                                // when the formula cell is inside the referenced range.
                                let insert_cell =
                                    |precedents: &mut HashSet<Precedent>,
                                     sheet_id: SheetId,
                                     addr: CellAddr| {
                                        if sheet_id == current_cell.sheet
                                            && addr == current_cell.addr
                                        {
                                            return;
                                        }
                                        precedents.insert(Precedent::Cell(CellId::new(
                                            sheet_id_for_graph(sheet_id),
                                            addr.row,
                                            addr.col,
                                        )));
                                    };

                                match &args[1] {
                                    Expr::CellRef(r) => {
                                        if let Some(sheets) = resolve_sheet_span(
                                            &r.sheet,
                                            current_cell.sheet,
                                            workbook,
                                        ) {
                                            let Some(addr) = r.addr.resolve(current_cell.addr)
                                            else {
                                                return;
                                            };
                                            for sheet_id in sheets {
                                                insert_cell(precedents, sheet_id, addr);
                                            }
                                        }
                                    }
                                    Expr::RangeRef(RangeRef { sheet, start, end }) => {
                                        if let Some(sheets) =
                                            resolve_sheet_span(sheet, current_cell.sheet, workbook)
                                        {
                                            let Some(start) = start.resolve(current_cell.addr)
                                            else {
                                                return;
                                            };
                                            let Some(end) = end.resolve(current_cell.addr) else {
                                                return;
                                            };
                                            let addr = CellAddr {
                                                row: start.row.min(end.row),
                                                col: start.col.min(end.col),
                                            };
                                            for sheet_id in sheets {
                                                insert_cell(precedents, sheet_id, addr);
                                            }
                                        }
                                    }
                                    Expr::StructuredRef(sref_expr) => {
                                        if matches!(&sref_expr.sheet, SheetReference::External(_)) {
                                            return;
                                        }
                                        if let Ok(ranges) =
                                            crate::structured_refs::resolve_structured_ref(
                                                tables_by_sheet,
                                                current_cell.sheet,
                                                current_cell.addr,
                                                &sref_expr.sref,
                                            )
                                        {
                                            for (sheet_id, start, end) in ranges {
                                                let addr = CellAddr {
                                                    row: start.row.min(end.row),
                                                    col: start.col.min(end.col),
                                                };
                                                insert_cell(precedents, sheet_id, addr);
                                            }
                                        }
                                    }
                                    Expr::SpillRange(inner) => {
                                        if let Some(target) =
                                            spill_range_target_cell(inner, current_cell)
                                        {
                                            insert_cell(precedents, target.sheet, target.addr);
                                        } else {
                                            walk_calc_expr_reference_context(
                                                inner,
                                                current_cell,
                                                tables_by_sheet,
                                                workbook,
                                                spills,
                                                precedents,
                                                visiting_names,
                                                lexical_scopes,
                                            );
                                        }
                                    }
                                    Expr::ImplicitIntersection(inner) => match inner.as_ref() {
                                        Expr::CellRef(r) => {
                                            if let Some(sheets) = resolve_sheet_span(
                                                &r.sheet,
                                                current_cell.sheet,
                                                workbook,
                                            ) {
                                                let Some(addr) = r.addr.resolve(current_cell.addr)
                                                else {
                                                    return;
                                                };
                                                for sheet_id in sheets {
                                                    insert_cell(precedents, sheet_id, addr);
                                                }
                                            }
                                        }
                                        Expr::RangeRef(RangeRef { sheet, start, end }) => {
                                            let Some(start) = start.resolve(current_cell.addr)
                                            else {
                                                return;
                                            };
                                            let Some(end) = end.resolve(current_cell.addr) else {
                                                return;
                                            };
                                            let row_start = start.row.min(end.row);
                                            let row_end = start.row.max(end.row);
                                            let col_start = start.col.min(end.col);
                                            let col_end = start.col.max(end.col);
                                            let cur = current_cell.addr;

                                            let intersected =
                                                if row_start == row_end && col_start == col_end {
                                                    Some(CellAddr {
                                                        row: row_start,
                                                        col: col_start,
                                                    })
                                                } else if col_start == col_end {
                                                    (cur.row >= row_start && cur.row <= row_end)
                                                        .then(|| CellAddr {
                                                            row: cur.row,
                                                            col: col_start,
                                                        })
                                                } else if row_start == row_end {
                                                    (cur.col >= col_start && cur.col <= col_end)
                                                        .then(|| CellAddr {
                                                            row: row_start,
                                                            col: cur.col,
                                                        })
                                                } else {
                                                    (cur.row >= row_start
                                                        && cur.row <= row_end
                                                        && cur.col >= col_start
                                                        && cur.col <= col_end)
                                                        .then(|| cur)
                                                };

                                            if let (Some(intersected), Some(sheets)) = (
                                                intersected,
                                                resolve_sheet_span(
                                                    sheet,
                                                    current_cell.sheet,
                                                    workbook,
                                                ),
                                            ) {
                                                for sheet_id in sheets {
                                                    insert_cell(precedents, sheet_id, intersected);
                                                }
                                            }
                                        }
                                        Expr::StructuredRef(sref_expr) => {
                                            if matches!(
                                                &sref_expr.sheet,
                                                SheetReference::External(_)
                                            ) {
                                                return;
                                            }
                                            if let Ok(ranges) =
                                                crate::structured_refs::resolve_structured_ref(
                                                    tables_by_sheet,
                                                    current_cell.sheet,
                                                    current_cell.addr,
                                                    &sref_expr.sref,
                                                )
                                            {
                                                for (sheet_id, start, end) in ranges {
                                                    let addr = CellAddr {
                                                        row: start.row.min(end.row),
                                                        col: start.col.min(end.col),
                                                    };
                                                    insert_cell(precedents, sheet_id, addr);
                                                }
                                            }
                                        }
                                        Expr::SpillRange(inner) => {
                                            if let Some(target) =
                                                spill_range_target_cell(inner, current_cell)
                                            {
                                                insert_cell(precedents, target.sheet, target.addr);
                                            } else {
                                                walk_calc_expr_reference_context(
                                                    inner,
                                                    current_cell,
                                                    tables_by_sheet,
                                                    workbook,
                                                    spills,
                                                    precedents,
                                                    visiting_names,
                                                    lexical_scopes,
                                                );
                                            }
                                        }
                                        other => {
                                            walk_calc_expr_reference_context(
                                                other,
                                                current_cell,
                                                tables_by_sheet,
                                                workbook,
                                                spills,
                                                precedents,
                                                visiting_names,
                                                lexical_scopes,
                                            );
                                        }
                                    },
                                    other => {
                                        walk_calc_expr_reference_context(
                                            other,
                                            current_cell,
                                            tables_by_sheet,
                                            workbook,
                                            spills,
                                            precedents,
                                            visiting_names,
                                            lexical_scopes,
                                        );
                                    }
                                }

                                for a in args.iter().skip(2) {
                                    walk_calc_expr(
                                        a,
                                        current_cell,
                                        tables_by_sheet,
                                        workbook,
                                        spills,
                                        precedents,
                                        visiting_names,
                                        lexical_scopes,
                                    );
                                }
                                return;
                            }

                            // Address-only `info_type` values (e.g. width/format/address) use the
                            // reference argument only for its address/shape. Walk it in
                            // "reference context" so we pick up dependencies used to *compute* the
                            // reference without introducing precedents on referenced cell values
                            // (avoids spurious circular references).
                            //
                            // Spill ranges (`A1#`) are dynamic, but these address-only CELL keys
                            // consult only the upper-left address. Treat direct spill references as
                            // address-only too to avoid pulling in dynamic spill dependencies.
                            match &args[1] {
                                Expr::SpillRange(_) => {}
                                Expr::ImplicitIntersection(inner)
                                    if matches!(inner.as_ref(), Expr::SpillRange(_)) => {}
                                other => walk_calc_expr_reference_context(
                                    other,
                                    current_cell,
                                    tables_by_sheet,
                                    workbook,
                                    spills,
                                    precedents,
                                    visiting_names,
                                    lexical_scopes,
                                ),
                            }

                            for a in args.iter().skip(2) {
                                walk_calc_expr(
                                    a,
                                    current_cell,
                                    tables_by_sheet,
                                    workbook,
                                    spills,
                                    precedents,
                                    visiting_names,
                                    lexical_scopes,
                                );
                            }
                            return;
                        }
                    }
                    "LET" => {
                        if args.len() < 3 || args.len() % 2 == 0 {
                            return;
                        }

                        lexical_scopes.push(HashSet::new());
                        for pair in args[..args.len() - 1].chunks_exact(2) {
                            let Some(name_key) = bare_identifier(&pair[0]) else {
                                lexical_scopes.pop();
                                return;
                            };
                            walk_calc_expr(
                                &pair[1],
                                current_cell,
                                tables_by_sheet,
                                workbook,
                                spills,
                                precedents,
                                visiting_names,
                                lexical_scopes,
                            );
                            lexical_scopes
                                .last_mut()
                                .expect("pushed scope")
                                .insert(name_key);
                        }

                        walk_calc_expr(
                            &args[args.len() - 1],
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    "LAMBDA" => {
                        if args.is_empty() {
                            return;
                        }

                        let mut scope = HashSet::new();
                        for param in &args[..args.len() - 1] {
                            let Some(name_key) = bare_identifier(param) else {
                                return;
                            };
                            if !scope.insert(name_key) {
                                return;
                            }
                        }

                        lexical_scopes.push(scope);
                        walk_calc_expr(
                            &args[args.len() - 1],
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                        lexical_scopes.pop();
                        return;
                    }
                    "FORMULATEXT" | "ISFORMULA" => {
                        let Some(arg0) = args.first() else {
                            return;
                        };

                        // These functions consult worksheet *metadata* (stored formula text /
                        // whether a cell has a formula), not the evaluated value of the referenced
                        // cell. Direct self-references like `=FORMULATEXT(A1)` in `A1` should not be
                        // treated as circular.
                        if is_direct_self_reference(arg0, current_cell) {
                            return;
                        }

                        // These functions accept a reference argument but only consult its metadata,
                        // not its evaluated value. Avoid introducing range-node cycles by
                        // representing direct references as single-cell precedents.
                        let insert_cell =
                            |precedents: &mut HashSet<Precedent>,
                             sheet_id: SheetId,
                             addr: CellAddr| {
                                if sheet_id == current_cell.sheet && addr == current_cell.addr {
                                    return;
                                }
                                precedents.insert(Precedent::Cell(CellId::new(
                                    sheet_id_for_graph(sheet_id),
                                    addr.row,
                                    addr.col,
                                )));
                            };

                        match arg0 {
                            Expr::CellRef(r) => {
                                if let Some(sheets) =
                                    resolve_sheet_span(&r.sheet, current_cell.sheet, workbook)
                                {
                                    let Some(addr) = r.addr.resolve(current_cell.addr) else {
                                        return;
                                    };
                                    for sheet_id in sheets {
                                        insert_cell(precedents, sheet_id, addr);
                                    }
                                }
                            }
                            Expr::RangeRef(RangeRef { sheet, start, end }) => {
                                if let Some(sheets) =
                                    resolve_sheet_span(sheet, current_cell.sheet, workbook)
                                {
                                    let Some(start) = start.resolve(current_cell.addr) else {
                                        return;
                                    };
                                    let Some(end) = end.resolve(current_cell.addr) else {
                                        return;
                                    };
                                    let addr = CellAddr {
                                        row: start.row.min(end.row),
                                        col: start.col.min(end.col),
                                    };
                                    for sheet_id in sheets {
                                        insert_cell(precedents, sheet_id, addr);
                                    }
                                }
                            }
                            Expr::StructuredRef(sref_expr) => {
                                if matches!(&sref_expr.sheet, SheetReference::External(_)) {
                                    return;
                                }
                                if let Ok(ranges) = crate::structured_refs::resolve_structured_ref(
                                    tables_by_sheet,
                                    current_cell.sheet,
                                    current_cell.addr,
                                    &sref_expr.sref,
                                ) {
                                    for (sheet_id, start, end) in ranges {
                                        let addr = CellAddr {
                                            row: start.row.min(end.row),
                                            col: start.col.min(end.col),
                                        };
                                        insert_cell(precedents, sheet_id, addr);
                                    }
                                }
                            }
                            Expr::SpillRange(inner) => {
                                if let Some(target) = spill_range_target_cell(inner, current_cell) {
                                    insert_cell(precedents, target.sheet, target.addr);
                                } else {
                                    walk_calc_expr_reference_context(
                                        inner,
                                        current_cell,
                                        tables_by_sheet,
                                        workbook,
                                        spills,
                                        precedents,
                                        visiting_names,
                                        lexical_scopes,
                                    );
                                }
                            }
                            Expr::ImplicitIntersection(inner) => match inner.as_ref() {
                                Expr::CellRef(r) => {
                                    if let Some(sheets) =
                                        resolve_sheet_span(&r.sheet, current_cell.sheet, workbook)
                                    {
                                        let Some(addr) = r.addr.resolve(current_cell.addr) else {
                                            return;
                                        };
                                        for sheet_id in sheets {
                                            insert_cell(precedents, sheet_id, addr);
                                        }
                                    }
                                }
                                Expr::RangeRef(RangeRef { sheet, start, end }) => {
                                    let Some(start) = start.resolve(current_cell.addr) else {
                                        return;
                                    };
                                    let Some(end) = end.resolve(current_cell.addr) else {
                                        return;
                                    };
                                    let row_start = start.row.min(end.row);
                                    let row_end = start.row.max(end.row);
                                    let col_start = start.col.min(end.col);
                                    let col_end = start.col.max(end.col);
                                    let cur = current_cell.addr;

                                    let intersected =
                                        if row_start == row_end && col_start == col_end {
                                            Some(CellAddr {
                                                row: row_start,
                                                col: col_start,
                                            })
                                        } else if col_start == col_end {
                                            (cur.row >= row_start && cur.row <= row_end).then(
                                                || CellAddr {
                                                    row: cur.row,
                                                    col: col_start,
                                                },
                                            )
                                        } else if row_start == row_end {
                                            (cur.col >= col_start && cur.col <= col_end).then(
                                                || CellAddr {
                                                    row: row_start,
                                                    col: cur.col,
                                                },
                                            )
                                        } else {
                                            (cur.row >= row_start
                                                && cur.row <= row_end
                                                && cur.col >= col_start
                                                && cur.col <= col_end)
                                                .then(|| cur)
                                        };

                                    if let (Some(intersected), Some(sheets)) = (
                                        intersected,
                                        resolve_sheet_span(sheet, current_cell.sheet, workbook),
                                    ) {
                                        for sheet_id in sheets {
                                            insert_cell(precedents, sheet_id, intersected);
                                        }
                                    }
                                }
                                Expr::StructuredRef(sref_expr) => {
                                    if matches!(&sref_expr.sheet, SheetReference::External(_)) {
                                        return;
                                    }
                                    if let Ok(ranges) =
                                        crate::structured_refs::resolve_structured_ref(
                                            tables_by_sheet,
                                            current_cell.sheet,
                                            current_cell.addr,
                                            &sref_expr.sref,
                                        )
                                    {
                                        for (sheet_id, start, end) in ranges {
                                            let addr = CellAddr {
                                                row: start.row.min(end.row),
                                                col: start.col.min(end.col),
                                            };
                                            insert_cell(precedents, sheet_id, addr);
                                        }
                                    }
                                }
                                Expr::SpillRange(inner) => {
                                    if let Some(target) =
                                        spill_range_target_cell(inner, current_cell)
                                    {
                                        insert_cell(precedents, target.sheet, target.addr);
                                    } else {
                                        walk_calc_expr_reference_context(
                                            inner,
                                            current_cell,
                                            tables_by_sheet,
                                            workbook,
                                            spills,
                                            precedents,
                                            visiting_names,
                                            lexical_scopes,
                                        );
                                    }
                                }
                                other => {
                                    walk_calc_expr_reference_context(
                                        other,
                                        current_cell,
                                        tables_by_sheet,
                                        workbook,
                                        spills,
                                        precedents,
                                        visiting_names,
                                        lexical_scopes,
                                    );
                                }
                            },
                            other => {
                                walk_calc_expr_reference_context(
                                    other,
                                    current_cell,
                                    tables_by_sheet,
                                    workbook,
                                    spills,
                                    precedents,
                                    visiting_names,
                                    lexical_scopes,
                                );
                            }
                        }

                        for a in args.iter().skip(1) {
                            walk_calc_expr(
                                a,
                                current_cell,
                                tables_by_sheet,
                                workbook,
                                spills,
                                precedents,
                                visiting_names,
                                lexical_scopes,
                            );
                        }
                        return;
                    }
                    "ROW" | "COLUMN" | "ROWS" | "COLUMNS" | "AREAS" | "SHEET" | "SHEETS" => {
                        let Some(arg0) = args.first() else {
                            return;
                        };
                        walk_calc_expr_reference_context(
                            arg0,
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                        return;
                    }
                    "ISREF" => {
                        let Some(arg0) = args.first() else {
                            return;
                        };
                        // ISREF is based on whether the argument is a reference, not the
                        // referenced cell's value.
                        walk_calc_expr_reference_context(
                            arg0,
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                        return;
                    }
                    _ => {}
                }
            } else {
                let name_key = normalize_defined_name(name);
                if !name_key.is_empty() && !name_is_local(lexical_scopes, &name_key) {
                    let sheet = current_cell.sheet;
                    let visit_key = (sheet, name_key.clone());
                    if visiting_names.insert(visit_key.clone()) {
                        if let Some(def) = resolve_defined_name(workbook, sheet, &name_key) {
                            if let Some(expr) = def.compiled.as_ref() {
                                walk_calc_expr(
                                    expr,
                                    CellKey {
                                        sheet,
                                        addr: current_cell.addr,
                                    },
                                    tables_by_sheet,
                                    workbook,
                                    spills,
                                    precedents,
                                    visiting_names,
                                    lexical_scopes,
                                );
                            }
                        }
                        visiting_names.remove(&visit_key);
                    }
                }
            }

            if let Some(spec) = crate::functions::lookup_function(name) {
                // `CELL("width", ref)` consults column metadata for `ref` but does not depend on the
                // contents of `ref`. Avoid registering plain reference literals as calc precedents
                // so formulas like `A1 = CELL("width", A1)` do not create spurious circular
                // references.
                if spec.name == "CELL"
                    && args.len() >= 2
                    && matches!(&args[0], Expr::Text(info) if info.trim().eq_ignore_ascii_case("width"))
                    && {
                        match &args[1] {
                            Expr::CellRef(_)
                            | Expr::RangeRef(_)
                            | Expr::StructuredRef(_)
                            | Expr::SpillRange(_) => true,
                            Expr::ImplicitIntersection(inner) => matches!(
                                inner.as_ref(),
                                Expr::CellRef(_)
                                    | Expr::RangeRef(_)
                                    | Expr::StructuredRef(_)
                                    | Expr::SpillRange(_)
                            ),
                            _ => false,
                        }
                    }
                {
                    for a in args.iter().skip(2) {
                        walk_calc_expr(
                            a,
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                    }
                    return;
                }
            }
            for a in args {
                walk_calc_expr(
                    a,
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    spills,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::Call { callee, args } => {
            walk_calc_expr(
                callee,
                current_cell,
                tables_by_sheet,
                workbook,
                spills,
                precedents,
                visiting_names,
                lexical_scopes,
            );
            for a in args {
                walk_calc_expr(
                    a,
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    spills,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::ArrayLiteral { values, .. } => {
            for el in values.iter() {
                walk_calc_expr(
                    el,
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    spills,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                );
            }
        }
        Expr::ImplicitIntersection(inner) => {
            match inner.as_ref() {
                // Implicit intersection over a static range only depends on the single intersected
                // cell (if any), rather than the entire rectangle.
                Expr::RangeRef(RangeRef { sheet, start, end }) => {
                    let Some(start) = start.resolve(current_cell.addr) else {
                        walk_calc_expr(
                            inner,
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                        return;
                    };
                    let Some(end) = end.resolve(current_cell.addr) else {
                        walk_calc_expr(
                            inner,
                            current_cell,
                            tables_by_sheet,
                            workbook,
                            spills,
                            precedents,
                            visiting_names,
                            lexical_scopes,
                        );
                        return;
                    };
                    let row_start = start.row.min(end.row);
                    let row_end = start.row.max(end.row);
                    let col_start = start.col.min(end.col);
                    let col_end = start.col.max(end.col);
                    let cur = current_cell.addr;

                    let intersected = if row_start == row_end && col_start == col_end {
                        Some(CellAddr {
                            row: row_start,
                            col: col_start,
                        })
                    } else if col_start == col_end {
                        (cur.row >= row_start && cur.row <= row_end).then(|| CellAddr {
                            row: cur.row,
                            col: col_start,
                        })
                    } else if row_start == row_end {
                        (cur.col >= col_start && cur.col <= col_end).then(|| CellAddr {
                            row: row_start,
                            col: cur.col,
                        })
                    } else {
                        (cur.row >= row_start
                            && cur.row <= row_end
                            && cur.col >= col_start
                            && cur.col <= col_end)
                            .then(|| cur)
                    };

                    if let (Some(intersected), Some(sheets)) = (
                        intersected,
                        resolve_sheet_span(sheet, current_cell.sheet, workbook),
                    ) {
                        for sheet_id in sheets {
                            precedents.insert(Precedent::Cell(CellId::new(
                                sheet_id_for_graph(sheet_id),
                                intersected.row,
                                intersected.col,
                            )));
                        }
                    }
                }
                _ => walk_calc_expr(
                    inner,
                    current_cell,
                    tables_by_sheet,
                    workbook,
                    spills,
                    precedents,
                    visiting_names,
                    lexical_scopes,
                ),
            }
        }
        Expr::Number(_) | Expr::Text(_) | Expr::Bool(_) | Expr::Blank | Expr::Error(_) => {}
    }
}

fn resolve_sheet(sheet: &SheetReference<usize>, current_sheet: SheetId) -> Option<SheetId> {
    match sheet {
        SheetReference::Current => Some(current_sheet),
        SheetReference::Sheet(id) => Some(*id),
        SheetReference::SheetRange(a, b) => {
            if a == b {
                Some(*a)
            } else {
                None
            }
        }
        SheetReference::External(_) => None,
    }
}

fn resolve_single_sheet(sheet: &SheetReference<usize>, current_sheet: SheetId) -> Option<SheetId> {
    resolve_sheet(sheet, current_sheet)
}

fn resolve_sheet_span(
    sheet: &SheetReference<usize>,
    current_sheet: SheetId,
    workbook: &Workbook,
) -> Option<Vec<SheetId>> {
    match sheet {
        SheetReference::Current => Some(vec![current_sheet]),
        SheetReference::Sheet(id) => workbook.sheet_exists(*id).then(|| vec![*id]),
        SheetReference::SheetRange(a, b) => workbook.sheet_span_ids(*a, *b),
        SheetReference::External(_) => None,
    }
}

fn sheet_id_for_graph(sheet: SheetId) -> u32 {
    sheet.try_into().expect("sheet id exceeds u32")
}

fn cell_id_from_key(key: CellKey) -> CellId {
    CellId::new(sheet_id_for_graph(key.sheet), key.addr.row, key.addr.col)
}

fn cell_key_from_id(id: CellId) -> CellKey {
    CellKey {
        sheet: usize::try_from(id.sheet_id).expect("sheet id exceeds usize"),
        addr: CellAddr {
            row: id.cell.row,
            col: id.cell.col,
        },
    }
}

fn value_delta(old: &Value, new: &Value) -> f64 {
    match (numeric_value(old), numeric_value(new)) {
        (Some(a), Some(b)) => (a - b).abs(),
        _ if old == new => 0.0,
        _ => f64::INFINITY,
    }
}

fn numeric_value(value: &Value) -> Option<f64> {
    match value {
        Value::Number(n) => Some(*n),
        Value::Blank => Some(0.0),
        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        Value::Entity(_) | Value::Record(_) => None,
        // Treat any other non-numeric values as non-numeric so iterative convergence uses
        // `INFINITY` deltas unless the values are identical.
        _ => None,
    }
}

fn normalize_defined_name(name: &str) -> String {
    crate::value::casefold(name.trim())
}

fn rewrite_defined_name_structural(
    engine: &Engine,
    def: &DefinedName,
    ctx_sheet: &str,
    edit: &StructuralEdit,
    sheet_order_indices: &HashMap<String, usize>,
) -> Result<Option<(NameDefinition, CompiledExpr)>, EngineError> {
    let origin = crate::CellAddr::new(0, 0);
    let (new_def, changed) = match &def.definition {
        NameDefinition::Constant(_) => return Ok(None),
        NameDefinition::Reference(formula) => {
            let (new_formula, changed) = rewrite_formula_for_structural_edit_with_resolver(
                formula,
                ctx_sheet,
                origin,
                edit,
                |name| sheet_order_indices.get(&Workbook::sheet_key(name)).copied(),
            );
            (NameDefinition::Reference(new_formula), changed)
        }
        NameDefinition::Formula(formula) => {
            let (new_formula, changed) = rewrite_formula_for_structural_edit_with_resolver(
                formula,
                ctx_sheet,
                origin,
                edit,
                |name| sheet_order_indices.get(&Workbook::sheet_key(name)).copied(),
            );
            (NameDefinition::Formula(new_formula), changed)
        }
    };

    if !changed {
        return Ok(None);
    }

    let formula = match &new_def {
        NameDefinition::Reference(f) | NameDefinition::Formula(f) => f,
        NameDefinition::Constant(_) => unreachable!("handled above"),
    };
    let ast = crate::parse_formula(
        formula,
        crate::ParseOptions {
            locale: crate::LocaleConfig::en_us(),
            reference_style: crate::ReferenceStyle::A1,
            normalize_relative_to: None,
        },
    )?;
    let parsed = lower_ast(&ast, None);
    let compiled = engine.compile_name_expr(&parsed);
    Ok(Some((new_def, compiled)))
}

fn rewrite_defined_name_range_map(
    engine: &Engine,
    def: &DefinedName,
    ctx_sheet: &str,
    edit: &RangeMapEdit,
    sheet_order_indices: &HashMap<String, usize>,
) -> Result<Option<(NameDefinition, CompiledExpr)>, EngineError> {
    let origin = crate::CellAddr::new(0, 0);
    let (new_def, changed) = match &def.definition {
        NameDefinition::Constant(_) => return Ok(None),
        NameDefinition::Reference(formula) => {
            let (new_formula, changed) = rewrite_formula_for_range_map_with_resolver(
                formula,
                ctx_sheet,
                origin,
                edit,
                |name| sheet_order_indices.get(&Workbook::sheet_key(name)).copied(),
            );
            (NameDefinition::Reference(new_formula), changed)
        }
        NameDefinition::Formula(formula) => {
            let (new_formula, changed) = rewrite_formula_for_range_map_with_resolver(
                formula,
                ctx_sheet,
                origin,
                edit,
                |name| sheet_order_indices.get(&Workbook::sheet_key(name)).copied(),
            );
            (NameDefinition::Formula(new_formula), changed)
        }
    };

    if !changed {
        return Ok(None);
    }

    let formula = match &new_def {
        NameDefinition::Reference(f) | NameDefinition::Formula(f) => f,
        NameDefinition::Constant(_) => unreachable!("handled above"),
    };
    let ast = crate::parse_formula(
        formula,
        crate::ParseOptions {
            locale: crate::LocaleConfig::en_us(),
            reference_style: crate::ReferenceStyle::A1,
            normalize_relative_to: None,
        },
    )?;
    let parsed = lower_ast(&ast, None);
    let compiled = engine.compile_name_expr(&parsed);
    Ok(Some((new_def, compiled)))
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::TimeZone;

    #[test]
    fn cell_style_id_persists_on_blank_cells() {
        let mut engine = Engine::new();
        let style_id = engine.intern_style(Style {
            number_format: Some("0.00".to_string()),
            ..Style::default()
        });
        engine
            .set_cell_style_id("Sheet1", "A1", style_id)
            .expect("set style id");

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            None,
            None,
            engine.info.clone(),
            engine.pivot_registry.clone(),
        );

        assert_eq!(snapshot.get_cell_value(sheet_id, addr), Value::Blank);
        assert_eq!(snapshot.cell_style_id(sheet_id, addr), style_id);
    }

    #[test]
    fn set_cell_value_preserves_style_id() {
        let mut engine = Engine::new();
        let style_id = engine.intern_style(Style {
            number_format: Some("0".to_string()),
            ..Style::default()
        });
        engine
            .set_cell_style_id("Sheet1", "A1", style_id)
            .expect("set style");
        engine
            .set_cell_value("Sheet1", "A1", 123.0_f64)
            .expect("set value");

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        let cell = engine.workbook.sheets[sheet_id]
            .cells
            .get(&addr)
            .expect("cell stored");
        assert_eq!(cell.style_id, style_id);
    }

    #[test]
    fn insert_delete_cols_shift_col_properties() {
        let mut engine = Engine::new();
        engine.set_col_width("Sheet1", 2, Some(42.0));

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        assert_eq!(
            engine.workbook.sheets[sheet_id]
                .col_properties
                .get(&2)
                .and_then(|p| p.width),
            Some(42.0)
        );

        engine
            .apply_operation(EditOp::InsertCols {
                sheet: "Sheet1".to_string(),
                col: 1,
                count: 2,
            })
            .expect("insert cols");

        assert!(
            !engine.workbook.sheets[sheet_id]
                .col_properties
                .contains_key(&2),
            "col properties should shift right on insert"
        );
        assert_eq!(
            engine.workbook.sheets[sheet_id]
                .col_properties
                .get(&4)
                .and_then(|p| p.width),
            Some(42.0)
        );

        engine
            .apply_operation(EditOp::DeleteCols {
                sheet: "Sheet1".to_string(),
                col: 1,
                count: 2,
            })
            .expect("delete cols");

        assert!(
            !engine.workbook.sheets[sheet_id]
                .col_properties
                .contains_key(&4),
            "col properties should shift left on delete"
        );
        assert_eq!(
            engine.workbook.sheets[sheet_id]
                .col_properties
                .get(&2)
                .and_then(|p| p.width),
            Some(42.0)
        );
    }

    #[test]
    fn sheet_name_lookup_is_nfkc_and_unicode_case_insensitive() {
        let mut engine = Engine::new();

        // Angstrom sign (U+212B) normalizes to Å (U+00C5) under NFKC.
        engine
            .set_cell_value("Å", "A1", Value::Number(1.0))
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Å").expect("sheet exists");
        assert_eq!(engine.workbook.sheet_ids_in_order().len(), 1);
        assert_eq!(engine.workbook.sheet_name(sheet_id), Some("Å"));

        // Lookup should succeed under Unicode normalization + case-insensitive compare.
        assert_eq!(engine.get_cell_value("Å", "A1"), Value::Number(1.0));

        // Creating/updating via an equivalent name should not create a new sheet and should keep
        // the original display name unchanged.
        engine
            .set_cell_value("Å", "B1", Value::Number(2.0))
            .unwrap();
        assert_eq!(engine.workbook.sheet_ids_in_order().len(), 1);
        assert_eq!(engine.workbook.sheet_name(sheet_id), Some("Å"));
        assert_eq!(engine.get_cell_value("Å", "B1"), Value::Number(2.0));
    }

    #[test]
    fn sheet_name_lookup_matches_unicode_uppercase_rules() {
        let mut engine = Engine::new();

        // Unicode uppercasing maps ß -> SS (Excel matches this).
        engine
            .set_cell_value("ß", "A1", Value::Number(1.0))
            .unwrap();
        engine
            .set_cell_value("SS", "B1", Value::Number(2.0))
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("ß").expect("sheet exists");
        assert_eq!(engine.workbook.sheet_ids_in_order().len(), 1);
        assert_eq!(engine.workbook.sheet_name(sheet_id), Some("ß"));
        assert_eq!(engine.get_cell_value("SS", "A1"), Value::Number(1.0));
        assert_eq!(engine.get_cell_value("ß", "B1"), Value::Number(2.0));
    }

    #[test]
    fn indirect_sheet_lookup_uses_unicode_normalization() {
        let mut engine = Engine::new();

        engine
            .set_cell_value("Å", "A1", Value::Number(10.0))
            .unwrap();
        engine
            .set_cell_formula("Å", "B1", "=SUM(INDIRECT(\"'Å'!A1\"))")
            .unwrap();

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Å", "B1"), Value::Number(10.0));
    }

    #[test]
    fn indirect_constant_external_refs_are_indexed_for_invalidation() {
        let mut engine = Engine::new();

        engine
            .set_cell_formula("Sheet1", "A1", "=INDIRECT(\"[Book.xlsx]Sheet1!B2\")")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").expect("addr");
        let key = CellKey {
            sheet: sheet_id,
            addr,
        };

        assert_eq!(
            engine
                .cell_external_sheet_refs
                .get(&key)
                .expect("cell should have external sheet refs"),
            &HashSet::from_iter([String::from("[Book.xlsx]Sheet1")])
        );
        assert_eq!(
            engine
                .cell_external_workbook_refs
                .get(&key)
                .expect("cell should have external workbook refs"),
            &HashSet::from_iter([String::from("Book.xlsx")])
        );
        assert!(
            engine
                .external_sheet_dependents
                .get("[Book.xlsx]Sheet1")
                .is_some_and(|deps| deps.contains(&key)),
            "reverse index should include the formula cell"
        );
        assert!(
            engine
                .external_workbook_dependents
                .get("Book.xlsx")
                .is_some_and(|deps| deps.contains(&key)),
            "workbook reverse index should include the formula cell"
        );
    }

    #[test]
    fn rename_sheet_rejects_nfkc_case_insensitive_duplicates() {
        let mut workbook = Workbook::default();
        let sheet_a = workbook.ensure_sheet("Å");
        let sheet_b = workbook.ensure_sheet("Data");

        // Renaming "Data" to a normalized-equivalent of "Å" should be rejected.
        let err = workbook
            .rename_sheet(sheet_b, "Å")
            .expect_err("expected duplicate rename to fail");
        assert_eq!(err, WorkbookRenameSheetError::DuplicateName);

        // Workbook state should remain unchanged after the failed rename.
        assert_eq!(workbook.sheet_name(sheet_a), Some("Å"));
        assert_eq!(workbook.sheet_name(sheet_b), Some("Data"));
        assert_eq!(workbook.sheet_id("Å"), Some(sheet_a));
        assert_eq!(workbook.sheet_id("Å"), Some(sheet_a));
    }

    #[test]
    fn set_range_values_writes_matrix() {
        let mut engine = Engine::new();
        let range = Range::from_a1("A1:B2").expect("range");
        let values = vec![
            vec![Value::Number(1.0), Value::Number(2.0)],
            vec![Value::Number(3.0), Value::Text("x".to_string())],
        ];

        engine
            .set_range_values("Sheet1", range, &values, false)
            .unwrap();

        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
        assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
        assert_eq!(
            engine.get_cell_value("Sheet1", "B2"),
            Value::Text("x".to_string())
        );
    }

    #[test]
    fn set_range_values_blank_preserves_style_only_cells() {
        let mut engine = Engine::new();
        let style_id = engine.intern_style(Style {
            number_format: Some("0".to_string()),
            ..Style::default()
        });
        engine.set_cell_value("Sheet1", "A1", 123.0).unwrap();
        engine.set_cell_style_id("Sheet1", "A1", style_id).unwrap();

        let range = Range::from_a1("A1:A1").expect("range");
        let values = vec![vec![Value::Blank]];
        engine
            .set_range_values("Sheet1", range, &values, false)
            .unwrap();

        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Blank);
        assert_eq!(
            engine.get_cell_style_id("Sheet1", "A1").unwrap(),
            Some(style_id)
        );
    }

    #[test]
    fn set_range_values_clears_stale_cell_metadata_even_when_value_is_unchanged() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        let key = CellKey { sheet: sheet_id, addr };

        // Simulate an inconsistent state (e.g. stale bytecode flags left behind after a refactor).
        {
            let cell = engine.workbook.get_or_create_cell_mut(key);
            cell.bytecode_compile_reason = Some(BytecodeCompileReason::Disabled);
            cell.volatile = true;
            cell.thread_safe = false;
            cell.dynamic_deps = true;
        }

        let range = Range::from_a1("A1:A1").expect("range");
        let values = vec![vec![Value::Number(1.0)]];
        engine
            .set_range_values("Sheet1", range, &values, false)
            .unwrap();

        let cell = engine.workbook.get_cell(key).expect("cell exists");
        assert_eq!(cell.value, Value::Number(1.0));
        assert!(cell.bytecode_compile_reason.is_none());
        assert!(!cell.volatile);
        assert!(cell.thread_safe);
        assert!(!cell.dynamic_deps);
    }

    #[test]
    fn set_range_values_blank_prunes_default_blank_cells_from_sparse_storage() {
        let mut engine = Engine::new();
        let sheet_id = engine.workbook.ensure_sheet("Sheet1");
        let addr = parse_a1("A1").unwrap();
        let key = CellKey { sheet: sheet_id, addr };

        // Insert a "default blank" cell that should not remain in sparse storage.
        {
            let cell = engine.workbook.get_or_create_cell_mut(key);
            cell.value = Value::Blank;
            cell.formula = None;
            cell.compiled = None;
            cell.bytecode_compile_reason = None;
            cell.volatile = false;
            cell.thread_safe = true;
            cell.dynamic_deps = false;
            cell.phonetic = None;
            cell.style_id = 0;
            cell.number_format = None;
        }
        assert!(
            engine.workbook.sheets[sheet_id].cells.contains_key(&addr),
            "expected setup to create a sparse cell entry"
        );

        let range = Range::from_a1("A1:A1").expect("range");
        let values = vec![vec![Value::Blank]];
        engine
            .set_range_values("Sheet1", range, &values, false)
            .unwrap();

        assert!(
            engine.workbook.sheets[sheet_id].cells.get(&addr).is_none(),
            "expected bulk blank write to prune default blank cell entries"
        );
    }

    #[test]
    fn set_range_values_rounds_numbers_in_precision_as_displayed_mode() {
        let mut engine = Engine::new();
        engine.set_calc_settings(CalcSettings {
            calculation_mode: CalculationMode::Manual,
            full_precision: false,
            ..CalcSettings::default()
        });

        engine
            .set_cell_number_format("Sheet1", "A1", Some("0.00".to_string()))
            .unwrap();

        let range = Range::from_a1("A1:A1").expect("range");
        let values = vec![vec![Value::Number(1.234)]];
        engine
            .set_range_values("Sheet1", range, &values, false)
            .unwrap();

        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.23));
    }

    #[test]
    fn clear_range_removes_cells_from_sparse_storage() {
        let mut engine = Engine::new();
        let range = Range::from_a1("A1:B2").expect("range");
        let values = vec![
            vec![Value::Number(1.0), Value::Number(2.0)],
            vec![Value::Number(3.0), Value::Number(4.0)],
        ];

        engine
            .set_range_values("Sheet1", range, &values, false)
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        assert_eq!(engine.workbook.sheets[sheet_id].cells.len(), 4);

        engine.clear_range("Sheet1", range, false).unwrap();
        assert!(
            engine.workbook.sheets[sheet_id].cells.is_empty(),
            "cleared cells should be removed from sparse storage"
        );
    }

    #[test]
    fn set_range_values_recalculates_dependents_when_requested() {
        let mut engine = Engine::new();
        engine.set_calc_settings(CalcSettings {
            calculation_mode: CalculationMode::Automatic,
            ..CalcSettings::default()
        });

        engine.set_cell_formula("Sheet1", "C1", "=A1+B1").unwrap();

        let range = Range::from_a1("A1:B1").expect("range");
        let values = vec![vec![Value::Number(2.0), Value::Number(3.0)]];

        engine
            .set_range_values("Sheet1", range, &values, true)
            .unwrap();

        assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(5.0));
    }

    #[test]
    fn let_lambda_calls_are_thread_safe() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=LET(f,LAMBDA(x,x+1),f(2))")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        let cell = engine.workbook.sheets[sheet_id]
            .cells
            .get(&addr)
            .expect("cell stored");
        assert!(
            cell.thread_safe,
            "LET/LAMBDA should be safe for parallel evaluation"
        );
    }

    #[test]
    fn lambda_parameter_calls_are_thread_safe() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula(
                "Sheet1",
                "A1",
                "=LET(apply,LAMBDA(f,f(1)),apply(LAMBDA(x,x+1)))",
            )
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        let cell = engine.workbook.sheets[sheet_id]
            .cells
            .get(&addr)
            .expect("cell stored");
        assert!(
            cell.thread_safe,
            "higher-order lambda calls should be parallel-safe"
        );

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    }

    #[test]
    fn lambda_invocation_syntax_evaluates() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=LAMBDA(x,x+1)(3)")
            .unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(4.0));
    }

    #[test]
    fn nested_lambda_invocation_syntax_evaluates() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=LAMBDA(x,LAMBDA(y,x+y))(1)(2)")
            .unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(3.0));
    }

    #[test]
    fn lambda_invocation_tracks_dependencies() {
        let mut engine = Engine::new();
        engine
            .set_cell_value("Sheet1", "B1", Value::Number(10.0))
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "A1", "=LAMBDA(x,B1+x)(1)")
            .unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(11.0));

        engine
            .set_cell_value("Sheet1", "B1", Value::Number(20.0))
            .unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(21.0));
    }

    #[test]
    fn parenthesized_name_invocation_preserves_lambda_recursion() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula(
                "Sheet1",
                "A1",
                "=LET(f,LAMBDA(n,IF(n=0,1,n*f(n-1))),(f)(5))",
            )
            .unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(120.0));
    }

    #[test]
    fn info_recalc_reflects_calc_settings() {
        let mut engine = Engine::new();

        engine
            .set_cell_formula("Sheet1", "A1", r#"=INFO("recalc")"#)
            .unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(
            engine.get_cell_value("Sheet1", "A1"),
            Value::Text("Manual".to_string())
        );

        engine.set_calc_settings(CalcSettings {
            calculation_mode: CalculationMode::Automatic,
            ..CalcSettings::default()
        });
        engine
            .set_cell_formula("Sheet1", "A2", r#"=INFO("recalc")"#)
            .unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(
            engine.get_cell_value("Sheet1", "A2"),
            Value::Text("Automatic".to_string())
        );

        engine.set_calc_settings(CalcSettings {
            calculation_mode: CalculationMode::AutomaticNoTable,
            ..CalcSettings::default()
        });
        engine
            .set_cell_formula("Sheet1", "A3", r#"=INFO("recalc")"#)
            .unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(
            engine.get_cell_value("Sheet1", "A3"),
            Value::Text("Automatic except for tables".to_string())
        );
    }

    #[test]
    fn cell_width_reflects_column_metadata() {
        let mut engine = Engine::new();
        engine
            // Place the formula outside of the referenced column so this test stays focused on
            // column metadata changes; self-reference behavior is covered by
            // `cell_width_self_reference_is_not_circular`.
            .set_cell_formula("Sheet1", "B1", r#"=CELL("width", A1)"#)
            .unwrap();
        engine.recalculate_single_threaded();
        // Excel returns the column width rounded down to whole characters, with a `0.0` fractional
        // marker when the column uses the sheet default width.
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(8.0));

        // Column widths are stored in Excel "character" units (OOXML `col/@width`).
        engine.set_col_width("Sheet1", 0, Some(15.0));
        engine.recalculate_single_threaded();
        // When a column uses an explicit width override, Excel reports a `0.1` fractional marker.
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(15.1));

        // Hidden columns should return 0 for CELL("width").
        engine.set_col_hidden("Sheet1", 0, true);
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.0));
    }

    #[test]
    fn cell_contents_self_reference_is_not_circular() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", r#"=CELL("contents", A1)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        // Match the CELL("contents") implementation: ensure the serialized formula has a leading '='.
        let formula = engine
            .get_cell_formula("Sheet1", "A1")
            .expect("formula stored");
        let mut expected = formula.to_string();
        if !expected.trim_start().starts_with('=') {
            expected.insert(0, '=');
        }
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Text(expected));
    }

    #[test]
    fn cell_contents_multi_cell_range_including_formula_cell_is_not_circular() {
        let mut engine = Engine::new();
        engine
            .set_cell_value("Sheet1", "A1", Value::Number(123.0))
            .unwrap();
        // Range includes the formula cell (A2), but CELL("contents") consults only the top-left.
        engine
            .set_cell_formula("Sheet1", "A2", r#"=CELL("contents", A1:A3)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(123.0));
    }

    #[test]
    fn cell_contents_range_top_left_self_reference_is_not_circular() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", r#"=CELL("contents", A1:A3)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);

        // Match the CELL("contents") implementation: ensure the serialized formula has a leading '='.
        let formula = engine
            .get_cell_formula("Sheet1", "A1")
            .expect("formula stored");
        let mut expected = formula.to_string();
        if !expected.trim_start().starts_with('=') {
            expected.insert(0, '=');
        }
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Text(expected));
    }

    #[test]
    fn cell_type_self_reference_is_not_circular() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", r#"=CELL("type", A1)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        // CELL("type") reports "v" when the referenced cell contains a value or a formula.
        assert_eq!(
            engine.get_cell_value("Sheet1", "A1"),
            Value::Text("v".to_string())
        );
    }

    #[test]
    fn cell_type_multi_cell_range_including_formula_cell_is_not_circular() {
        let mut engine = Engine::new();
        // Range includes the formula cell (A2), but CELL("type") consults only the top-left.
        engine
            .set_cell_formula("Sheet1", "A2", r#"=CELL("type", A1:A3)"#)
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);
        // Top-left cell is blank => "b".
        assert_eq!(
            engine.get_cell_value("Sheet1", "A2"),
            Value::Text("b".to_string())
        );
    }

    #[test]
    fn formulatext_self_reference_is_not_circular() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=FORMULATEXT(A1)")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);

        let stored = engine
            .get_cell_formula("Sheet1", "A1")
            .expect("formula stored");
        let expected = crate::functions::information::workbook::normalize_formula_text(stored);
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Text(expected));
    }

    #[test]
    fn isformula_self_reference_is_not_circular() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=ISFORMULA(A1)")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(true));
    }

    #[test]
    fn formulatext_multi_cell_range_including_formula_cell_is_not_circular() {
        let mut engine = Engine::new();
        // This reference is invalid for FORMULATEXT (expects a single cell) but should still return
        // `#N/A` rather than being forced into circular-reference handling when the range includes
        // the formula cell.
        engine
            .set_cell_formula("Sheet1", "A2", "=FORMULATEXT(A1:A3)")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(
            engine.get_cell_value("Sheet1", "A2"),
            Value::Error(ErrorKind::NA)
        );
    }

    #[test]
    fn isformula_multi_cell_range_including_formula_cell_is_not_circular() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A2", "=ISFORMULA(A1:A3)")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(
            engine.get_cell_value("Sheet1", "A2"),
            Value::Error(ErrorKind::Value)
        );
    }

    #[test]
    fn isref_self_reference_is_not_circular() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=ISREF(A1)")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(true));
    }

    #[test]
    fn isref_multi_cell_range_including_formula_cell_is_not_circular() {
        let mut engine = Engine::new();
        // Range includes the formula cell (A2), but ISREF does not dereference values.
        engine
            .set_cell_formula("Sheet1", "A2", "=ISREF(A1:A3)")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(true));
    }

    #[test]
    fn isref_index_reference_does_not_create_range_node_cycles() {
        let mut engine = Engine::new();
        // The INDEX range includes the formula cell (A2), but INDEX is only used to compute a
        // reference here; ISREF should not introduce a range-node cycle.
        engine
            .set_cell_formula("Sheet1", "A2", "=ISREF(INDEX(A1:A3,1))")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(true));
    }

    #[test]
    fn row_choose_self_reference_is_not_circular() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=ROW(CHOOSE(1, A1, B1))")
            .unwrap();
        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    }

    #[test]
    fn reference_info_functions_do_not_create_spurious_cycles() {
        let mut engine = Engine::new();

        // Self-referential single-cell references should not become calc-graph cycles.
        engine.set_cell_formula("Sheet1", "D1", "=ROW(D1)").unwrap();
        engine
            .set_cell_formula("Sheet1", "D2", "=COLUMN(D2)")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "D3", "=ROWS(D1:D3)")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "D4", "=COLUMNS(B:D)")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "D5", "=SHEET(D5)")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "D6", "=SHEETS(D6)")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "D7", "=AREAS(D7)")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "D8", "=ROW(OFFSET(D8,0,0))")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "D9", "=COLUMN(OFFSET(D9,0,0))")
            .unwrap();

        // Range arguments that include the formula cell should not create range-node cycles.
        engine
            .set_cell_formula("Sheet1", "A1", "=SUM(ROW(1:5))")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "A2", "=SUM(COLUMN(A:C))")
            .unwrap();

        engine.recalculate_single_threaded();

        assert_eq!(engine.circular_reference_count(), 0);

        assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
        assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(4.0));
        assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(3.0));
        assert_eq!(engine.get_cell_value("Sheet1", "D4"), Value::Number(3.0));
        assert_eq!(engine.get_cell_value("Sheet1", "D5"), Value::Number(1.0));
        assert_eq!(engine.get_cell_value("Sheet1", "D6"), Value::Number(1.0));
        assert_eq!(engine.get_cell_value("Sheet1", "D7"), Value::Number(1.0));
        assert_eq!(engine.get_cell_value("Sheet1", "D8"), Value::Number(8.0));
        assert_eq!(engine.get_cell_value("Sheet1", "D9"), Value::Number(4.0));

        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(15.0));
        assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(6.0));
    }

    #[test]
    fn cell_width_self_reference_is_not_circular() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", r#"=CELL("width", A1)"#)
            .unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(8.0));

        engine.set_col_width("Sheet1", 0, Some(15.0));
        engine.recalculate_single_threaded();
        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(15.1));

        engine.set_col_hidden("Sheet1", 0, true);
        engine.recalculate_single_threaded();
        assert_eq!(engine.circular_reference_count(), 0);
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.0));
    }

    #[test]
    fn multithreaded_and_singlethreaded_match_for_volatiles_given_same_recalc_context() {
        fn setup(engine: &mut Engine) {
            engine
                .set_cell_formula("Sheet1", "A1", "=NOW()")
                .expect("set NOW()");
            engine
                .set_cell_formula("Sheet1", "A2", "=RAND()")
                .expect("set RAND()");
            engine
                .set_cell_formula("Sheet1", "A3", "=RANDBETWEEN(10, 20)")
                .expect("set RANDBETWEEN()");
            engine
                .set_cell_formula("Sheet1", "B1", "=A1+A2+A3")
                .expect("set dependent");
        }

        let mut single = Engine::new();
        setup(&mut single);
        let mut multi = Engine::new();
        setup(&mut multi);

        let recalc_ctx = crate::eval::RecalcContext {
            now_utc: chrono::Utc
                .timestamp_opt(1_700_000_000, 123_456_789)
                .single()
                .unwrap(),
            recalc_id: 42,
            number_locale: crate::value::NumberLocale::en_us(),
            calculation_mode: CalculationMode::Manual,
        };

        let levels_single = single
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = single.recalculate_levels(
            levels_single,
            RecalcMode::SingleThreaded,
            &recalc_ctx,
            single.date_system,
            None,
        );

        let levels_multi = multi
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = multi.recalculate_levels(
            levels_multi,
            RecalcMode::MultiThreaded,
            &recalc_ctx,
            multi.date_system,
            None,
        );

        for addr in ["A1", "A2", "A3", "B1"] {
            assert_eq!(
                multi.get_cell_value("Sheet1", addr),
                single.get_cell_value("Sheet1", addr),
                "mismatch at {addr}"
            );
        }
    }

    #[test]
    fn now_and_today_compile_to_bytecode() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=NOW()")
            .expect("set NOW()");
        engine
            .set_cell_formula("Sheet1", "A2", "=TODAY()")
            .expect("set TODAY()");

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        for addr in ["A1", "A2"] {
            let addr = parse_a1(addr).unwrap();
            let cell = engine.workbook.sheets[sheet_id]
                .cells
                .get(&addr)
                .expect("cell stored");
            assert!(
                matches!(cell.compiled.as_ref(), Some(CompiledFormula::Bytecode(_))),
                "{addr:?} should compile to bytecode"
            );
        }
    }

    #[test]
    fn now_and_today_bytecode_matches_ast_within_recalc_context() {
        fn setup(engine: &mut Engine) {
            engine
                .set_cell_formula("Sheet1", "A1", "=NOW()")
                .expect("set NOW()");
            engine
                .set_cell_formula("Sheet1", "A2", "=TODAY()")
                .expect("set TODAY()");
        }

        let mut bytecode = Engine::new();
        setup(&mut bytecode);

        let mut ast = Engine::new();
        ast.set_bytecode_enabled(false);
        setup(&mut ast);

        let recalc_ctx = crate::eval::RecalcContext {
            now_utc: chrono::Utc
                .timestamp_opt(1_700_000_000, 123_456_789)
                .single()
                .unwrap(),
            recalc_id: 42,
            number_locale: crate::value::NumberLocale::en_us(),
            calculation_mode: CalculationMode::Manual,
        };

        let levels_ast = ast.calc_graph.calc_levels_for_dirty().expect("calc levels");
        let _ = ast.recalculate_levels(
            levels_ast,
            RecalcMode::SingleThreaded,
            &recalc_ctx,
            ast.date_system,
            None,
        );

        let levels_bytecode = bytecode
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = bytecode.recalculate_levels(
            levels_bytecode,
            RecalcMode::SingleThreaded,
            &recalc_ctx,
            bytecode.date_system,
            None,
        );

        for addr in ["A1", "A2"] {
            assert_eq!(
                bytecode.get_cell_value("Sheet1", addr),
                ast.get_cell_value("Sheet1", addr),
                "mismatch at {addr}"
            );
        }
    }

    #[test]
    fn now_and_today_bytecode_respects_date_system() {
        use crate::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
        use chrono::{Datelike, Timelike};

        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=TODAY()")
            .expect("set TODAY()");
        engine
            .set_cell_formula("Sheet1", "A2", "=NOW()")
            .expect("set NOW()");

        let recalc_ctx = crate::eval::RecalcContext {
            now_utc: chrono::Utc
                .timestamp_opt(1_700_000_000, 123_456_789)
                .single()
                .unwrap(),
            recalc_id: 42,
            number_locale: crate::value::NumberLocale::en_us(),
            calculation_mode: CalculationMode::Manual,
        };

        let run = |engine: &mut Engine, ctx: &crate::eval::RecalcContext| {
            let levels = engine
                .calc_graph
                .calc_levels_for_dirty()
                .expect("calc levels");
            let _ = engine.recalculate_levels(
                levels,
                RecalcMode::SingleThreaded,
                ctx,
                engine.date_system,
                None,
            );
        };

        // Default: 1900 date system.
        run(&mut engine, &recalc_ctx);
        let now = recalc_ctx.now_utc;
        let date = now.date_naive();
        let seconds = now.time().num_seconds_from_midnight() as f64
            + (now.time().nanosecond() as f64 / 1_000_000_000.0);

        let base_1900 = ymd_to_serial(
            ExcelDate::new(date.year(), date.month() as u8, date.day() as u8),
            ExcelDateSystem::EXCEL_1900,
        )
        .unwrap() as f64;

        assert_eq!(
            engine.get_cell_value("Sheet1", "A1"),
            Value::Number(base_1900)
        );
        assert_eq!(
            engine.get_cell_value("Sheet1", "A2"),
            Value::Number(base_1900 + seconds / 86_400.0)
        );

        // Switch to 1904 date system and ensure NOW/TODAY shift appropriately.
        engine.set_date_system(ExcelDateSystem::Excel1904);
        run(&mut engine, &recalc_ctx);

        let base_1904 = ymd_to_serial(
            ExcelDate::new(date.year(), date.month() as u8, date.day() as u8),
            ExcelDateSystem::Excel1904,
        )
        .unwrap() as f64;

        assert_eq!(
            engine.get_cell_value("Sheet1", "A1"),
            Value::Number(base_1904)
        );
        assert_eq!(
            engine.get_cell_value("Sheet1", "A2"),
            Value::Number(base_1904 + seconds / 86_400.0)
        );
    }

    #[test]
    fn bytecode_rand_matches_ast_given_same_recalc_context() {
        fn setup(engine: &mut Engine) {
            engine
                .set_cell_formula("Sheet1", "A1", "=RAND()")
                .expect("set RAND()");
            engine
                .set_cell_formula("Sheet1", "A2", "=RAND()+RAND()")
                .expect("set RAND()+RAND()");
            engine
                .set_cell_formula("Sheet1", "A3", "=RANDBETWEEN(10, 20)")
                .expect("set RANDBETWEEN()");
            engine
                .set_cell_formula("Sheet1", "B1", "=A1+A2+A3")
                .expect("set dependent");
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);

        let recalc_ctx = crate::eval::RecalcContext {
            now_utc: chrono::Utc
                .timestamp_opt(1_700_000_000, 123_456_789)
                .single()
                .unwrap(),
            recalc_id: 123,
            number_locale: crate::value::NumberLocale::en_us(),
            calculation_mode: CalculationMode::Manual,
        };

        // Ensure the volatile RNG formulas compile to bytecode when the backend is enabled.
        let sheet_id = bytecode_engine
            .workbook
            .sheet_id("Sheet1")
            .expect("sheet exists");
        for addr in ["A1", "A2", "A3", "B1"] {
            let addr = parse_a1(addr).unwrap();
            let cell = bytecode_engine.workbook.sheets[sheet_id]
                .cells
                .get(&addr)
                .expect("cell stored");
            assert!(
                matches!(cell.compiled.as_ref(), Some(CompiledFormula::Bytecode(_))),
                "expected {addr:?} to compile to bytecode"
            );
        }

        let levels_bc = bytecode_engine
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = bytecode_engine.recalculate_levels(
            levels_bc,
            RecalcMode::SingleThreaded,
            &recalc_ctx,
            bytecode_engine.date_system,
            None,
        );

        let levels_ast = ast_engine
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = ast_engine.recalculate_levels(
            levels_ast,
            RecalcMode::SingleThreaded,
            &recalc_ctx,
            ast_engine.date_system,
            None,
        );

        for addr in ["A1", "A2", "A3", "B1"] {
            assert_eq!(
                bytecode_engine.get_cell_value("Sheet1", addr),
                ast_engine.get_cell_value("Sheet1", addr),
                "bytecode vs ast mismatch at {addr}"
            );
        }
    }

    #[test]
    fn recalculate_with_value_changes_includes_spill_outputs() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "A1", "=SEQUENCE(1,2)")
            .unwrap();

        let changes = engine.recalculate_with_value_changes(RecalcMode::SingleThreaded);
        assert_eq!(
            changes,
            vec![
                RecalcValueChange {
                    sheet: "Sheet1".to_string(),
                    addr: parse_a1("A1").unwrap(),
                    value: Value::Number(1.0),
                },
                RecalcValueChange {
                    sheet: "Sheet1".to_string(),
                    addr: parse_a1("B1").unwrap(),
                    value: Value::Number(2.0),
                },
            ]
        );
    }

    #[test]
    fn clear_cell_removes_literal_entry_from_sheet_map() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A1", 123.0).unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        assert!(
            engine.workbook.sheets[sheet_id].cells.contains_key(&addr),
            "literal cell should be stored"
        );

        engine.clear_cell("Sheet1", "A1").unwrap();

        assert!(
            !engine.workbook.sheets[sheet_id].cells.contains_key(&addr),
            "cleared literal cell should be removed from sparse storage"
        );
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Blank);
    }

    #[test]
    fn clear_cell_marks_dependents_dirty() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A1", 2.0).unwrap();
        engine.set_cell_formula("Sheet1", "B1", "=A1*2").unwrap();
        engine.recalculate();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(4.0));

        engine.clear_cell("Sheet1", "A1").unwrap();
        engine.recalculate();
        assert_eq!(
            engine.get_cell_value("Sheet1", "B1"),
            Value::Number(0.0),
            "clearing an input should propagate to dependent formulas"
        );
    }

    #[test]
    fn recalculate_with_value_changes_tracks_scalar_formula_updates() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
        engine.set_cell_formula("Sheet1", "A1", "=A2*2").unwrap();

        let changes = engine.recalculate_with_value_changes(RecalcMode::SingleThreaded);
        assert_eq!(
            changes,
            vec![RecalcValueChange {
                sheet: "Sheet1".to_string(),
                addr: parse_a1("A1").unwrap(),
                value: Value::Number(6.0),
            }]
        );

        engine.set_cell_value("Sheet1", "A2", 4.0).unwrap();
        let changes = engine.recalculate_with_value_changes(RecalcMode::SingleThreaded);
        assert_eq!(
            changes,
            vec![RecalcValueChange {
                sheet: "Sheet1".to_string(),
                addr: parse_a1("A1").unwrap(),
                value: Value::Number(8.0),
            }]
        );
    }

    #[test]
    fn bytecode_sparse_iteration_matches_ast_for_huge_sparse_ranges() {
        fn setup(engine: &mut Engine) {
            // Sparse values spread across a full Excel column.
            engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
            engine.set_cell_value("Sheet1", "A500000", 2.0).unwrap();
            engine.set_cell_value("Sheet1", "A1048576", 3.0).unwrap();

            // Sum/average value range (aligned with A:A).
            engine.set_cell_value("Sheet1", "C1", 10.0).unwrap();
            engine.set_cell_value("Sheet1", "C500000", 20.0).unwrap();
            engine.set_cell_value("Sheet1", "C1048576", 30.0).unwrap();

            // Secondary criteria range (aligned with A:A).
            engine.set_cell_value("Sheet1", "D1", 100.0).unwrap();
            engine.set_cell_value("Sheet1", "D500000", 200.0).unwrap();
            engine.set_cell_value("Sheet1", "D1048576", 100.0).unwrap();

            // Boolean values spread across a full Excel column.
            engine.set_cell_value("Sheet1", "E1", true).unwrap();
            engine.set_cell_value("Sheet1", "E500000", false).unwrap();
            engine.set_cell_value("Sheet1", "E1048576", true).unwrap();

            engine
                .set_cell_formula("Sheet1", "B1", "=SUM(A:A)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B2", "=COUNTIF(A:A, 0)")
                .unwrap();

            // Criteria aggregates over full-column ranges should also take the sparse iteration path.
            engine
                .set_cell_formula("Sheet1", "B3", r#"=SUMIF(A:A,">1",C:C)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B4", r#"=SUMIFS(C:C,A:A,">1",D:D,100)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B5", r#"=COUNTIFS(A:A,">1",D:D,100)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B6", r#"=AVERAGEIF(A:A,">1",C:C)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B7", r#"=AVERAGEIFS(C:C,A:A,">1",D:D,200)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B8", r#"=MINIFS(C:C,A:A,">1",D:D,200)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B9", r#"=MAXIFS(C:C,A:A,">1",D:D,200)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B10", r#"=COUNTIFS(A:A,"",D:D,"")"#)
                .unwrap();

            // Non-criteria aggregates over full-column ranges should also take the sparse iteration path.
            engine
                .set_cell_formula("Sheet1", "B11", "=COUNT(A:A)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B12", "=AVERAGE(A:A)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B13", "=MIN(A:A)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B14", "=MAX(A:A)")
                .unwrap();

            // Logical aggregations over full-column ranges should also take the sparse iteration path.
            engine
                .set_cell_formula("Sheet1", "B15", "=AND(E:E)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B16", "=OR(E:E)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B17", "=XOR(E:E)")
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        // Ensure the full-column formulas are actually bytecode-compiled.
        let sheet_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();
        let formula_cells = [
            "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13",
            "B14", "B15", "B16", "B17",
        ];
        let mut tasks: Vec<(CellKey, CompiledFormula)> = Vec::with_capacity(formula_cells.len());
        for cell in formula_cells {
            let addr = parse_a1(cell).unwrap();
            let compiled = bytecode_engine.workbook.sheets[sheet_id]
                .cells
                .get(&addr)
                .and_then(|c| c.compiled.as_ref())
                .cloned()
                .expect("compiled formula");
            assert!(
                matches!(compiled, CompiledFormula::Bytecode(_)),
                "expected {cell} to compile to bytecode"
            );
            tasks.push((
                CellKey {
                    sheet: sheet_id,
                    addr,
                },
                compiled,
            ));
        }

        // Column caches should *not* allocate a full-column buffer for `A:A` (or any other full-column
        // references used by the formulas above).
        let snapshot = Snapshot::from_workbook(
            &bytecode_engine.workbook,
            &bytecode_engine.spills,
            bytecode_engine.external_value_provider.clone(),
            bytecode_engine.external_data_provider.clone(),
            bytecode_engine.info.clone(),
            bytecode_engine.pivot_registry.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(bytecode_engine.workbook.sheets.len(), &snapshot, &tasks);
        assert!(
            column_cache.by_sheet[sheet_id].is_empty(),
            "expected full-column ranges to skip column-slice cache allocation"
        );

        bytecode_engine.recalculate_single_threaded();
        let bc_sum = bytecode_engine.get_cell_value("Sheet1", "B1");
        let bc_countif = bytecode_engine.get_cell_value("Sheet1", "B2");
        let bc_sumif = bytecode_engine.get_cell_value("Sheet1", "B3");
        let bc_sumifs = bytecode_engine.get_cell_value("Sheet1", "B4");
        let bc_countifs = bytecode_engine.get_cell_value("Sheet1", "B5");
        let bc_averageif = bytecode_engine.get_cell_value("Sheet1", "B6");
        let bc_averageifs = bytecode_engine.get_cell_value("Sheet1", "B7");
        let bc_minifs = bytecode_engine.get_cell_value("Sheet1", "B8");
        let bc_maxifs = bytecode_engine.get_cell_value("Sheet1", "B9");
        let bc_countifs_blank = bytecode_engine.get_cell_value("Sheet1", "B10");
        let bc_count = bytecode_engine.get_cell_value("Sheet1", "B11");
        let bc_average = bytecode_engine.get_cell_value("Sheet1", "B12");
        let bc_min = bytecode_engine.get_cell_value("Sheet1", "B13");
        let bc_max = bytecode_engine.get_cell_value("Sheet1", "B14");
        let bc_and = bytecode_engine.get_cell_value("Sheet1", "B15");
        let bc_or = bytecode_engine.get_cell_value("Sheet1", "B16");
        let bc_xor = bytecode_engine.get_cell_value("Sheet1", "B17");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        let ast_sum = ast_engine.get_cell_value("Sheet1", "B1");
        let ast_countif = ast_engine.get_cell_value("Sheet1", "B2");
        let ast_sumif = ast_engine.get_cell_value("Sheet1", "B3");
        let ast_sumifs = ast_engine.get_cell_value("Sheet1", "B4");
        let ast_countifs = ast_engine.get_cell_value("Sheet1", "B5");
        let ast_averageif = ast_engine.get_cell_value("Sheet1", "B6");
        let ast_averageifs = ast_engine.get_cell_value("Sheet1", "B7");
        let ast_minifs = ast_engine.get_cell_value("Sheet1", "B8");
        let ast_maxifs = ast_engine.get_cell_value("Sheet1", "B9");
        let ast_countifs_blank = ast_engine.get_cell_value("Sheet1", "B10");
        let ast_count = ast_engine.get_cell_value("Sheet1", "B11");
        let ast_average = ast_engine.get_cell_value("Sheet1", "B12");
        let ast_min = ast_engine.get_cell_value("Sheet1", "B13");
        let ast_max = ast_engine.get_cell_value("Sheet1", "B14");
        let ast_and = ast_engine.get_cell_value("Sheet1", "B15");
        let ast_or = ast_engine.get_cell_value("Sheet1", "B16");
        let ast_xor = ast_engine.get_cell_value("Sheet1", "B17");

        assert_eq!(bc_sum, ast_sum, "SUM mismatch");
        assert_eq!(bc_countif, ast_countif, "COUNTIF mismatch");
        assert_eq!(bc_sumif, ast_sumif, "SUMIF mismatch");
        assert_eq!(bc_sumifs, ast_sumifs, "SUMIFS mismatch");
        assert_eq!(bc_countifs, ast_countifs, "COUNTIFS mismatch");
        assert_eq!(bc_averageif, ast_averageif, "AVERAGEIF mismatch");
        assert_eq!(bc_averageifs, ast_averageifs, "AVERAGEIFS mismatch");
        assert_eq!(bc_minifs, ast_minifs, "MINIFS mismatch");
        assert_eq!(bc_maxifs, ast_maxifs, "MAXIFS mismatch");
        assert_eq!(
            bc_countifs_blank, ast_countifs_blank,
            "COUNTIFS blank mismatch"
        );
        assert_eq!(bc_count, ast_count, "COUNT mismatch");
        assert_eq!(bc_average, ast_average, "AVERAGE mismatch");
        assert_eq!(bc_min, ast_min, "MIN mismatch");
        assert_eq!(bc_max, ast_max, "MAX mismatch");
        assert_eq!(bc_and, ast_and, "AND mismatch");
        assert_eq!(bc_or, ast_or, "OR mismatch");
        assert_eq!(bc_xor, ast_xor, "XOR mismatch");

        // Sanity check expected values.
        assert_eq!(bc_sum, Value::Number(6.0));
        assert_eq!(bc_countif, Value::Number(1_048_573.0));
        assert_eq!(bc_sumif, Value::Number(50.0));
        assert_eq!(bc_sumifs, Value::Number(30.0));
        assert_eq!(bc_countifs, Value::Number(1.0));
        assert_eq!(bc_averageif, Value::Number(25.0));
        assert_eq!(bc_averageifs, Value::Number(20.0));
        assert_eq!(bc_minifs, Value::Number(20.0));
        assert_eq!(bc_maxifs, Value::Number(20.0));
        assert_eq!(bc_countifs_blank, Value::Number(1_048_573.0));
        assert_eq!(bc_count, Value::Number(3.0));
        assert_eq!(bc_average, Value::Number(2.0));
        assert_eq!(bc_min, Value::Number(1.0));
        assert_eq!(bc_max, Value::Number(3.0));
        assert_eq!(bc_and, Value::Bool(false));
        assert_eq!(bc_or, Value::Bool(true));
        assert_eq!(bc_xor, Value::Bool(false));
    }

    #[test]
    fn bytecode_sparse_iteration_matches_ast_for_huge_sparse_ranges_counta_countblank() {
        fn setup(engine: &mut Engine) {
            engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
            engine.set_cell_value("Sheet1", "A2", "").unwrap(); // empty string
            engine.set_cell_value("Sheet1", "A500000", 2.0).unwrap();
            engine.set_cell_value("Sheet1", "A1048576", 3.0).unwrap();

            engine
                .set_cell_formula("Sheet1", "B1", "=COUNTA(A:A)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B2", "=COUNTBLANK(A:A)")
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        // Ensure the COUNTA formula is actually bytecode-compiled.
        let sheet_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();
        let b1 = parse_a1("B1").unwrap();
        let cell_b1 = bytecode_engine.workbook.sheets[sheet_id]
            .cells
            .get(&b1)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula");
        assert!(
            matches!(cell_b1, CompiledFormula::Bytecode(_)),
            "expected COUNTA(A:A) to compile to bytecode"
        );

        // Column caches should *not* allocate a full-column buffer for `A:A`.
        let snapshot = Snapshot::from_workbook(
            &bytecode_engine.workbook,
            &bytecode_engine.spills,
            bytecode_engine.external_value_provider.clone(),
            bytecode_engine.external_data_provider.clone(),
            bytecode_engine.info.clone(),
            bytecode_engine.pivot_registry.clone(),
        );
        let key_b1 = CellKey {
            sheet: sheet_id,
            addr: b1,
        };
        let tasks = vec![(key_b1, cell_b1.clone())];
        let column_cache =
            BytecodeColumnCache::build(bytecode_engine.workbook.sheets.len(), &snapshot, &tasks);
        assert!(
            !column_cache
                .by_sheet
                .get(sheet_id)
                .map(|cols| cols.contains_key(&0))
                .unwrap_or(false),
            "expected full-column range to skip column-slice cache allocation"
        );

        bytecode_engine.recalculate_single_threaded();
        let bc_counta = bytecode_engine.get_cell_value("Sheet1", "B1");
        let bc_countblank = bytecode_engine.get_cell_value("Sheet1", "B2");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        let ast_counta = ast_engine.get_cell_value("Sheet1", "B1");
        let ast_countblank = ast_engine.get_cell_value("Sheet1", "B2");

        assert_eq!(bc_counta, ast_counta, "COUNTA mismatch");
        assert_eq!(bc_countblank, ast_countblank, "COUNTBLANK mismatch");

        // Sanity check expected values.
        assert_eq!(bc_counta, Value::Number(4.0));
        assert_eq!(bc_countblank, Value::Number(1_048_573.0));
    }

    #[test]
    fn bytecode_sparse_sumproduct_matches_ast_for_huge_sparse_ranges() {
        fn setup(engine: &mut Engine) {
            engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
            engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
            engine.set_cell_value("Sheet1", "A500000", 2.0).unwrap();
            engine.set_cell_value("Sheet1", "B500000", 20.0).unwrap();
            engine.set_cell_value("Sheet1", "A1048576", 3.0).unwrap();
            engine.set_cell_value("Sheet1", "B1048576", 30.0).unwrap();

            engine
                .set_cell_formula("Sheet1", "C1", "=SUMPRODUCT(A:A,B:B)")
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        // Ensure SUMPRODUCT compiled to bytecode.
        let sheet_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();
        let c1 = parse_a1("C1").unwrap();
        let cell_c1 = bytecode_engine.workbook.sheets[sheet_id]
            .cells
            .get(&c1)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula");
        assert!(
            matches!(cell_c1, CompiledFormula::Bytecode(_)),
            "expected SUMPRODUCT(A:A,B:B) to compile to bytecode"
        );

        // Full-column ranges should skip building column slices.
        let snapshot = Snapshot::from_workbook(
            &bytecode_engine.workbook,
            &bytecode_engine.spills,
            bytecode_engine.external_value_provider.clone(),
            bytecode_engine.external_data_provider.clone(),
            bytecode_engine.info.clone(),
            bytecode_engine.pivot_registry.clone(),
        );
        let key_c1 = CellKey {
            sheet: sheet_id,
            addr: c1,
        };
        let tasks = vec![(key_c1, cell_c1.clone())];
        let column_cache =
            BytecodeColumnCache::build(bytecode_engine.workbook.sheets.len(), &snapshot, &tasks);
        assert!(
            !column_cache
                .by_sheet
                .get(sheet_id)
                .map(|cols| cols.contains_key(&0) || cols.contains_key(&1))
                .unwrap_or(false),
            "expected full-column ranges to skip column-slice cache allocation"
        );

        bytecode_engine.recalculate_single_threaded();
        let bc_sumproduct = bytecode_engine.get_cell_value("Sheet1", "C1");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        let ast_sumproduct = ast_engine.get_cell_value("Sheet1", "C1");

        assert_eq!(bc_sumproduct, ast_sumproduct, "SUMPRODUCT mismatch");
        assert_eq!(bc_sumproduct, Value::Number(140.0));
    }

    #[test]
    fn bytecode_sparse_iteration_matches_ast_for_huge_sparse_3d_ranges_counta_countblank() {
        fn setup(engine: &mut Engine) {
            for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
                engine.set_cell_value(sheet, "A1", 1.0).unwrap();
                engine.set_cell_value(sheet, "A2", "").unwrap(); // empty string
                engine.set_cell_value(sheet, "A500000", 2.0).unwrap();
                engine.set_cell_value(sheet, "A1048576", 3.0).unwrap();
            }

            engine
                .set_cell_formula("Sheet1", "B1", "=COUNTA(Sheet1:Sheet3!A:A)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B2", "=COUNTBLANK(Sheet1:Sheet3!A:A)")
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        // Ensure the COUNTA formula is actually bytecode-compiled.
        let sheet1_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();
        let b1 = parse_a1("B1").unwrap();
        let b2 = parse_a1("B2").unwrap();
        let cell_b1 = bytecode_engine.workbook.sheets[sheet1_id]
            .cells
            .get(&b1)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula");
        let cell_b2 = bytecode_engine.workbook.sheets[sheet1_id]
            .cells
            .get(&b2)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula");
        assert!(
            matches!(cell_b1, CompiledFormula::Bytecode(_)),
            "expected COUNTA(Sheet1:Sheet3!A:A) to compile to bytecode"
        );
        assert!(
            matches!(cell_b2, CompiledFormula::Bytecode(_)),
            "expected COUNTBLANK(Sheet1:Sheet3!A:A) to compile to bytecode"
        );

        // Column caches should *not* allocate full-column buffers for 3D spans over `A:A`.
        let snapshot = Snapshot::from_workbook(
            &bytecode_engine.workbook,
            &bytecode_engine.spills,
            bytecode_engine.external_value_provider.clone(),
            bytecode_engine.external_data_provider.clone(),
            bytecode_engine.info.clone(),
            bytecode_engine.pivot_registry.clone(),
        );
        let key_b1 = CellKey {
            sheet: sheet1_id,
            addr: b1,
        };
        let key_b2 = CellKey {
            sheet: sheet1_id,
            addr: b2,
        };
        let tasks = vec![(key_b1, cell_b1.clone()), (key_b2, cell_b2.clone())];
        let column_cache =
            BytecodeColumnCache::build(bytecode_engine.workbook.sheets.len(), &snapshot, &tasks);

        for sheet_name in ["Sheet1", "Sheet2", "Sheet3"] {
            let sheet_id = bytecode_engine.workbook.sheet_id(sheet_name).unwrap();
            assert!(
                !column_cache
                    .by_sheet
                    .get(sheet_id)
                    .map(|cols| cols.contains_key(&0))
                    .unwrap_or(false),
                "expected full-column 3D span to skip column-slice cache allocation for {sheet_name}"
            );
        }

        bytecode_engine.recalculate_single_threaded();
        let bc_counta = bytecode_engine.get_cell_value("Sheet1", "B1");
        let bc_countblank = bytecode_engine.get_cell_value("Sheet1", "B2");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        let ast_counta = ast_engine.get_cell_value("Sheet1", "B1");
        let ast_countblank = ast_engine.get_cell_value("Sheet1", "B2");

        assert_eq!(bc_counta, ast_counta, "COUNTA mismatch");
        assert_eq!(bc_countblank, ast_countblank, "COUNTBLANK mismatch");

        // Sanity check expected values.
        assert_eq!(bc_counta, Value::Number(12.0));
        assert_eq!(bc_countblank, Value::Number(3_145_719.0));
    }

    #[test]
    fn bytecode_sparse_iteration_matches_ast_for_huge_sparse_3d_ranges() {
        fn setup(engine: &mut Engine) {
            for (sheet, values) in [
                ("Sheet1", [1.0, 2.0, 3.0]),
                ("Sheet2", [4.0, 5.0, 6.0]),
                ("Sheet3", [7.0, 8.0, 9.0]),
            ] {
                engine.set_cell_value(sheet, "A1", values[0]).unwrap();
                engine.set_cell_value(sheet, "A500000", values[1]).unwrap();
                engine.set_cell_value(sheet, "A1048576", values[2]).unwrap();
            }

            // Boolean values spread across a full Excel column on each sheet.
            //
            // Keep the overall XOR result non-trivial by using an odd number of TRUE values across
            // the full 3D span.
            for (sheet, end_true) in [("Sheet1", true), ("Sheet2", false), ("Sheet3", true)] {
                engine.set_cell_value(sheet, "E1", true).unwrap();
                engine.set_cell_value(sheet, "E500000", false).unwrap();
                engine.set_cell_value(sheet, "E1048576", end_true).unwrap();
            }

            engine
                .set_cell_formula("Sheet1", "B1", "=SUM(Sheet1:Sheet3!A:A)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B2", "=COUNTIF(Sheet1:Sheet3!A:A, 0)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B3", "=MIN(Sheet1:Sheet3!A:A)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B4", "=MAX(Sheet1:Sheet3!A:A)")
                .unwrap();

            engine
                .set_cell_formula("Sheet1", "B5", "=COUNT(Sheet1:Sheet3!A:A)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B6", "=AVERAGE(Sheet1:Sheet3!A:A)")
                .unwrap();

            engine
                .set_cell_formula("Sheet1", "B7", "=AND(Sheet1:Sheet3!E:E)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B8", "=OR(Sheet1:Sheet3!E:E)")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B9", "=XOR(Sheet1:Sheet3!E:E)")
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        // Ensure the full-column formulas are actually bytecode-compiled.
        let sheet1_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();
        let formula_cells = ["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9"];
        let mut tasks: Vec<(CellKey, CompiledFormula)> = Vec::with_capacity(formula_cells.len());
        for cell in formula_cells {
            let addr = parse_a1(cell).unwrap();
            let compiled = bytecode_engine.workbook.sheets[sheet1_id]
                .cells
                .get(&addr)
                .and_then(|c| c.compiled.as_ref())
                .cloned()
                .expect("compiled formula");
            assert!(
                matches!(compiled, CompiledFormula::Bytecode(_)),
                "expected {cell} to compile to bytecode"
            );
            tasks.push((
                CellKey {
                    sheet: sheet1_id,
                    addr,
                },
                compiled,
            ));
        }

        // Column caches should *not* allocate full-column buffers for 3D spans over `A:A` / `E:E`.
        let snapshot = Snapshot::from_workbook(
            &bytecode_engine.workbook,
            &bytecode_engine.spills,
            bytecode_engine.external_value_provider.clone(),
            bytecode_engine.external_data_provider.clone(),
            bytecode_engine.info.clone(),
            bytecode_engine.pivot_registry.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(bytecode_engine.workbook.sheets.len(), &snapshot, &tasks);

        for sheet_name in ["Sheet1", "Sheet2", "Sheet3"] {
            let sheet_id = bytecode_engine.workbook.sheet_id(sheet_name).unwrap();
            assert!(
                column_cache.by_sheet[sheet_id].is_empty(),
                "expected full-column 3D span to skip column-slice cache allocation for {sheet_name}",
            );
        }

        bytecode_engine.recalculate_single_threaded();
        let bc_sum = bytecode_engine.get_cell_value("Sheet1", "B1");
        let bc_countif = bytecode_engine.get_cell_value("Sheet1", "B2");
        let bc_min = bytecode_engine.get_cell_value("Sheet1", "B3");
        let bc_max = bytecode_engine.get_cell_value("Sheet1", "B4");
        let bc_count = bytecode_engine.get_cell_value("Sheet1", "B5");
        let bc_average = bytecode_engine.get_cell_value("Sheet1", "B6");
        let bc_and = bytecode_engine.get_cell_value("Sheet1", "B7");
        let bc_or = bytecode_engine.get_cell_value("Sheet1", "B8");
        let bc_xor = bytecode_engine.get_cell_value("Sheet1", "B9");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        let ast_sum = ast_engine.get_cell_value("Sheet1", "B1");
        let ast_countif = ast_engine.get_cell_value("Sheet1", "B2");
        let ast_min = ast_engine.get_cell_value("Sheet1", "B3");
        let ast_max = ast_engine.get_cell_value("Sheet1", "B4");
        let ast_count = ast_engine.get_cell_value("Sheet1", "B5");
        let ast_average = ast_engine.get_cell_value("Sheet1", "B6");
        let ast_and = ast_engine.get_cell_value("Sheet1", "B7");
        let ast_or = ast_engine.get_cell_value("Sheet1", "B8");
        let ast_xor = ast_engine.get_cell_value("Sheet1", "B9");

        assert_eq!(bc_sum, ast_sum, "SUM mismatch");
        assert_eq!(bc_countif, ast_countif, "COUNTIF mismatch");
        assert_eq!(bc_min, ast_min, "MIN mismatch");
        assert_eq!(bc_max, ast_max, "MAX mismatch");
        assert_eq!(bc_count, ast_count, "COUNT mismatch");
        assert_eq!(bc_average, ast_average, "AVERAGE mismatch");
        assert_eq!(bc_and, ast_and, "AND mismatch");
        assert_eq!(bc_or, ast_or, "OR mismatch");
        assert_eq!(bc_xor, ast_xor, "XOR mismatch");

        // Sanity check expected values.
        assert_eq!(bc_sum, Value::Number(45.0));
        assert_eq!(bc_countif, Value::Number(3_145_719.0));
        assert_eq!(bc_min, Value::Number(1.0));
        assert_eq!(bc_max, Value::Number(9.0));
        assert_eq!(bc_count, Value::Number(9.0));
        assert_eq!(bc_average, Value::Number(5.0));
        assert_eq!(bc_and, Value::Bool(false));
        assert_eq!(bc_or, Value::Bool(true));
        assert_eq!(bc_xor, Value::Bool(true));
    }

    #[test]
    fn bytecode_sparse_iteration_matches_ast_for_huge_sparse_criteria_aggregates() {
        fn setup(engine: &mut Engine) {
            // Sparse values spread across a full Excel column.
            engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
            engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
            engine.set_cell_value("Sheet1", "A500000", 0.0).unwrap();
            engine.set_cell_value("Sheet1", "A1048576", 3.0).unwrap();

            engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
            engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
            // Row 3 has a number in the sum range but a blank criteria cell (implicit blank).
            engine.set_cell_value("Sheet1", "B3", 7.0).unwrap();
            engine.set_cell_value("Sheet1", "B500000", 5.0).unwrap();
            engine.set_cell_value("Sheet1", "B1048576", 30.0).unwrap();

            engine
                .set_cell_formula("Sheet1", "C1", r#"=SUMIF(A:A,">1",B:B)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C2", r#"=AVERAGEIF(A:A,">0",B:B)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C3", r#"=SUMIFS(B:B,A:A,">0")"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C4", r#"=AVERAGEIFS(B:B,A:A,">0")"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C5", r#"=COUNTIFS(A:A,">0")"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C6", r#"=COUNTIFS(A:A,0)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C7", r#"=MINIFS(B:B,A:A,">0")"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C8", r#"=MAXIFS(B:B,A:A,">0")"#)
                .unwrap();

            // Multiple-criteria variants.
            engine
                .set_cell_formula("Sheet1", "C9", r#"=SUMIFS(B:B,A:A,">0",B:B,">15")"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C10", r#"=AVERAGEIFS(B:B,A:A,">0",B:B,">15")"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C11", r#"=COUNTIFS(A:A,">0",B:B,">15")"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C12", r#"=MINIFS(B:B,A:A,">0",B:B,">15")"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C13", r#"=MAXIFS(B:B,A:A,">0",B:B,">15")"#)
                .unwrap();

            // Blank-criteria cases: criteria range is mostly implicit blanks.
            engine
                .set_cell_formula("Sheet1", "C14", r#"=SUMIF(A:A,"",B:B)"#)
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "C15", r#"=COUNTIFS(A:A,"")"#)
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        let sheet_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();

        // Ensure formulas are bytecode-compiled.
        let addrs = [
            "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13",
            "C14", "C15",
        ];

        let mut tasks: Vec<(CellKey, CompiledFormula)> = Vec::new();
        for a1 in addrs {
            let addr = parse_a1(a1).unwrap();
            let compiled = bytecode_engine.workbook.sheets[sheet_id]
                .cells
                .get(&addr)
                .and_then(|c| c.compiled.clone())
                .expect("compiled formula");
            assert!(
                matches!(compiled, CompiledFormula::Bytecode(_)),
                "expected {a1} to compile to bytecode"
            );
            tasks.push((
                CellKey {
                    sheet: sheet_id,
                    addr,
                },
                compiled,
            ));
        }

        // Column caches should *not* allocate full-column buffers for `A:A` / `B:B`.
        let snapshot = Snapshot::from_workbook(
            &bytecode_engine.workbook,
            &bytecode_engine.spills,
            bytecode_engine.external_value_provider.clone(),
            bytecode_engine.external_data_provider.clone(),
            bytecode_engine.info.clone(),
            bytecode_engine.pivot_registry.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(bytecode_engine.workbook.sheets.len(), &snapshot, &tasks);
        assert!(
            !column_cache
                .by_sheet
                .get(sheet_id)
                .map(|cols| cols.contains_key(&0) || cols.contains_key(&1))
                .unwrap_or(false),
            "expected full-column criteria aggregates to skip column-slice cache allocation"
        );

        bytecode_engine.recalculate_single_threaded();

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();

        let expected = [
            ("C1", Value::Number(50.0)),
            ("C2", Value::Number(20.0)),
            ("C3", Value::Number(60.0)),
            ("C4", Value::Number(20.0)),
            ("C5", Value::Number(3.0)),
            ("C6", Value::Number(1_048_573.0)),
            ("C7", Value::Number(10.0)),
            ("C8", Value::Number(30.0)),
            ("C9", Value::Number(50.0)),
            ("C10", Value::Number(25.0)),
            ("C11", Value::Number(2.0)),
            ("C12", Value::Number(20.0)),
            ("C13", Value::Number(30.0)),
            ("C14", Value::Number(7.0)),
            ("C15", Value::Number(1_048_572.0)),
        ];

        for (addr, value) in expected {
            let bc = bytecode_engine.get_cell_value("Sheet1", addr);
            let ast = ast_engine.get_cell_value("Sheet1", addr);
            assert_eq!(bc, ast, "mismatch at {addr}");
            assert_eq!(bc, value, "unexpected value at {addr}");
        }
    }

    #[test]
    fn recalculate_with_value_changes_clears_shrunk_spill_cells() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "A1", "=SEQUENCE(1,A2)")
            .unwrap();

        let _ = engine.recalculate_with_value_changes(RecalcMode::SingleThreaded);

        // Shrink the spill width from 2 columns to 1. The previous spill cell (B1)
        // should be returned as a delta back to blank.
        engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
        let changes = engine.recalculate_with_value_changes(RecalcMode::SingleThreaded);
        assert_eq!(
            changes,
            vec![RecalcValueChange {
                sheet: "Sheet1".to_string(),
                addr: parse_a1("B1").unwrap(),
                value: Value::Blank,
            }]
        );
    }

    #[test]
    fn clear_cell_removes_formula_entry_from_sheet_map() {
        let mut engine = Engine::new();
        engine.set_cell_formula("Sheet1", "A1", "=1+1").unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr = parse_a1("A1").unwrap();
        assert!(
            engine.workbook.sheets[sheet_id].cells.contains_key(&addr),
            "formula cell should be stored"
        );
        assert_eq!(engine.get_cell_formula("Sheet1", "A1"), Some("=1+1"));

        engine.clear_cell("Sheet1", "A1").unwrap();

        assert!(
            !engine.workbook.sheets[sheet_id].cells.contains_key(&addr),
            "cleared formula cell should be removed from sparse storage"
        );
        assert_eq!(engine.get_cell_formula("Sheet1", "A1"), None);
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Blank);
    }

    #[test]
    fn bytecode_compile_report_allows_cross_sheet_references() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet2", "A1", 1.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "A1", "=Sheet2!A1+1")
            .unwrap();

        let report = engine.bytecode_compile_report(10);
        assert_eq!(report.len(), 0, "expected formula to compile to bytecode");

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    }

    #[test]
    fn bytecode_compile_report_allows_range_expressions() {
        let mut engine = Engine::new();
        engine.set_cell_formula("Sheet1", "A1", "=A2:A3").unwrap();

        let report = engine.bytecode_compile_report(10);
        assert_eq!(report.len(), 0);
    }

    #[test]
    fn bytecode_compile_report_orders_by_tab_order_after_sheet_reorder() {
        let mut engine = Engine::new();

        // Use a non-thread-safe function so formulas deterministically fall back from bytecode.
        for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
            engine
                .set_cell_formula(sheet, "A1", "=RTD(\"prog\",\"server\",\"topic\")")
                .unwrap();
        }

        let report = engine.bytecode_compile_report(10);
        assert_eq!(report.len(), 3);
        assert_eq!(
            report.iter().map(|e| e.sheet.as_str()).collect::<Vec<_>>(),
            vec!["Sheet1", "Sheet2", "Sheet3"],
            "expected report order to match the default tab order"
        );

        let sheet1_id = engine.workbook.sheet_id("Sheet1").unwrap();
        let sheet2_id = engine.workbook.sheet_id("Sheet2").unwrap();
        let sheet3_id = engine.workbook.sheet_id("Sheet3").unwrap();
        engine
            .workbook
            .set_sheet_order(vec![sheet2_id, sheet3_id, sheet1_id]);

        let reordered = engine.bytecode_compile_report(10);
        assert_eq!(reordered.len(), 3);
        assert_eq!(
            reordered
                .iter()
                .map(|e| e.sheet.as_str())
                .collect::<Vec<_>>(),
            vec!["Sheet2", "Sheet3", "Sheet1"],
            "expected report order to match the updated tab order"
        );
    }

    #[test]
    fn workbook_sheet_ids_are_stable_and_sheet_spans_follow_tab_order() {
        let mut workbook = Workbook::default();
        let sheet1 = workbook.ensure_sheet("Sheet1");
        let sheet2 = workbook.ensure_sheet("Sheet2");
        let sheet3 = workbook.ensure_sheet("Sheet3");

        // Sheet names are case-insensitive.
        assert_eq!(workbook.ensure_sheet("sheet1"), sheet1);
        assert_eq!(workbook.sheet_id("SHEET2"), Some(sheet2));

        // Default tab order matches creation order.
        assert_eq!(workbook.sheet_ids_in_order(), &[sheet1, sheet2, sheet3]);
        assert_eq!(workbook.sheet_order_index(sheet1), Some(0));
        assert_eq!(workbook.sheet_name(sheet3), Some("Sheet3"));

        // Reorder sheets without changing ids.
        workbook.set_sheet_order(vec![sheet2, sheet3, sheet1]);
        assert_eq!(workbook.sheet_ids_in_order(), &[sheet2, sheet3, sheet1]);
        assert_eq!(workbook.sheet_id("Sheet1"), Some(sheet1));
        assert_eq!(workbook.sheet_id("Sheet2"), Some(sheet2));
        assert_eq!(workbook.sheet_id("Sheet3"), Some(sheet3));

        // Excel-style 3D spans use tab order (and support reversed spans).
        assert_eq!(
            workbook.sheet_span_ids(sheet2, sheet1),
            Some(vec![sheet2, sheet3, sheet1])
        );
        assert_eq!(
            workbook.sheet_span_ids(sheet1, sheet2),
            Some(vec![sheet2, sheet3, sheet1])
        );
    }

    #[test]
    fn workbook_sheet_ids_are_not_reused_after_sheet_delete() {
        let mut engine = Engine::new();
        let sheet1 = engine.workbook.ensure_sheet("Sheet1");
        let sheet2 = engine.workbook.ensure_sheet("Sheet2");
        let sheet3 = engine.workbook.ensure_sheet("Sheet3");

        engine.delete_sheet("Sheet2").unwrap();

        assert!(engine.workbook.sheet_exists(sheet1));
        assert!(!engine.workbook.sheet_exists(sheet2));
        assert!(engine.workbook.sheet_exists(sheet3));
        assert_eq!(engine.workbook.sheet_id("Sheet2"), None);
        assert_eq!(engine.workbook.sheet_name(sheet2), None);
        assert_eq!(engine.workbook.sheet_ids_in_order(), &[sheet1, sheet3]);

        let sheet4 = engine.workbook.ensure_sheet("Sheet4");
        assert_eq!(sheet4, 3, "deleted sheet ids should not be reused");
        assert_eq!(
            engine.workbook.sheet_ids_in_order(),
            &[sheet1, sheet3, sheet4]
        );

        // 3D spans are driven by the current tab order and ignore deleted sheets.
        assert_eq!(
            engine.workbook.sheet_span_ids(sheet1, sheet3),
            Some(vec![sheet1, sheet3])
        );
        assert_eq!(engine.workbook.sheet_span_ids(sheet2, sheet3), None);
    }

    #[test]
    fn delete_sheet_prunes_pivots_even_when_the_pivot_uses_the_stable_sheet_key() {
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");
        engine.ensure_sheet("Sheet2");

        // Give the sheet a display name distinct from its stable key so we can ensure pivot cleanup
        // matches both forms.
        engine.set_sheet_display_name("Sheet2", "Report");

        // Attach a table to the sheet so we can also validate table-backed pivot pruning.
        let sheet2_id = engine.workbook.sheet_id("Sheet2").unwrap();
        engine.workbook.set_tables(
            sheet2_id,
            vec![Table {
                id: 42,
                name: "Table1".to_string(),
                display_name: "Table1".to_string(),
                range: Range::from_a1("A1:B2").unwrap(),
                header_row_count: 1,
                totals_row_count: 0,
                columns: vec![],
                style: None,
                auto_filter: None,
                relationship_id: None,
                part_path: None,
            }],
        );

        let pivot_dest_key = engine.add_pivot_table(PivotTableDefinition {
            id: 0,
            name: "PivotDestKey".to_string(),
            source: PivotSource::Range {
                sheet: "Sheet1".to_string(),
                range: None,
            },
            destination: crate::pivot::PivotDestination {
                // Reference the *stable key* for Sheet2.
                sheet: "Sheet2".to_string(),
                cell: CellRef::new(0, 0),
            },
            config: crate::pivot::PivotConfig::default(),
            apply_number_formats: true,
            last_output_range: None,
            needs_refresh: false,
        });

        let pivot_source_table = engine.add_pivot_table(PivotTableDefinition {
            id: 0,
            name: "PivotSourceTable".to_string(),
            source: PivotSource::Table { table_id: 42 },
            destination: crate::pivot::PivotDestination {
                sheet: "Sheet1".to_string(),
                cell: CellRef::new(0, 0),
            },
            config: crate::pivot::PivotConfig::default(),
            apply_number_formats: true,
            last_output_range: None,
            needs_refresh: false,
        });

        assert!(engine.pivot_table(pivot_dest_key).is_some());
        assert!(engine.pivot_table(pivot_source_table).is_some());

        // Delete using the *display name*.
        engine.delete_sheet("Report").unwrap();

        assert!(
            engine.pivot_table(pivot_dest_key).is_none(),
            "expected pivot definitions referencing the deleted sheet key to be pruned"
        );
        assert!(
            engine.pivot_table(pivot_source_table).is_none(),
            "expected table-backed pivots referencing tables from the deleted sheet to be pruned"
        );
    }

    #[test]
    fn delete_sheet_prunes_pivots_even_when_the_pivot_uses_the_sheet_display_name() {
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");
        engine.ensure_sheet("Sheet2");

        engine.set_sheet_display_name("Sheet2", "Report");

        let pivot_dest_display = engine.add_pivot_table(PivotTableDefinition {
            id: 0,
            name: "PivotDestDisplay".to_string(),
            source: PivotSource::Range {
                sheet: "Sheet1".to_string(),
                range: None,
            },
            destination: crate::pivot::PivotDestination {
                // Reference the *display name* for Sheet2.
                sheet: "Report".to_string(),
                cell: CellRef::new(0, 0),
            },
            config: crate::pivot::PivotConfig::default(),
            apply_number_formats: true,
            last_output_range: None,
            needs_refresh: false,
        });

        assert!(engine.pivot_table(pivot_dest_display).is_some());

        // Delete using the *stable key*.
        engine.delete_sheet("Sheet2").unwrap();

        assert!(
            engine.pivot_table(pivot_dest_display).is_none(),
            "expected pivot definitions referencing the deleted sheet display name to be pruned"
        );
    }

    #[test]
    fn index_area_num_over_sheet_span_uses_tab_order_after_reorder() {
        fn setup(engine: &mut Engine) {
            for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
                engine.ensure_sheet(sheet);
            }
            engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
            engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
            engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

            // Reorder tabs to [Sheet3, Sheet2, Sheet1].
            let sheet1_id = engine.workbook.sheet_id("Sheet1").unwrap();
            let sheet2_id = engine.workbook.sheet_id("Sheet2").unwrap();
            let sheet3_id = engine.workbook.sheet_id("Sheet3").unwrap();
            engine
                .workbook
                .set_sheet_order(vec![sheet3_id, sheet2_id, sheet1_id]);

            engine
                .set_cell_formula("Sheet1", "B1", "=SUM(INDEX(Sheet1:Sheet3!A1,1,1,1))")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B2", "=SUM(INDEX(Sheet1:Sheet3!A1,1,1,2))")
                .unwrap();
            engine
                .set_cell_formula("Sheet1", "B3", "=SUM(INDEX(Sheet1:Sheet3!A1,1,1,3))")
                .unwrap();
        }

        let mut engine = Engine::new();
        setup(&mut engine);
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
        assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
        assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(1.0));

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        assert_eq!(
            ast_engine.get_cell_value("Sheet1", "B1"),
            Value::Number(3.0)
        );
        assert_eq!(
            ast_engine.get_cell_value("Sheet1", "B2"),
            Value::Number(2.0)
        );
        assert_eq!(
            ast_engine.get_cell_value("Sheet1", "B3"),
            Value::Number(1.0)
        );
    }

    #[test]
    fn bytecode_dynamic_deref_matches_ast_for_sheet_spans() {
        fn setup(engine: &mut Engine) {
            for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
                engine.ensure_sheet(sheet);
            }
            engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
            engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
            engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
            engine
                .set_cell_formula("Sheet1", "B1", "=Sheet1:Sheet3!A1")
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        // Ensure the formula is actually bytecode-compiled.
        let sheet_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();
        let b1 = parse_a1("B1").unwrap();
        let cell_b1 = bytecode_engine.workbook.sheets[sheet_id]
            .cells
            .get(&b1)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula");
        assert!(
            matches!(cell_b1, CompiledFormula::Bytecode(_)),
            "expected Sheet1:Sheet3!A1 to compile to bytecode"
        );

        bytecode_engine.recalculate_single_threaded();
        let bc_value = bytecode_engine.get_cell_value("Sheet1", "B1");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        let ast_value = ast_engine.get_cell_value("Sheet1", "B1");

        assert_eq!(bc_value, ast_value, "bytecode/AST mismatch");
        // Discontiguous unions (e.g. 3D sheet spans) cannot be spilled as a single array.
        assert_eq!(bc_value, Value::Error(ErrorKind::Value));
    }

    #[test]
    fn bytecode_implicit_intersection_over_sheet_spans_matches_ast() {
        fn setup(engine: &mut Engine) {
            for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
                engine.ensure_sheet(sheet);
            }
            engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
            engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
            engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
            engine
                .set_cell_formula("Sheet1", "B1", "=ISNUMBER(@Sheet1:Sheet3!A1)")
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        // Ensure the formula is actually bytecode-compiled.
        let sheet_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();
        let b1 = parse_a1("B1").unwrap();
        let cell_b1 = bytecode_engine.workbook.sheets[sheet_id]
            .cells
            .get(&b1)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula");
        assert!(
            matches!(cell_b1, CompiledFormula::Bytecode(_)),
            "expected ISNUMBER(@Sheet1:Sheet3!A1) to compile to bytecode"
        );

        bytecode_engine.recalculate_single_threaded();
        let bc_value = bytecode_engine.get_cell_value("Sheet1", "B1");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        let ast_value = ast_engine.get_cell_value("Sheet1", "B1");

        assert_eq!(bc_value, ast_value, "bytecode/AST mismatch");
        // Implicit intersection over a 3D span is ambiguous, yielding #VALUE!, and information
        // functions should treat that as a non-number rather than propagating the error.
        assert_eq!(bc_value, Value::Bool(false));
    }

    #[test]
    fn reorder_sheet_rejects_unknown_sheets_and_out_of_range_indices() {
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");
        engine.ensure_sheet("Sheet2");

        // Unknown sheet name.
        assert!(!engine.reorder_sheet("Missing", 0));
        // Out-of-range tab index (len == 2).
        assert!(!engine.reorder_sheet("Sheet1", 2));

        // No-op reorder should succeed.
        assert!(engine.reorder_sheet("Sheet1", 0));
        // Valid move should succeed.
        assert!(engine.reorder_sheet("Sheet2", 0));
    }

    #[test]
    fn reorder_sheet_updates_tab_order_for_forward_and_backward_moves() {
        let mut engine = Engine::new();
        for sheet in ["Sheet1", "Sheet2", "Sheet3", "Sheet4"] {
            engine.ensure_sheet(sheet);
        }

        assert_eq!(
            engine.sheet_names_in_order(),
            vec![
                "Sheet1".to_string(),
                "Sheet2".to_string(),
                "Sheet3".to_string(),
                "Sheet4".to_string()
            ]
        );

        // Move forward (lower -> higher index).
        assert!(engine.reorder_sheet("Sheet1", 2));
        assert_eq!(
            engine.sheet_names_in_order(),
            vec![
                "Sheet2".to_string(),
                "Sheet3".to_string(),
                "Sheet1".to_string(),
                "Sheet4".to_string()
            ]
        );

        // Move backward (higher -> lower index).
        assert!(engine.reorder_sheet("Sheet4", 1));
        assert_eq!(
            engine.sheet_names_in_order(),
            vec![
                "Sheet2".to_string(),
                "Sheet4".to_string(),
                "Sheet3".to_string(),
                "Sheet1".to_string()
            ]
        );

        // No-op reorder should succeed and keep order unchanged.
        assert!(engine.reorder_sheet("Sheet4", 1));
        assert_eq!(
            engine.sheet_names_in_order(),
            vec![
                "Sheet2".to_string(),
                "Sheet4".to_string(),
                "Sheet3".to_string(),
                "Sheet1".to_string()
            ]
        );

        // Out-of-range index should fail and keep order unchanged.
        assert!(!engine.reorder_sheet("Sheet4", 4));
        assert_eq!(
            engine.sheet_names_in_order(),
            vec![
                "Sheet2".to_string(),
                "Sheet4".to_string(),
                "Sheet3".to_string(),
                "Sheet1".to_string()
            ]
        );
    }

    #[test]
    fn workbook_reorder_sheet_semantics_match_engine_contract() {
        // Exercise `Workbook::reorder_sheet` directly so we cover its contract independently of
        // the engine wrapper (which may rebuild graphs, etc).
        let mut workbook = Workbook::default();
        let sheet1 = workbook.ensure_sheet("Sheet1");
        let sheet2 = workbook.ensure_sheet("Sheet2");
        let sheet3 = workbook.ensure_sheet("Sheet3");

        assert_eq!(workbook.sheet_ids_in_order(), &[sheet1, sheet2, sheet3]);

        // Unknown sheet id should fail.
        assert!(!workbook.reorder_sheet(usize::MAX, 0));
        assert_eq!(workbook.sheet_ids_in_order(), &[sheet1, sheet2, sheet3]);

        // Out-of-range index should fail.
        assert!(!workbook.reorder_sheet(sheet1, 3));
        assert_eq!(workbook.sheet_ids_in_order(), &[sheet1, sheet2, sheet3]);

        // No-op reorder should succeed and keep order unchanged.
        assert!(workbook.reorder_sheet(sheet1, 0));
        assert_eq!(workbook.sheet_ids_in_order(), &[sheet1, sheet2, sheet3]);

        // Move forward (lower -> higher index).
        assert!(workbook.reorder_sheet(sheet1, 2));
        assert_eq!(workbook.sheet_ids_in_order(), &[sheet2, sheet3, sheet1]);

        // Move backward (higher -> lower index).
        assert!(workbook.reorder_sheet(sheet1, 0));
        assert_eq!(workbook.sheet_ids_in_order(), &[sheet1, sheet2, sheet3]);

        // Simulate an inconsistent state where a live sheet is missing from `sheet_order`; this
        // should fail without mutating the remaining order.
        workbook.sheet_order.retain(|&id| id != sheet1);
        let before = workbook.sheet_order.clone();
        assert!(!workbook.reorder_sheet(sheet1, 0));
        assert_eq!(workbook.sheet_order, before);
    }

    #[test]
    fn bytecode_concat_over_sheet_span_uses_tab_order_after_reorder() {
        fn setup(engine: &mut Engine) {
            for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
                engine.ensure_sheet(sheet);
            }
            engine.set_cell_value("Sheet1", "A1", "1").unwrap();
            engine.set_cell_value("Sheet2", "A1", "2").unwrap();
            engine.set_cell_value("Sheet3", "A1", "3").unwrap();

            // Reorder tabs to [Sheet3, Sheet2, Sheet1].
            let sheet1_id = engine.workbook.sheet_id("Sheet1").unwrap();
            let sheet2_id = engine.workbook.sheet_id("Sheet2").unwrap();
            let sheet3_id = engine.workbook.sheet_id("Sheet3").unwrap();
            engine
                .workbook
                .set_sheet_order(vec![sheet3_id, sheet2_id, sheet1_id]);

            engine
                .set_cell_formula("Sheet1", "B1", "=CONCAT(Sheet1:Sheet3!A1)")
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        // Ensure the formula is actually bytecode-compiled.
        let sheet_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();
        let b1 = parse_a1("B1").unwrap();
        let compiled_b1 = bytecode_engine.workbook.sheets[sheet_id]
            .cells
            .get(&b1)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula");
        assert!(
            matches!(compiled_b1, CompiledFormula::Bytecode(_)),
            "expected CONCAT(Sheet1:Sheet3!A1) to compile to bytecode"
        );

        bytecode_engine.recalculate_single_threaded();
        let bc_value = bytecode_engine.get_cell_value("Sheet1", "B1");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        let ast_value = ast_engine.get_cell_value("Sheet1", "B1");

        assert_eq!(bc_value, ast_value, "bytecode/AST mismatch");
        assert_eq!(bc_value, Value::from("321"));
    }

    #[test]
    fn bytecode_sum_over_sheet_span_error_precedence_follows_tab_order() {
        fn setup(engine: &mut Engine) {
            for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
                engine.ensure_sheet(sheet);
            }

            engine
                .set_cell_value("Sheet1", "A1", Value::Error(ErrorKind::Div0))
                .unwrap();
            engine
                .set_cell_value("Sheet2", "A1", Value::Error(ErrorKind::Value))
                .unwrap();
            engine
                .set_cell_value("Sheet3", "A1", Value::Error(ErrorKind::Name))
                .unwrap();

            // Reorder tabs to [Sheet3, Sheet2, Sheet1].
            let sheet1_id = engine.workbook.sheet_id("Sheet1").unwrap();
            let sheet2_id = engine.workbook.sheet_id("Sheet2").unwrap();
            let sheet3_id = engine.workbook.sheet_id("Sheet3").unwrap();
            engine
                .workbook
                .set_sheet_order(vec![sheet3_id, sheet2_id, sheet1_id]);

            engine
                .set_cell_formula("Sheet1", "B1", "=SUM(Sheet1:Sheet3!A1)")
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        // Ensure the formula is actually bytecode-compiled.
        let sheet_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();
        let b1 = parse_a1("B1").unwrap();
        let compiled_b1 = bytecode_engine.workbook.sheets[sheet_id]
            .cells
            .get(&b1)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula");
        assert!(
            matches!(compiled_b1, CompiledFormula::Bytecode(_)),
            "expected SUM(Sheet1:Sheet3!A1) to compile to bytecode"
        );

        bytecode_engine.recalculate_single_threaded();
        let bc_value = bytecode_engine.get_cell_value("Sheet1", "B1");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        let ast_value = ast_engine.get_cell_value("Sheet1", "B1");

        assert_eq!(bc_value, ast_value, "bytecode/AST mismatch");
        assert_eq!(bc_value, Value::Error(ErrorKind::Name));
    }

    #[test]
    fn bytecode_and_over_sheet_span_error_precedence_follows_tab_order() {
        fn setup(engine: &mut Engine) {
            for sheet in ["Sheet1", "Sheet2", "Sheet3"] {
                engine.ensure_sheet(sheet);
            }

            engine
                .set_cell_value("Sheet1", "A1", Value::Error(ErrorKind::Div0))
                .unwrap();
            engine
                .set_cell_value("Sheet2", "A1", Value::Error(ErrorKind::Value))
                .unwrap();
            engine
                .set_cell_value("Sheet3", "A1", Value::Error(ErrorKind::Name))
                .unwrap();

            // Reorder tabs to [Sheet3, Sheet2, Sheet1].
            let sheet1_id = engine.workbook.sheet_id("Sheet1").unwrap();
            let sheet2_id = engine.workbook.sheet_id("Sheet2").unwrap();
            let sheet3_id = engine.workbook.sheet_id("Sheet3").unwrap();
            engine
                .workbook
                .set_sheet_order(vec![sheet3_id, sheet2_id, sheet1_id]);

            engine
                .set_cell_formula("Sheet1", "B1", "=AND(Sheet1:Sheet3!A1)")
                .unwrap();
        }

        let mut bytecode_engine = Engine::new();
        setup(&mut bytecode_engine);

        // Ensure the formula is actually bytecode-compiled.
        let sheet_id = bytecode_engine.workbook.sheet_id("Sheet1").unwrap();
        let b1 = parse_a1("B1").unwrap();
        let compiled_b1 = bytecode_engine.workbook.sheets[sheet_id]
            .cells
            .get(&b1)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula");
        assert!(
            matches!(compiled_b1, CompiledFormula::Bytecode(_)),
            "expected AND(Sheet1:Sheet3!A1) to compile to bytecode"
        );

        bytecode_engine.recalculate_single_threaded();
        let bc_value = bytecode_engine.get_cell_value("Sheet1", "B1");

        let mut ast_engine = Engine::new();
        ast_engine.set_bytecode_enabled(false);
        setup(&mut ast_engine);
        ast_engine.recalculate_single_threaded();
        let ast_value = ast_engine.get_cell_value("Sheet1", "B1");

        assert_eq!(bc_value, ast_value, "bytecode/AST mismatch");
        assert_eq!(bc_value, Value::Error(ErrorKind::Name));
    }

    #[test]
    fn bytecode_sheet_span_expansion_respects_tab_order_after_reorder_and_rebuild() {
        let mut engine = Engine::new();
        for sheet in ["Sheet1", "Sheet2", "Sheet3", "Sheet4"] {
            engine.ensure_sheet(sheet);
        }

        engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
        engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
        engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
        engine.set_cell_value("Sheet4", "A1", 4.0).unwrap();

        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(Sheet1:Sheet3!A1)")
            .unwrap();

        // Ensure the formula is actually bytecode-compiled.
        let sheet1_id = engine.workbook.sheet_id("Sheet1").unwrap();
        let b1 = parse_a1("B1").unwrap();
        let compiled_b1 = engine.workbook.sheets[sheet1_id]
            .cells
            .get(&b1)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula");
        assert!(
            matches!(compiled_b1, CompiledFormula::Bytecode(_)),
            "expected SUM(Sheet1:Sheet3!A1) to compile to bytecode"
        );

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));

        // Reorder the tab order so Sheet4 falls within the Sheet1:Sheet3 span.
        // This also rebuilds the dependency graph so bytecode-expanded spans refresh.
        assert!(engine.reorder_sheet("Sheet4", 1));

        let compiled_b1_after = engine.workbook.sheets[sheet1_id]
            .cells
            .get(&b1)
            .and_then(|c| c.compiled.as_ref())
            .expect("compiled formula after rebuild");
        assert!(
            matches!(compiled_b1_after, CompiledFormula::Bytecode(_)),
            "expected SUM(Sheet1:Sheet3!A1) to remain bytecode-compiled after rebuild"
        );

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));
    }

    #[test]
    fn ast_sheet_span_expansion_respects_tab_order_after_reorder_and_rebuild() {
        let mut engine = Engine::new();
        engine.set_bytecode_enabled(false);
        for sheet in ["Sheet1", "Sheet2", "Sheet3", "Sheet4"] {
            engine.ensure_sheet(sheet);
        }

        engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
        engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
        engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();
        engine.set_cell_value("Sheet4", "A1", 4.0).unwrap();

        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(Sheet1:Sheet3!A1)")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "B2", "=SUM(Sheet3:Sheet1!A1)")
            .unwrap();

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
        assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(6.0));

        // Reorder the tab order so Sheet4 falls within the Sheet1:Sheet3 span.
        let sheet1_id = engine.workbook.sheet_id("Sheet1").unwrap();
        let sheet2_id = engine.workbook.sheet_id("Sheet2").unwrap();
        let sheet3_id = engine.workbook.sheet_id("Sheet3").unwrap();
        let sheet4_id = engine.workbook.sheet_id("Sheet4").unwrap();
        engine
            .workbook
            .set_sheet_order(vec![sheet1_id, sheet4_id, sheet2_id, sheet3_id]);

        engine.rebuild_graph().unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));
        assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(10.0));
    }

    #[test]
    fn bytecode_compile_report_classifies_unsupported_expressions() {
        let mut engine = Engine::new();
        engine
            // Sheet-qualified defined names are currently handled by the AST evaluator.
            .set_cell_formula("Sheet1", "B1", "=Sheet1!Foo")
            .unwrap();

        let report = engine.bytecode_compile_report(10);
        assert_eq!(report.len(), 1);
        // `LowerError::Unsupported` is mapped to `IneligibleExpr` so compile reports can distinguish
        // "missing bytecode implementation" from structural lowering errors (e.g. cross-sheet refs).
        assert_eq!(report[0].reason, BytecodeCompileReason::IneligibleExpr);
    }

    #[test]
    fn bytecode_compile_report_allows_spill_ranges() {
        let mut engine = Engine::new();
        engine.set_cell_formula("Sheet1", "B1", "=A1#").unwrap();

        let report = engine.bytecode_compile_report(10);
        assert_eq!(report.len(), 0);
    }

    #[test]
    fn bytecode_compile_report_classifies_not_thread_safe_formulas() {
        let mut engine = Engine::new();
        // RTD is volatile + not thread-safe (requires an external data provider).
        engine
            .set_cell_formula("Sheet1", "A1", "=RTD(\"prog\",\"server\",\"topic\")")
            .unwrap();

        let report = engine.bytecode_compile_report(10);
        assert_eq!(report.len(), 1);
        assert_eq!(report[0].reason, BytecodeCompileReason::NotThreadSafe);
    }

    #[test]
    fn bytecode_compile_report_allows_non_default_sheet_dimensions() {
        let mut engine = Engine::new();
        engine
            .set_sheet_dimensions("Sheet1", 100, EXCEL_MAX_COLS)
            .unwrap();
        engine.set_cell_formula("Sheet1", "A1", "=1+1").unwrap();

        let report = engine.bytecode_compile_report(10);
        assert!(
            report.is_empty(),
            "expected formula to compile to bytecode on non-default sheet dimensions; report: {report:?}"
        );
        assert_eq!(engine.bytecode_program_count(), 1);
    }

    #[test]
    fn bytecode_compile_report_allows_large_ranges_in_aggregate_contexts() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "AA1", "=SUM(A1:Z200000)")
            .unwrap();

        let report = engine.bytecode_compile_report(10);
        assert!(
            report.is_empty(),
            "expected large aggregate ranges to compile to bytecode; report: {report:?}"
        );
    }

    #[test]
    fn bytecode_column_cache_uses_range_min_row() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(A900000:A900010)")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let key = CellKey {
            sheet: sheet_id,
            addr: parse_a1("B1").unwrap(),
        };
        let compiled = engine
            .workbook
            .get_cell(key)
            .and_then(|c| c.compiled.clone())
            .expect("compiled formula stored");

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
            engine.external_data_provider.clone(),
            engine.info.clone(),
            engine.pivot_registry.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(engine.workbook.sheets.len(), &snapshot, &[(key, compiled)]);

        let col = column_cache.by_sheet[sheet_id]
            .get(&0)
            .expect("column A is cached");
        assert_eq!(col.segments.len(), 1);
        let seg = &col.segments[0];
        assert_eq!(seg.row_start, 899_999);
        assert_eq!(seg.values.len(), 11);
    }

    #[test]
    fn bytecode_column_cache_builds_disjoint_segments() {
        let mut engine = Engine::new();
        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(A1:A10)+SUM(A900000:A900010)")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let key = CellKey {
            sheet: sheet_id,
            addr: parse_a1("B1").unwrap(),
        };
        let compiled = engine
            .workbook
            .get_cell(key)
            .and_then(|c| c.compiled.clone())
            .expect("compiled formula stored");

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
            engine.external_data_provider.clone(),
            engine.info.clone(),
            engine.pivot_registry.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(engine.workbook.sheets.len(), &snapshot, &[(key, compiled)]);

        let col = column_cache.by_sheet[sheet_id]
            .get(&0)
            .expect("column A is cached");
        assert_eq!(col.segments.len(), 2);
        assert_eq!(col.segments[0].row_start, 0);
        assert_eq!(col.segments[0].values.len(), 10);
        assert_eq!(col.segments[1].row_start, 899_999);
        assert_eq!(col.segments[1].values.len(), 11);
    }

    #[test]
    fn bytecode_column_cache_populates_segment_values() {
        let mut engine = Engine::new();
        engine.set_cell_value("Sheet1", "A5", 42.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(A1:A10)")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let key = CellKey {
            sheet: sheet_id,
            addr: parse_a1("B1").unwrap(),
        };
        let compiled = engine
            .workbook
            .get_cell(key)
            .and_then(|c| c.compiled.clone())
            .expect("compiled formula stored");

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
            engine.external_data_provider.clone(),
            engine.info.clone(),
            engine.pivot_registry.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(engine.workbook.sheets.len(), &snapshot, &[(key, compiled)]);

        let col = column_cache.by_sheet[sheet_id]
            .get(&0)
            .expect("column A is cached");
        assert_eq!(col.segments.len(), 1);
        let seg = &col.segments[0];
        assert_eq!(seg.row_start, 0);
        assert_eq!(seg.values[4], 42.0);
    }

    #[test]
    fn bytecode_column_cache_treats_rich_values_like_text() {
        let mut engine = Engine::new();
        engine
            .set_cell_value(
                "Sheet1",
                "A1",
                Value::Entity(crate::value::EntityValue::new("Entity display")),
            )
            .unwrap();
        engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(A1:A2)")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let key = CellKey {
            sheet: sheet_id,
            addr: parse_a1("B1").unwrap(),
        };
        let compiled = engine
            .workbook
            .get_cell(key)
            .and_then(|c| c.compiled.clone())
            .expect("compiled formula stored");

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
            engine.external_data_provider.clone(),
            engine.info.clone(),
            engine.pivot_registry.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(engine.workbook.sheets.len(), &snapshot, &[(key, compiled)]);

        let col = column_cache.by_sheet[sheet_id]
            .get(&0)
            .expect("column A is cached");
        assert_eq!(col.segments.len(), 1);
        let seg = &col.segments[0];
        assert_eq!(seg.row_start, 0);

        assert!(
            seg.values[0].is_nan(),
            "rich value should not write a numeric slot"
        );
        assert_eq!(seg.values[1], 3.0);

        assert_eq!(seg.blocked_rows_strict, vec![0]);
        assert!(
            seg.blocked_rows_ignore_nonnumeric.is_empty(),
            "rich values should not block IgnoreNonNumeric slices"
        );
    }

    #[test]
    fn value_delta_treats_rich_values_as_non_numeric() {
        use crate::value::{EntityValue, RecordValue};

        let entity_a = Value::Entity(EntityValue::new("A"));
        let entity_b = Value::Entity(EntityValue::new("B"));
        assert!(numeric_value(&entity_a).is_none());
        assert_eq!(value_delta(&entity_a, &entity_b), f64::INFINITY);
        assert_eq!(value_delta(&entity_a, &entity_a), 0.0);

        let record_a = Value::Record(RecordValue::new("A"));
        let record_b = Value::Record(RecordValue::new("B"));
        assert!(numeric_value(&record_a).is_none());
        assert_eq!(value_delta(&record_a, &record_b), f64::INFINITY);
        assert_eq!(value_delta(&record_a, &record_a), 0.0);
    }

    #[test]
    fn bytecode_column_cache_ignores_ranges_used_only_for_implicit_intersection() {
        let mut engine = Engine::new();
        engine.set_cell_formula("Sheet1", "B1", "=@A1:A10").unwrap();

        assert_eq!(
            engine.bytecode_program_count(),
            1,
            "implicit intersection formulas should compile to bytecode"
        );

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let key = CellKey {
            sheet: sheet_id,
            addr: parse_a1("B1").unwrap(),
        };
        let compiled = engine
            .workbook
            .get_cell(key)
            .and_then(|c| c.compiled.clone())
            .expect("compiled formula stored");

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
            engine.external_data_provider.clone(),
            engine.info.clone(),
            engine.pivot_registry.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(engine.workbook.sheets.len(), &snapshot, &[(key, compiled)]);

        assert!(
            column_cache.by_sheet[sheet_id].is_empty(),
            "implicit intersection ranges should not force column cache allocation"
        );
    }

    #[test]
    fn bytecode_column_cache_ignores_out_of_bounds_ranges() {
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");
        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");

        let addr = parse_a1("A1").unwrap();
        let origin = bytecode::CellCoord {
            row: addr.row as i32,
            col: addr.col as i32,
        };
        let expr = bytecode::parse_formula("=SUM(XFE1:XFE2)", origin).expect("bytecode parse");
        let program = engine.bytecode_cache.get_or_compile(&expr);

        // The bytecode column cache ignores out-of-bounds ranges, but still needs a dummy AST
        // payload to satisfy the `CompiledFormula::Bytecode` wrapper.
        let parsed = crate::parse_formula("=1", crate::ParseOptions::default()).unwrap();
        let mut resolve_sheet = |name: &str| engine.workbook.sheet_id(name);
        let mut sheet_dims = |sheet_id: usize| {
            engine
                .workbook
                .sheets
                .get(sheet_id)
                .map(|s| (s.row_count, s.col_count))
                .unwrap_or((EXCEL_MAX_ROWS, EXCEL_MAX_COLS))
        };
        let ast = compile_canonical_expr(
            &parsed.expr,
            sheet_id,
            addr,
            &mut resolve_sheet,
            &mut sheet_dims,
        );

        let key = CellKey {
            sheet: sheet_id,
            addr,
        };
        let compiled = CompiledFormula::Bytecode(BytecodeFormula {
            ast,
            program,
            sheet_dims_generation: engine.sheet_dims_generation,
        });

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
            engine.external_data_provider.clone(),
            engine.info.clone(),
            engine.pivot_registry.clone(),
        );
        let column_cache =
            BytecodeColumnCache::build(engine.workbook.sheets.len(), &snapshot, &[(key, compiled)]);

        assert!(column_cache.by_sheet[sheet_id].is_empty());
    }

    #[test]
    fn bytecode_compiler_handles_huge_ranges_via_sparse_iteration() {
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet2");

        // Put some sparse values on Sheet2 so the full-sheet range has a non-trivial result.
        engine.set_cell_value("Sheet2", "A1", 1.0).unwrap();
        engine.set_cell_value("Sheet2", "C3", 2.0).unwrap();

        // A full-sheet range reference is extremely large (rows * cols), but bytecode evaluation
        // can still handle it efficiently via sparse iteration over the stored cells.
        //
        // Avoid self-references by summing a different sheet than the formula cell.
        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(Sheet2!A1:XFD1048576)")
            .unwrap();

        // Full-sheet ranges are enormous; the bytecode runtime can still evaluate aggregates over
        // these ranges via sparse iteration. The column cache must not allocate dense buffers for
        // them.
        assert_eq!(engine.bytecode_program_count(), 1);

        let sheet1_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let sheet2_id = engine.workbook.sheet_id("Sheet2").expect("sheet exists");
        let key_b1 = CellKey {
            sheet: sheet1_id,
            addr: parse_a1("B1").unwrap(),
        };
        let compiled = engine
            .workbook
            .get_cell(key_b1)
            .and_then(|c| c.compiled.clone())
            .expect("compiled formula stored");
        assert!(
            matches!(compiled, CompiledFormula::Bytecode(_)),
            "expected compiled formula to take the bytecode path"
        );

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
            engine.external_data_provider.clone(),
            engine.info.clone(),
            engine.pivot_registry.clone(),
        );
        let column_cache = BytecodeColumnCache::build(
            engine.workbook.sheets.len(),
            &snapshot,
            &[(key_b1, compiled)],
        );
        assert!(
            column_cache.by_sheet[sheet2_id].is_empty(),
            "expected full-sheet range to skip column-slice cache allocation"
        );

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    }

    fn assert_bytecode_matches_ast(formula: &str, expected: Value) {
        let addr = "A1";
        let recalc_ctx = crate::eval::RecalcContext {
            now_utc: chrono::Utc
                .timestamp_opt(1_700_000_000, 123_456_789)
                .single()
                .unwrap(),
            recalc_id: 42,
            number_locale: crate::value::NumberLocale::en_us(),
            calculation_mode: CalculationMode::Manual,
        };

        // Bytecode-enabled engine.
        let mut engine_bc = Engine::new();
        engine_bc.set_cell_formula("Sheet1", addr, formula).unwrap();
        assert_eq!(
            engine_bc.bytecode_program_count(),
            1,
            "expected formula to compile to bytecode"
        );
        let sheet_id = engine_bc.workbook.sheet_id("Sheet1").expect("sheet exists");
        let key = CellKey {
            sheet: sheet_id,
            addr: parse_a1(addr).unwrap(),
        };
        assert!(
            matches!(
                engine_bc
                    .workbook
                    .get_cell(key)
                    .and_then(|c| c.compiled.as_ref()),
                Some(CompiledFormula::Bytecode(_))
            ),
            "expected compiled formula to take the bytecode path"
        );
        let levels = engine_bc
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = engine_bc.recalculate_levels(
            levels,
            RecalcMode::SingleThreaded,
            &recalc_ctx,
            engine_bc.date_system,
            None,
        );
        let value_bc = engine_bc.get_cell_value("Sheet1", addr);
        assert_eq!(value_bc, expected);

        // Bytecode-disabled engine (AST-only).
        let mut engine_ast = Engine::new();
        engine_ast.set_bytecode_enabled(false);
        engine_ast
            .set_cell_formula("Sheet1", addr, formula)
            .unwrap();
        assert_eq!(
            engine_ast.bytecode_program_count(),
            0,
            "bytecode-disabled engine should not compile programs"
        );
        let levels = engine_ast
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = engine_ast.recalculate_levels(
            levels,
            RecalcMode::SingleThreaded,
            &recalc_ctx,
            engine_ast.date_system,
            None,
        );
        let value_ast = engine_ast.get_cell_value("Sheet1", addr);
        assert_eq!(value_ast, expected);

        assert_eq!(value_bc, value_ast);
    }

    fn assert_bytecode_eq_ast(formula: &str) -> Value {
        let addr = "A1";
        let recalc_ctx = crate::eval::RecalcContext {
            now_utc: chrono::Utc
                .timestamp_opt(1_700_000_000, 123_456_789)
                .single()
                .unwrap(),
            recalc_id: 42,
            number_locale: crate::value::NumberLocale::en_us(),
            calculation_mode: CalculationMode::Manual,
        };

        // Bytecode-enabled engine.
        let mut engine_bc = Engine::new();
        engine_bc.set_cell_formula("Sheet1", addr, formula).unwrap();
        assert_eq!(
            engine_bc.bytecode_program_count(),
            1,
            "expected formula to compile to bytecode"
        );
        let sheet_id = engine_bc.workbook.sheet_id("Sheet1").expect("sheet exists");
        let key = CellKey {
            sheet: sheet_id,
            addr: parse_a1(addr).unwrap(),
        };
        assert!(
            matches!(
                engine_bc
                    .workbook
                    .get_cell(key)
                    .and_then(|c| c.compiled.as_ref()),
                Some(CompiledFormula::Bytecode(_))
            ),
            "expected compiled formula to take the bytecode path"
        );
        let levels = engine_bc
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = engine_bc.recalculate_levels(
            levels,
            RecalcMode::SingleThreaded,
            &recalc_ctx,
            engine_bc.date_system,
            None,
        );
        let value_bc = engine_bc.get_cell_value("Sheet1", addr);

        // Bytecode-disabled engine (AST-only).
        let mut engine_ast = Engine::new();
        engine_ast.set_bytecode_enabled(false);
        engine_ast
            .set_cell_formula("Sheet1", addr, formula)
            .unwrap();
        assert_eq!(
            engine_ast.bytecode_program_count(),
            0,
            "bytecode-disabled engine should not compile programs"
        );
        let levels = engine_ast
            .calc_graph
            .calc_levels_for_dirty()
            .expect("calc levels");
        let _ = engine_ast.recalculate_levels(
            levels,
            RecalcMode::SingleThreaded,
            &recalc_ctx,
            engine_ast.date_system,
            None,
        );
        let value_ast = engine_ast.get_cell_value("Sheet1", addr);

        assert_eq!(value_bc, value_ast);
        value_bc
    }

    #[test]
    fn bytecode_if_is_lazy_in_false_branch() {
        assert_bytecode_matches_ast("=IF(TRUE, \"x\", 1/0)", Value::Text("x".to_string()));
    }

    #[test]
    fn bytecode_if_is_lazy_in_true_branch() {
        assert_bytecode_matches_ast("=IF(FALSE, 1/0, 7)", Value::Number(7.0));
    }

    #[test]
    fn bytecode_if_missing_false_defaults_to_false() {
        assert_bytecode_matches_ast("=IF(FALSE, 7)", Value::Bool(false));
    }

    #[test]
    fn bytecode_iferror_is_lazy_in_fallback() {
        assert_bytecode_matches_ast("=IFERROR(1, 1/0)", Value::Number(1.0));
    }

    #[test]
    fn bytecode_iferror_evaluates_fallback_on_error() {
        assert_bytecode_matches_ast("=IFERROR(1/0, 7)", Value::Number(7.0));
    }

    #[test]
    fn bytecode_iferror_evaluates_fallback_on_na() {
        assert_bytecode_matches_ast("=IFERROR(NA(), 7)", Value::Number(7.0));
    }

    #[test]
    fn bytecode_ifna_is_lazy_in_fallback() {
        assert_bytecode_matches_ast("=IFNA(1, 1/0)", Value::Number(1.0));
    }

    #[test]
    fn bytecode_ifna_does_not_use_fallback_for_non_na_errors() {
        assert_bytecode_matches_ast("=IFNA(1/0, 7)", Value::Error(ErrorKind::Div0));
    }

    #[test]
    fn bytecode_ifna_evaluates_fallback_on_na() {
        assert_bytecode_matches_ast("=IFNA(NA(), 7)", Value::Number(7.0));
    }

    #[test]
    fn bytecode_if_text_condition_is_coerced_to_bool() {
        assert_bytecode_matches_ast("=IF(\"FALSE\", 1/0, 7)", Value::Number(7.0));
    }

    #[test]
    fn bytecode_if_short_circuits_unused_volatile_branch() {
        let v = assert_bytecode_eq_ast("=IF(TRUE, 1, RAND()) + RAND()");
        match v {
            Value::Number(n) => assert!((1.0..2.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_if_short_circuits_unused_volatile_true_branch() {
        let v = assert_bytecode_eq_ast("=IF(FALSE, RAND(), 1) + RAND()");
        match v {
            Value::Number(n) => assert!((1.0..2.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_if_does_not_eval_branches_when_condition_is_error_even_when_volatile() {
        let v = assert_bytecode_eq_ast("=IFERROR(IF(1/0, RAND(), 1) + RAND(), RAND())");
        match v {
            Value::Number(n) => assert!((0.0..1.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_iferror_short_circuits_unused_volatile_fallback() {
        let v = assert_bytecode_eq_ast("=IFERROR(1, RAND()) + RAND()");
        match v {
            Value::Number(n) => assert!((1.0..2.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_ifna_short_circuits_unused_volatile_fallback() {
        let v = assert_bytecode_eq_ast("=IFNA(1, RAND()) + RAND()");
        match v {
            Value::Number(n) => assert!((1.0..2.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_ifna_does_not_eval_fallback_for_non_na_errors_even_when_volatile() {
        // `IFNA` should not evaluate its fallback when the first argument is a non-#N/A error.
        // Use RAND() draw indexing to detect accidental eager evaluation:
        // - The inner `+ RAND()` is always evaluated (even though the left side is an error),
        //   and the outer IFERROR fallback returns a subsequent RAND() draw.
        // - If IFNA eagerly evaluated its fallback RAND(), the visible returned draw would shift.
        let v = assert_bytecode_eq_ast("=IFERROR(IFNA(1/0, RAND()) + RAND(), RAND())");
        match v {
            Value::Number(n) => assert!((0.0..1.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_choose_short_circuits_unused_volatile_choice() {
        let v = assert_bytecode_eq_ast("=CHOOSE(2, RAND(), 1) + RAND()");
        match v {
            Value::Number(n) => assert!((1.0..2.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_choose_does_not_eval_choices_when_index_is_error_even_when_volatile() {
        let v = assert_bytecode_eq_ast("=IFERROR(CHOOSE(1/0, RAND(), 1) + RAND(), RAND())");
        match v {
            Value::Number(n) => assert!((0.0..1.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_choose_does_not_eval_choices_when_index_is_out_of_range_even_when_volatile() {
        // Use RAND() draw indexing to detect accidental eager evaluation of CHOOSE choices when
        // the index is invalid (#VALUE!).
        let v = assert_bytecode_eq_ast("=IFERROR(CHOOSE(3, RAND(), 1) + RAND(), RAND())");
        match v {
            Value::Number(n) => assert!((0.0..1.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_ifs_short_circuits_later_conditions() {
        assert_bytecode_matches_ast("=IFS(TRUE, 1, 1/0, 2)", Value::Number(1.0));
    }

    #[test]
    fn bytecode_ifs_short_circuits_values_for_false_conditions() {
        assert_bytecode_matches_ast("=IFS(FALSE, 1/0, TRUE, 2)", Value::Number(2.0));
    }

    #[test]
    fn bytecode_ifs_does_not_eval_values_when_no_condition_matches_even_when_volatile() {
        let v =
            assert_bytecode_eq_ast("=IFERROR(IFS(FALSE, RAND(), FALSE, RAND()) + RAND(), RAND())");
        match v {
            Value::Number(n) => assert!((0.0..1.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_ifs_short_circuits_unused_volatile_args() {
        let v = assert_bytecode_eq_ast("=IFS(TRUE, 1, TRUE, RAND()) + RAND()");
        match v {
            Value::Number(n) => assert!((1.0..2.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_ifs_short_circuits_unused_volatile_condition() {
        let v = assert_bytecode_eq_ast("=IFS(TRUE, 1, RAND(), 2) + RAND()");
        match v {
            Value::Number(n) => assert!((1.0..2.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_ifs_does_not_eval_later_pairs_when_condition_is_error_even_when_volatile() {
        let v = assert_bytecode_eq_ast("=IFERROR(IFS(1/0, RAND(), TRUE, 1) + RAND(), RAND())");
        match v {
            Value::Number(n) => assert!((0.0..1.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_switch_short_circuits_later_case_values() {
        assert_bytecode_matches_ast("=SWITCH(1, 1, 10, 1/0, 20)", Value::Number(10.0));
    }

    #[test]
    fn bytecode_switch_short_circuits_results_for_unmatched_cases() {
        assert_bytecode_matches_ast("=SWITCH(2, 1, 1/0, 2, 20)", Value::Number(20.0));
    }

    #[test]
    fn bytecode_switch_does_not_eval_results_when_no_case_matches_even_when_volatile() {
        let v =
            assert_bytecode_eq_ast("=IFERROR(SWITCH(3, 1, RAND(), 2, RAND()) + RAND(), RAND())");
        match v {
            Value::Number(n) => assert!((0.0..1.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_switch_short_circuits_unused_volatile_default() {
        let v = assert_bytecode_eq_ast("=SWITCH(1, 1, 10, RAND()) + RAND()");
        match v {
            Value::Number(n) => assert!((10.0..11.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_switch_short_circuits_unused_volatile_case_value() {
        let v = assert_bytecode_eq_ast("=SWITCH(1, 1, 10, RAND(), 20) + RAND()");
        match v {
            Value::Number(n) => assert!((10.0..11.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_switch_short_circuits_unused_volatile_case_result() {
        let v = assert_bytecode_eq_ast("=SWITCH(2, 1, RAND(), 2, 20) + RAND()");
        match v {
            Value::Number(n) => assert!((20.0..21.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_switch_does_not_eval_case_values_when_discriminant_is_error_even_when_volatile() {
        let v = assert_bytecode_eq_ast("=IFERROR(SWITCH(1/0, RAND(), 10, 2, 20) + RAND(), RAND())");
        match v {
            Value::Number(n) => assert!((0.0..1.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_switch_does_not_eval_case_results_when_case_value_is_error_even_when_volatile() {
        let v = assert_bytecode_eq_ast(
            "=IFERROR(SWITCH(2, 1/0, RAND(), RAND(), RAND()) + RAND(), RAND())",
        );
        match v {
            Value::Number(n) => assert!((0.0..1.0).contains(&n), "got {n}"),
            other => panic!("expected number, got {other:?}"),
        }
    }

    #[test]
    fn bytecode_compiler_allows_huge_ranges_for_implicit_intersection() {
        let mut engine = Engine::new();

        // A nearly full-sheet range (excluding the last column) is far beyond the bytecode range
        // cell-count limit, but implicit intersection only needs to check membership and/or
        // dereference a single cell.
        engine
            .set_cell_formula("Sheet1", "XFD1", "=@A1:XFC1048576")
            .unwrap();

        // Ensure we're exercising the bytecode path.
        assert_eq!(engine.bytecode_program_count(), 1);

        engine.recalculate_single_threaded();

        // XFD1 is outside the rectangle (the range ends at XFC), so implicit intersection should
        // fail with #VALUE!.
        assert_eq!(
            engine.get_cell_value("Sheet1", "XFD1"),
            Value::Error(ErrorKind::Value)
        );
    }

    #[test]
    fn bytecode_compiler_inlines_defined_name_constants_under_implicit_intersection() {
        let mut engine = Engine::new();
        engine
            .define_name(
                "X",
                NameScope::Workbook,
                NameDefinition::Constant(Value::Number(5.0)),
            )
            .unwrap();
        engine.set_cell_formula("Sheet1", "A1", "=@X").unwrap();

        // Ensure the name constant was inlined so the bytecode backend can compile the `@`
        // expression (bytecode lowering does not support NameRef directly).
        assert_eq!(engine.bytecode_program_count(), 1);

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(5.0));
    }

    #[test]
    fn implicit_intersection_range_dependencies_point_to_intersected_cell() {
        let mut engine = Engine::new();
        engine.set_cell_formula("Sheet1", "B2", "=@A1:A3").unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let b2 = parse_a1("B2").unwrap();
        let key = CellKey {
            sheet: sheet_id,
            addr: b2,
        };

        let precedents = engine.calc_graph.precedents_of(cell_id_from_key(key));
        assert_eq!(
            precedents,
            vec![Precedent::Cell(CellId::new(
                sheet_id_for_graph(sheet_id),
                1,
                0
            ))],
            "=@A1:A3 in row 2 should only depend on A2"
        );
    }

    #[test]
    fn bytecode_supports_structured_refs_by_resolving_tables() {
        use formula_model::table::TableColumn;

        fn table_fixture(range: &str) -> Table {
            Table {
                id: 1,
                name: "Table1".into(),
                display_name: "Table1".into(),
                range: Range::from_a1(range).unwrap(),
                header_row_count: 1,
                totals_row_count: 0,
                columns: vec![
                    TableColumn {
                        id: 1,
                        name: "Col1".into(),
                        formula: None,
                        totals_formula: None,
                    },
                    TableColumn {
                        id: 2,
                        name: "Col2".into(),
                        formula: None,
                        totals_formula: None,
                    },
                    TableColumn {
                        id: 3,
                        name: "Col3".into(),
                        formula: None,
                        totals_formula: None,
                    },
                ],
                style: None,
                auto_filter: None,
                relationship_id: None,
                part_path: None,
            }
        }

        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");
        engine.set_sheet_tables("Sheet1", vec![table_fixture("A1:C3")]);

        engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
        engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
        engine.set_cell_value("Sheet1", "B3", 20.0).unwrap();

        engine
            .set_cell_formula("Sheet1", "D1", "=SUM(Table1[Col1])")
            .unwrap();
        engine
            .set_cell_formula("Sheet1", "C2", "=[@Col1]+[@Col2]")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr_d1 = parse_a1("D1").unwrap();
        let addr_c2 = parse_a1("C2").unwrap();
        let cell_d1 = engine.workbook.sheets[sheet_id]
            .cells
            .get(&addr_d1)
            .expect("D1 stored");
        let cell_c2 = engine.workbook.sheets[sheet_id]
            .cells
            .get(&addr_c2)
            .expect("C2 stored");

        assert!(
            matches!(cell_d1.compiled, Some(CompiledFormula::Bytecode(_))),
            "structured refs should be eligible for bytecode after lowering"
        );
        assert!(
            matches!(cell_c2.compiled, Some(CompiledFormula::Bytecode(_))),
            "this-row structured refs should be eligible for bytecode after lowering"
        );

        engine.recalculate_single_threaded();
        let d1_bc = engine.get_cell_value("Sheet1", "D1");
        let c2_bc = engine.get_cell_value("Sheet1", "C2");
        assert_eq!(d1_bc, Value::Number(3.0));
        assert_eq!(c2_bc, Value::Number(11.0));

        // Compare bytecode vs AST evaluation.
        engine.set_bytecode_enabled(false);
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "D1"), d1_bc);
        assert_eq!(engine.get_cell_value("Sheet1", "C2"), c2_bc);
    }

    #[test]
    fn structured_ref_bytecode_recompiles_on_table_resize() {
        use formula_model::table::TableColumn;

        fn table_fixture(range: &str) -> Table {
            Table {
                id: 1,
                name: "Table1".into(),
                display_name: "Table1".into(),
                range: Range::from_a1(range).unwrap(),
                header_row_count: 1,
                totals_row_count: 0,
                columns: vec![TableColumn {
                    id: 1,
                    name: "Col1".into(),
                    formula: None,
                    totals_formula: None,
                }],
                style: None,
                auto_filter: None,
                relationship_id: None,
                part_path: None,
            }
        }

        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");
        engine.set_sheet_tables("Sheet1", vec![table_fixture("A1:A3")]);

        engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();

        engine
            .set_cell_formula("Sheet1", "B1", "=SUM(Table1[Col1])")
            .unwrap();

        assert_eq!(engine.bytecode_program_count(), 1);
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));

        // Resize the table to include an additional data row.
        engine.set_sheet_tables("Sheet1", vec![table_fixture("A1:A4")]);
        engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();

        // The new table extents should cause the structured ref to be re-lowered and recompiled,
        // producing a new program key.
        assert_eq!(engine.bytecode_program_count(), 2);
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
    }

    #[test]
    fn bytecode_supports_multi_area_structured_refs() {
        use formula_model::table::TableColumn;

        fn table_fixture(range: &str) -> Table {
            Table {
                id: 1,
                name: "Table1".into(),
                display_name: "Table1".into(),
                range: Range::from_a1(range).unwrap(),
                header_row_count: 1,
                totals_row_count: 0,
                columns: vec![
                    TableColumn {
                        id: 1,
                        name: "Col1".into(),
                        formula: None,
                        totals_formula: None,
                    },
                    TableColumn {
                        id: 2,
                        name: "Col2".into(),
                        formula: None,
                        totals_formula: None,
                    },
                    TableColumn {
                        id: 3,
                        name: "Col3".into(),
                        formula: None,
                        totals_formula: None,
                    },
                ],
                style: None,
                auto_filter: None,
                relationship_id: None,
                part_path: None,
            }
        }

        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");
        engine.set_sheet_tables("Sheet1", vec![table_fixture("A1:C3")]);

        engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
        engine.set_cell_value("Sheet1", "C2", 3.0).unwrap();
        engine.set_cell_value("Sheet1", "C3", 4.0).unwrap();

        engine
            .set_cell_formula("Sheet1", "D1", "=SUM(Table1[[Col1],[Col3]])")
            .unwrap();

        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let addr_d1 = parse_a1("D1").unwrap();
        let cell_d1 = engine.workbook.sheets[sheet_id]
            .cells
            .get(&addr_d1)
            .expect("D1 stored");
        assert!(
            matches!(cell_d1.compiled, Some(CompiledFormula::Bytecode(_))),
            "multi-area structured refs should be eligible for bytecode after lowering"
        );

        engine.recalculate_single_threaded();
        let d1_bc = engine.get_cell_value("Sheet1", "D1");
        assert_eq!(d1_bc, Value::Number(10.0));

        engine.set_bytecode_enabled(false);
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "D1"), d1_bc);
    }

    #[test]
    fn bytecode_dependency_trace_engine_updates_dependents_from_bytecode_eval() {
        let mut engine = Engine::new();
        engine
            .set_cell_value("Sheet1", "A1", Value::Number(10.0))
            .unwrap();
        engine.set_cell_formula("Sheet1", "B1", "=1").unwrap();

        // Force B1 to be treated as a dynamic-dependency formula, but swap in a bytecode program
        // that dereferences A1 via a reference value. This validates that bytecode evaluation can
        // produce dynamic dependency traces that the engine consumes to update its calc graph.
        let sheet_id = engine.workbook.sheet_id("Sheet1").expect("sheet exists");
        let b1_addr = parse_a1("B1").unwrap();
        let cell = engine.workbook.sheets[sheet_id]
            .cells
            .get_mut(&b1_addr)
            .expect("B1 stored");

        cell.dynamic_deps = true;
        let ast = cell.compiled.as_ref().expect("compiled").ast().clone();

        let mut program =
            bytecode::Program::new(Arc::from("bytecode_dependency_trace_engine_test"));
        program.range_refs.push(bytecode::RangeRef::new(
            bytecode::Ref::new(0, 0, true, true), // A1
            bytecode::Ref::new(0, 0, true, true), // A1
        ));
        program.instrs.push(bytecode::Instruction::new(
            bytecode::OpCode::LoadRange,
            0,
            0,
        ));

        cell.compiled = Some(CompiledFormula::Bytecode(BytecodeFormula {
            ast,
            program: Arc::new(program),
            sheet_dims_generation: engine.sheet_dims_generation,
        }));

        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));

        let dependents = engine.dependents("Sheet1", "A1").unwrap();
        assert!(
            dependents.contains(&PrecedentNode::Cell {
                sheet: 0,
                addr: CellAddr { row: 0, col: 1 } // B1
            }),
            "B1 should become a dependent of A1 after bytecode dependency tracing"
        );

        engine
            .set_cell_value("Sheet1", "A1", Value::Number(20.0))
            .unwrap();
        engine.recalculate_single_threaded();
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(20.0));
    }

    #[test]
    fn engine_bytecode_grid_column_slice_rejects_out_of_bounds_columns() {
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");

        let snapshot = Snapshot::from_workbook(
            &engine.workbook,
            &engine.spills,
            engine.external_value_provider.clone(),
            engine.external_data_provider.clone(),
            engine.info.clone(),
            engine.pivot_registry.clone(),
        );
        let sheet_cols = i32::try_from(snapshot.sheet_dimensions[0].1).unwrap_or(i32::MAX);

        let mut cols = HashMap::new();
        cols.insert(
            sheet_cols,
            BytecodeColumn {
                segments: vec![BytecodeColumnSegment {
                    row_start: 0,
                    values: vec![1.0],
                    blocked_rows_strict: Vec::new(),
                    blocked_rows_ignore_nonnumeric: Vec::new(),
                }],
            },
        );

        let grid = EngineBytecodeGrid {
            snapshot: &snapshot,
            sheet_id: 0,
            cols: &cols,
            cols_by_sheet: std::slice::from_ref(&cols),
            slice_mode: ColumnSliceMode::IgnoreNonNumeric,
            trace: None,
        };

        assert!(
            bytecode::grid::Grid::column_slice(&grid, sheet_cols, 0, 0).is_none(),
            "out-of-bounds columns should never be eligible for SIMD slicing"
        );
    }

    #[test]
    fn bytecode_multi_sheet_column_slice_uses_correct_sheet() {
        // Regression test for multi-sheet SIMD slices:
        // `Grid::column_slice_on_sheet` must return slices for the requested sheet, not the current
        // sheet. The default `Grid` implementation ignores the sheet id and can incorrectly
        // double-count the current sheet when SIMD fast paths are used for 3D aggregates.
        let mut engine = Engine::new();
        engine.ensure_sheet("Sheet1");
        engine.ensure_sheet("Sheet2");

        engine
            .set_cell_value("Sheet1", "A1", Value::Number(1.0))
            .unwrap();
        engine
            .set_cell_value("Sheet1", "A2", Value::Number(2.0))
            .unwrap();
        engine
            .set_cell_value("Sheet2", "A1", Value::Number(10.0))
            .unwrap();
        engine
            .set_cell_value("Sheet2", "A2", Value::Number(20.0))
            .unwrap();

        engine
            .set_cell_formula(
                "Sheet1",
                "B1",
                // The first SUM forces a column slice to exist for Sheet1, while the second SUM
                // evaluates a 3D sheet span that must read Sheet2's slice via `column_slice_on_sheet`.
                "=SUM(A1:A2) + SUM(Sheet1:Sheet2!A1:A2)",
            )
            .unwrap();

        engine.recalculate_single_threaded();

        // Expected:
        // SUM(Sheet1!A1:A2) = 3
        // SUM(Sheet1:Sheet2!A1:A2) = 3 + 30 = 33
        // Total = 36
        assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(36.0));
    }
}
