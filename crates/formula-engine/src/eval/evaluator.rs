use crate::calc_settings::CalculationMode;
use crate::date::ExcelDateSystem;
use crate::error::ExcelError;
use crate::eval::address::CellAddr;
use crate::eval::ast::{
    BinaryOp, CompareOp, CompiledExpr, Expr, PostfixOp, SheetReference, UnaryOp,
};
use crate::functions::{
    ArgValue as FnArgValue, FunctionContext, Reference as FnReference, SheetId as FnSheetId,
};
use crate::locale::ValueLocaleConfig;
use crate::value::{casefold, cmp_case_insensitive, Array, ErrorKind, Lambda, NumberLocale, Value};
use crate::LocaleConfig;
use formula_model::HorizontalAlignment;
use std::cell::{Cell, RefCell};
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use std::sync::Arc;

/// Synthetic call name used for anonymous `LAMBDA(...)(...)` invocations.
///
/// The leading NUL byte ensures this key cannot be referenced by user formulas.
const ANON_LAMBDA_CALL_NAME: &str = "\u{0}ANON_LAMBDA_CALL";

/// Maximum number of cells the engine will materialize into an in-memory [`Value::Array`].
///
/// This is primarily used to keep evaluation robust when dynamic sheet dimensions make ranges
/// effectively unbounded. For example:
/// - A bare reference result like `=Sheet1!A:XFD` would otherwise attempt to allocate ~17B cells.
/// - Array-producing functions like `ROW(A:A)` could become enormous if a sheet grows to billions
///   of rows.
///
/// When an operation would exceed this limit, evaluation returns `#SPILL!` ("Spill range is too
/// big") instead of attempting a huge allocation / long-running loop.
pub(crate) const MAX_MATERIALIZED_ARRAY_CELLS: usize = 5_000_000;

// Excel has various nesting limits (e.g. 64 nested function calls). Keep lambda recursion bounded
// well below the Rust stack limit to avoid process aborts for accidental infinite recursion.
const LAMBDA_RECURSION_LIMIT: u32 = 64;

#[derive(Debug, Clone, Copy)]
pub struct EvalContext {
    pub current_sheet: usize,
    pub current_cell: CellAddr,
}

#[derive(Debug, Clone)]
pub struct RecalcContext {
    pub now_utc: chrono::DateTime<chrono::Utc>,
    pub recalc_id: u64,
    pub number_locale: NumberLocale,
    pub calculation_mode: CalculationMode,
}

impl RecalcContext {
    pub fn new(recalc_id: u64) -> Self {
        Self {
            now_utc: chrono::Utc::now(),
            recalc_id,
            number_locale: NumberLocale::en_us(),
            calculation_mode: CalculationMode::Automatic,
        }
    }
}

/// Dynamic-dependency trace captured during formula evaluation.
///
/// This records which references were actually dereferenced (directly or via functions). The
/// engine uses this to update the dependency graph for formulas with dynamic reference behavior
/// (e.g. INDIRECT/OFFSET).
#[derive(Debug, Default, Clone)]
pub struct DependencyTrace {
    precedents: HashSet<FnReference>,
}

impl DependencyTrace {
    pub fn record_reference(&mut self, reference: FnReference) {
        let reference = reference.normalized();

        // Keep the trace compact by discarding precedents that are fully subsumed by a larger
        // rectangle on the same sheet.
        //
        // This matters for reference-returning functions like INDEX: callers may record the
        // entire input range as a dynamic precedent while the evaluator later dereferences the
        // selected single-cell reference for the final value. In that case, the single cell is
        // redundant once the full range has been recorded.
        if self.precedents.iter().any(|existing| {
            existing.sheet_id == reference.sheet_id
                && existing.start.row <= reference.start.row
                && existing.start.col <= reference.start.col
                && existing.end.row >= reference.end.row
                && existing.end.col >= reference.end.col
        }) {
            return;
        }

        self.precedents.retain(|existing| {
            !(existing.sheet_id == reference.sheet_id
                && reference.start.row <= existing.start.row
                && reference.start.col <= existing.start.col
                && reference.end.row >= existing.end.row
                && reference.end.col >= existing.end.col)
        });

        self.precedents.insert(reference);
    }

    #[must_use]
    pub fn precedents(&self, sheet_tab_index: impl Fn(usize) -> usize) -> Vec<FnReference> {
        let mut out: Vec<FnReference> = self.precedents.iter().cloned().collect();
        out.sort_by(|a, b| {
            match (&a.sheet_id, &b.sheet_id) {
                (FnSheetId::Local(a_sheet), FnSheetId::Local(b_sheet)) => sheet_tab_index(*a_sheet)
                    .cmp(&sheet_tab_index(*b_sheet))
                    .then_with(|| a_sheet.cmp(b_sheet)),
                (FnSheetId::Local(_), FnSheetId::External(_)) => Ordering::Less,
                (FnSheetId::External(_), FnSheetId::Local(_)) => Ordering::Greater,
                (FnSheetId::External(a_key), FnSheetId::External(b_key)) => a_key.cmp(b_key),
            }
            .then_with(|| a.start.row.cmp(&b.start.row))
            .then_with(|| a.start.col.cmp(&b.start.col))
            .then_with(|| a.end.row.cmp(&b.end.row))
            .then_with(|| a.end.col.cmp(&b.end.col))
        });
        out
    }
}

pub trait ValueResolver {
    fn sheet_exists(&self, sheet_id: usize) -> bool;
    /// Returns the current worksheet tab order index for `sheet_id`.
    ///
    /// Excel defines the ordering of multi-area references (e.g. 3D sheet spans like
    /// `Sheet1:Sheet3!A1`, or `INDEX(..., area_num)`) based on workbook sheet tab order, not the
    /// internal numeric sheet id.
    ///
    /// Most resolvers historically used sheet ids that matched tab order, so the default
    /// implementation preserves the old behavior by treating the sheet id itself as the order
    /// index. Resolvers with stable sheet ids should override this to return the current tab
    /// position.
    fn sheet_order_index(&self, sheet_id: usize) -> Option<usize> {
        self.sheet_exists(sheet_id).then_some(sheet_id)
    }
    /// Workbook legacy text codepage used for DBCS (`*B`) text functions.
    ///
    /// Defaults to Excel's en-US codepage (1252).
    fn text_codepage(&self) -> u16 {
        1252
    }
    /// Return the number of worksheets available in the workbook.
    ///
    /// This is used by worksheet information functions like `INFO("numfile")`.
    fn sheet_count(&self) -> usize {
        1
    }
    /// Expand a 3D sheet span into its component sheets in workbook tab order.
    ///
    /// Excel resolves `Sheet1:Sheet3!A1` by including all sheets between the boundary sheets in the
    /// workbook's **current sheet tab order**, inclusive. This means:
    /// - Intermediate sheets are determined by the current tab order (not by numeric sheet id).
    /// - Reversed spans are allowed (e.g. `Sheet3:Sheet1!A1`) and refer to the same set of sheets.
    ///
    /// Returning `None` indicates either boundary sheet could not be located in the workbook order
    /// (and should therefore evaluate to `#REF!`).
    ///
    /// The default implementation preserves historical behavior by expanding to the numeric
    /// `min..=max` sheet id range.
    fn expand_sheet_span(&self, start_sheet_id: usize, end_sheet_id: usize) -> Option<Vec<usize>> {
        if !self.sheet_exists(start_sheet_id) || !self.sheet_exists(end_sheet_id) {
            return None;
        }
        let (start, end) = if start_sheet_id <= end_sheet_id {
            (start_sheet_id, end_sheet_id)
        } else {
            (end_sheet_id, start_sheet_id)
        };
        // Historical behavior expands to the numeric id range. Filter out missing sheets so deleted
        // ids (which are never reused) do not introduce spurious `#REF!` results.
        Some((start..=end).filter(|id| self.sheet_exists(*id)).collect())
    }

    /// Host-provided system metadata used by the Excel `INFO()` worksheet function.
    ///
    /// The engine does not attempt to query the real OS at runtime; to keep evaluation portable
    /// and deterministic, hosts may populate these values explicitly (e.g. via `EngineInfo`).
    fn info_system(&self) -> Option<&str> {
        None
    }
    fn info_directory(&self) -> Option<&str> {
        None
    }
    fn info_osversion(&self) -> Option<&str> {
        None
    }
    fn info_release(&self) -> Option<&str> {
        None
    }
    fn info_version(&self) -> Option<&str> {
        None
    }
    fn info_memavail(&self) -> Option<f64> {
        None
    }
    fn info_totmem(&self) -> Option<f64> {
        None
    }
    /// Returns the upper-left visible cell for `sheet_id`, as an absolute A1 reference (e.g.
    /// `"$A$1"`).
    fn info_origin(&self, _sheet_id: usize) -> Option<&str> {
        None
    }
    /// Returns the current (row_count, col_count) dimensions for a sheet.
    ///
    /// Coordinates are in-bounds iff:
    /// - `row < row_count`
    /// - `col < col_count`
    ///
    /// Implementations that don't track dynamic dimensions can rely on the default, which matches
    /// Excel's default worksheet bounds.
    fn sheet_dimensions(&self, _sheet_id: usize) -> (u32, u32) {
        (formula_model::EXCEL_MAX_ROWS, formula_model::EXCEL_MAX_COLS)
    }
    /// Returns the sheet default column width in Excel "character" units.
    ///
    /// This corresponds to the worksheet's `<sheetFormatPr defaultColWidth="...">` metadata.
    fn sheet_default_col_width(&self, _sheet_id: usize) -> Option<f32> {
        None
    }

    /// Returns the top-left visible cell ("origin") for a worksheet view, if provided.
    ///
    /// This corresponds to Excel's `INFO("origin")` semantics (driven by scroll position + frozen
    /// panes). The core engine is deterministic and must not consult UI/window state directly, so
    /// hosts should plumb this metadata through.
    fn sheet_origin_cell(&self, _sheet_id: usize) -> Option<CellAddr> {
        None
    }
    fn get_cell_value(&self, sheet_id: usize, addr: CellAddr) -> Value;
    /// Returns the effective horizontal alignment for the given cell, if available.
    ///
    /// This is used by worksheet information functions like `CELL("prefix")`.
    fn cell_horizontal_alignment(
        &self,
        _sheet_id: usize,
        _addr: CellAddr,
    ) -> Option<HorizontalAlignment> {
        None
    }
    /// Resolve a sheet id back to its display name.
    ///
    /// This is used by worksheet information functions like `CELL("address")`.
    /// Resolvers that do not track sheet display names can return `None`.
    fn sheet_name(&self, _sheet_id: usize) -> Option<&str> {
        None
    }
    /// Return the stored formula text for a cell (including the leading `=`), if available.
    ///
    /// This is used by `CELL("contents")` and future formula-text functions.
    fn get_cell_formula(&self, _sheet_id: usize, _addr: CellAddr) -> Option<&str> {
        None
    }

    /// Returns the stored phonetic guide (furigana) text for a cell, if available.
    ///
    /// This metadata is used by Excel's `PHONETIC(...)` worksheet function. Most resolvers do not
    /// model phonetic guides, so the default implementation returns `None`.
    fn get_cell_phonetic(&self, _sheet_id: usize, _addr: CellAddr) -> Option<&str> {
        None
    }

    /// Returns the workbook's style table, if available.
    fn style_table(&self) -> Option<&formula_model::StyleTable> {
        None
    }

    /// Return the default style id for an entire worksheet, if present.
    fn sheet_default_style_id(&self, _sheet_id: usize) -> Option<u32> {
        None
    }

    /// Return the style id for a specific cell.
    ///
    /// Style id `0` is always the default (empty) style.
    fn cell_style_id(&self, _sheet_id: usize, _addr: CellAddr) -> u32 {
        0
    }

    /// Return the style id from the compressed range-run formatting layer for a cell, if present.
    ///
    /// Runs are defined per-column as row intervals and have precedence:
    /// `sheet < col < row < range-run < cell`.
    ///
    /// Style id `0` indicates no range-run override.
    fn format_run_style_id(&self, _sheet_id: usize, _addr: CellAddr) -> u32 {
        0
    }

    /// Return the default style id for an entire row, if present.
    fn row_style_id(&self, _sheet_id: usize, _row: u32) -> Option<u32> {
        None
    }

    /// Return per-column properties (width/hidden/default style), if present.
    fn col_properties(&self, _sheet_id: usize, _col: u32) -> Option<formula_model::ColProperties> {
        None
    }

    /// Return the style id from the range-run formatting layer for a cell, if present.
    ///
    /// This corresponds to DocumentController's `formatRunsByCol` layer (large range formatting
    /// rectangles compressed into per-column runs).
    ///
    /// Style id `0` indicates "no run applies".
    fn range_run_style_id(&self, _sheet_id: usize, _addr: CellAddr) -> u32 {
        0
    }

    /// Optional workbook directory metadata (typically with a trailing path separator).
    fn workbook_directory(&self) -> Option<&str> {
        None
    }

    /// Optional workbook filename metadata (e.g. `Book1.xlsx`).
    fn workbook_filename(&self) -> Option<&str> {
        None
    }

    /// Return the number format string for a cell, if available.
    ///
    /// This is used by worksheet information functions like `CELL("format")` /
    /// `CELL("color")` / `CELL("parentheses")`.
    fn get_cell_number_format(&self, _sheet_id: usize, _addr: CellAddr) -> Option<&str> {
        None
    }
    /// Resolve a value from an external workbook reference like `[Book.xlsx]Sheet1!A1`.
    ///
    /// The `sheet` key is the canonical bracketed form (`"[Book.xlsx]Sheet1"`).
    /// Returning `None` indicates the external reference could not be resolved and should
    /// evaluate to `#REF!`.
    fn get_external_value(&self, _sheet: &str, _addr: CellAddr) -> Option<Value> {
        None
    }
    /// Return the sheet order for an external workbook.
    ///
    /// This is used to expand external-workbook 3D spans like `[Book.xlsx]Sheet1:Sheet3!A1`,
    /// which must be resolved by workbook sheet order. Implementations should return sheet names
    /// (without the `[Book.xlsx]` prefix) in workbook order.
    ///
    /// Expected semantics: endpoint matching should behave like Excel's sheet name comparison
    /// (Unicode-aware, NFKC + case-insensitive), as implemented by
    /// [`formula_model::sheet_name_eq_case_insensitive`].
    ///
    /// Returning `None` indicates that the sheet order is unavailable, in which case external
    /// 3D spans evaluate to `#REF!`.
    fn external_sheet_order(&self, _workbook: &str) -> Option<Vec<String>> {
        None
    }

    /// Return the sheet order for an external workbook as an `Arc` slice.
    ///
    /// This mirrors [`crate::ExternalValueProvider::workbook_sheet_names`] and allows resolvers to
    /// return cached sheet lists without cloning. The default implementation forwards to
    /// [`ValueResolver::external_sheet_order`].
    fn workbook_sheet_names(&self, workbook: &str) -> Option<Arc<[String]>> {
        self.external_sheet_order(workbook).map(Arc::from)
    }

    /// Return table metadata for an external workbook.
    ///
    /// This is used to resolve external workbook structured references like
    /// `"[Book.xlsx]Sheet1!Table1[Col]"`.
    ///
    /// The input `workbook` is the raw name inside the bracketed prefix (e.g. `"Book.xlsx"`),
    /// matching what [`crate::external_refs::parse_external_key`] extracts from a
    /// `"[workbook]Sheet"` key.
    ///
    /// Returning `None` indicates that table metadata is unavailable, in which case external
    /// structured references evaluate to `#REF!`.
    fn external_workbook_table(
        &self,
        _workbook: &str,
        _table_name: &str,
    ) -> Option<(String, formula_model::Table)> {
        None
    }
    /// Optional external data provider used by RTD / CUBE* functions.
    fn external_data_provider(&self) -> Option<&dyn crate::ExternalDataProvider> {
        None
    }
    /// Returns the in-memory pivot registry (if available) for resolving `GETPIVOTDATA`.
    fn pivot_registry(&self) -> Option<&crate::pivot_registry::PivotRegistry> {
        None
    }
    /// Resolve a worksheet name to an internal sheet id.
    ///
    /// This is used by volatile reference functions like `INDIRECT` that parse sheet names
    /// at runtime.
    ///
    /// Expected semantics: match Excel's Unicode-aware case-insensitive sheet name comparison.
    /// The engine approximates this by applying Unicode NFKC (compatibility normalization) and
    /// then Unicode uppercasing, as implemented by [`formula_model::sheet_name_eq_case_insensitive`].
    ///
    /// Resolvers that do not support name-based sheet lookup can return `None`.
    fn sheet_id(&self, _name: &str) -> Option<usize> {
        None
    }
    /// Iterates stored cells in `sheet_id`.
    ///
    /// For sparse backends, this should enumerate only populated addresses. Evaluators use this
    /// to implement sparse-aware aggregation over large ranges (e.g. `A:A`).
    fn iter_sheet_cells(
        &self,
        _sheet_id: usize,
    ) -> Option<Box<dyn Iterator<Item = CellAddr> + '_>> {
        None
    }
    fn resolve_structured_ref(
        &self,
        ctx: EvalContext,
        sref: &crate::structured_refs::StructuredRef,
    ) -> Result<Vec<(usize, CellAddr, CellAddr)>, ErrorKind>;
    fn resolve_name(&self, _sheet_id: usize, _name: &str) -> Option<ResolvedName> {
        None
    }
    /// If `addr` is part of a spilled array, returns the spill origin cell.
    fn spill_origin(&self, _sheet_id: usize, _addr: CellAddr) -> Option<CellAddr> {
        None
    }
    /// If `origin` is the origin of a spilled array, returns the full spill range (inclusive).
    fn spill_range(&self, _sheet_id: usize, _origin: CellAddr) -> Option<(CellAddr, CellAddr)> {
        None
    }

    /// System identifier surfaced by `INFO("system")` (e.g. `pcdos`).
    fn system_info(&self) -> Option<&str> {
        None
    }

    /// Origin string surfaced by `INFO("origin")`.
    fn origin(&self) -> Option<&str> {
        None
    }

    /// Resolve effective formatting/protection metadata for a single cell.
    ///
    /// This is used by worksheet information functions like `CELL("prefix")` and
    /// `CELL("protect")`.
    ///
    /// Implementations that do not track formatting can return the default style values.
    fn effective_cell_style(
        &self,
        _sheet_id: usize,
        _addr: CellAddr,
    ) -> crate::style_patch::EffectiveStyle {
        crate::style_patch::EffectiveStyle::default()
    }
}

#[derive(Debug, Clone)]
pub enum ResolvedName {
    Constant(Value),
    Expr(CompiledExpr),
}

#[derive(Debug, Clone)]
struct ResolvedRange {
    sheet_id: FnSheetId,
    start: CellAddr,
    end: CellAddr,
}

impl ResolvedRange {
    fn normalized(&self) -> Self {
        let (r1, r2) = if self.start.row <= self.end.row {
            (self.start.row, self.end.row)
        } else {
            (self.end.row, self.start.row)
        };
        let (c1, c2) = if self.start.col <= self.end.col {
            (self.start.col, self.end.col)
        } else {
            (self.end.col, self.start.col)
        };
        Self {
            sheet_id: self.sheet_id.clone(),
            start: CellAddr { row: r1, col: c1 },
            end: CellAddr { row: r2, col: c2 },
        }
    }

    fn is_single_cell(&self) -> bool {
        self.start == self.end
    }
}

#[derive(Debug, Clone)]
enum EvalValue {
    Scalar(Value),
    Reference(Vec<ResolvedRange>),
}

pub struct Evaluator<'a, R: ValueResolver> {
    resolver: &'a R,
    ctx: EvalContext,
    recalc_ctx: &'a RecalcContext,
    tracer: Option<&'a RefCell<DependencyTrace>>,
    name_stack: Rc<RefCell<Vec<(usize, String)>>>,
    lexical_scopes: Rc<RefCell<Vec<HashMap<String, Value>>>>,
    lambda_depth: Rc<Cell<u32>>,
    date_system: ExcelDateSystem,
    value_locale: ValueLocaleConfig,
    rng_counter: Rc<Cell<u64>>,
    locale: LocaleConfig,
    text_codepage: u16,
}

struct LexicalScopeGuard {
    stack: Rc<RefCell<Vec<HashMap<String, Value>>>>,
}

impl Drop for LexicalScopeGuard {
    fn drop(&mut self) {
        let mut stack = self.stack.borrow_mut();
        stack.pop();
    }
}

impl<'a, R: ValueResolver> Evaluator<'a, R> {
    fn cmp_sheet_ids_in_tab_order(&self, a: &FnSheetId, b: &FnSheetId) -> Ordering {
        match (a, b) {
            (FnSheetId::Local(a_id), FnSheetId::Local(b_id)) => {
                let a_idx = self.resolver.sheet_order_index(*a_id).unwrap_or(*a_id);
                let b_idx = self.resolver.sheet_order_index(*b_id).unwrap_or(*b_id);
                a_idx.cmp(&b_idx).then_with(|| a_id.cmp(b_id))
            }
            (FnSheetId::Local(_), FnSheetId::External(_)) => Ordering::Less,
            (FnSheetId::External(_), FnSheetId::Local(_)) => Ordering::Greater,
            (FnSheetId::External(a_key), FnSheetId::External(b_key)) => {
                // External references are keyed by the canonical sheet key (`"[Book.xlsx]Sheet1"`).
                // When the ValueResolver provides external workbook tab order, preserve that order
                // when sorting reference unions. This matters for Excel semantics like:
                // - INDEX(..., area_num)
                // - error precedence across multi-area unions
                //
                // Fall back to lexicographic ordering when workbook order is unavailable.
                match (
                    split_external_sheet_key_parts(a_key),
                    split_external_sheet_key_parts(b_key),
                ) {
                    (Some((a_wb, a_sheet)), Some((b_wb, b_sheet))) if a_wb == b_wb => {
                        match self.resolver.workbook_sheet_names(a_wb) {
                            Some(order) => {
                                let mut a_idx: Option<usize> = None;
                                let mut b_idx: Option<usize> = None;
                                for (idx, name) in order.iter().enumerate() {
                                    if a_idx.is_none()
                                        && formula_model::sheet_name_eq_case_insensitive(
                                            name, a_sheet,
                                        )
                                    {
                                        a_idx = Some(idx);
                                    }
                                    if b_idx.is_none()
                                        && formula_model::sheet_name_eq_case_insensitive(
                                            name, b_sheet,
                                        )
                                    {
                                        b_idx = Some(idx);
                                    }
                                    if a_idx.is_some() && b_idx.is_some() {
                                        break;
                                    }
                                }
                                match (a_idx, b_idx) {
                                    (Some(a_idx), Some(b_idx)) => {
                                        a_idx.cmp(&b_idx).then_with(|| a_key.cmp(b_key))
                                    }
                                    _ => a_key.cmp(b_key),
                                }
                            }
                            None => a_key.cmp(b_key),
                        }
                    }
                    _ => a_key.cmp(b_key),
                }
            }
        }
    }

    fn sort_resolved_ranges(&self, ranges: &mut [ResolvedRange]) {
        ranges.sort_by(|a, b| {
            self.cmp_sheet_ids_in_tab_order(&a.sheet_id, &b.sheet_id)
                .then_with(|| a.start.row.cmp(&b.start.row))
                .then_with(|| a.start.col.cmp(&b.start.col))
                .then_with(|| a.end.row.cmp(&b.end.row))
                .then_with(|| a.end.col.cmp(&b.end.col))
        });
    }

    pub fn new(resolver: &'a R, ctx: EvalContext, recalc_ctx: &'a RecalcContext) -> Self {
        Self::new_with_date_system_and_locales(
            resolver,
            ctx,
            recalc_ctx,
            ExcelDateSystem::EXCEL_1900,
            ValueLocaleConfig::default(),
            LocaleConfig::en_us(),
        )
    }

    pub fn new_with_date_system(
        resolver: &'a R,
        ctx: EvalContext,
        recalc_ctx: &'a RecalcContext,
        date_system: ExcelDateSystem,
    ) -> Self {
        Self::new_with_date_system_and_locales(
            resolver,
            ctx,
            recalc_ctx,
            date_system,
            ValueLocaleConfig::default(),
            LocaleConfig::en_us(),
        )
    }

    pub fn new_with_date_system_and_locale(
        resolver: &'a R,
        ctx: EvalContext,
        recalc_ctx: &'a RecalcContext,
        date_system: ExcelDateSystem,
        value_locale: ValueLocaleConfig,
    ) -> Self {
        Self::new_with_date_system_and_locales(
            resolver,
            ctx,
            recalc_ctx,
            date_system,
            value_locale,
            LocaleConfig::en_us(),
        )
    }

    pub fn new_with_date_system_and_locales(
        resolver: &'a R,
        ctx: EvalContext,
        recalc_ctx: &'a RecalcContext,
        date_system: ExcelDateSystem,
        value_locale: ValueLocaleConfig,
        locale: LocaleConfig,
    ) -> Self {
        let text_codepage = resolver.text_codepage();
        Self {
            resolver,
            ctx,
            recalc_ctx,
            tracer: None,
            name_stack: Rc::new(RefCell::new(Vec::new())),
            lexical_scopes: Rc::new(RefCell::new(Vec::new())),
            lambda_depth: Rc::new(Cell::new(0)),
            date_system,
            value_locale,
            rng_counter: Rc::new(Cell::new(0)),
            locale,
            // Default to the resolver's configured workbook text codepage.
            //
            // Most resolvers return 1252 (en-US / single-byte) which preserves the engine's
            // historical behavior. Engine-backed resolvers (e.g. Snapshot) can override this so
            // legacy DBCS functions (LENB/LEFTB/ASC/DBCS/...) respect workbook locale semantics.
            text_codepage,
        }
    }

    pub fn with_text_codepage(mut self, text_codepage: u16) -> Self {
        self.text_codepage = text_codepage;
        self
    }

    fn with_ctx(&self, ctx: EvalContext) -> Self {
        Self {
            resolver: self.resolver,
            ctx,
            recalc_ctx: self.recalc_ctx,
            tracer: self.tracer,
            name_stack: Rc::clone(&self.name_stack),
            lexical_scopes: Rc::clone(&self.lexical_scopes),
            lambda_depth: Rc::clone(&self.lambda_depth),
            date_system: self.date_system,
            value_locale: self.value_locale,
            rng_counter: Rc::clone(&self.rng_counter),
            locale: self.locale.clone(),
            text_codepage: self.text_codepage,
        }
    }

    fn with_lexical_scopes(&self, scopes: Vec<HashMap<String, Value>>) -> Self {
        Self {
            resolver: self.resolver,
            ctx: self.ctx,
            recalc_ctx: self.recalc_ctx,
            tracer: self.tracer,
            name_stack: Rc::clone(&self.name_stack),
            lexical_scopes: Rc::new(RefCell::new(scopes)),
            lambda_depth: Rc::clone(&self.lambda_depth),
            date_system: self.date_system,
            value_locale: self.value_locale,
            rng_counter: Rc::clone(&self.rng_counter),
            locale: self.locale.clone(),
            text_codepage: self.text_codepage,
        }
    }

    fn push_lexical_scope(&self, scope: HashMap<String, Value>) -> LexicalScopeGuard {
        self.lexical_scopes.borrow_mut().push(scope);
        LexicalScopeGuard {
            stack: Rc::clone(&self.lexical_scopes),
        }
    }

    fn lookup_lexical_value(&self, name: &str) -> Option<Value> {
        let key = casefold(name.trim());
        let scopes = self.lexical_scopes.borrow();
        for scope in scopes.iter().rev() {
            if let Some(value) = scope.get(&key) {
                return Some(value.clone());
            }
        }
        None
    }

    fn capture_lexical_env_map(&self) -> HashMap<String, Value> {
        let scopes = self.lexical_scopes.borrow();
        let mut out = HashMap::new();
        for scope in scopes.iter() {
            for (k, v) in scope {
                out.insert(k.clone(), v.clone());
            }
        }
        out
    }

    pub fn with_dependency_trace(mut self, trace: &'a RefCell<DependencyTrace>) -> Self {
        self.tracer = Some(trace);
        self
    }

    fn trace_reference(&self, reference: &FnReference) {
        let Some(trace) = self.tracer else {
            return;
        };
        trace.borrow_mut().record_reference(reference.clone());
    }

    fn trace_cell(&self, sheet_id: &FnSheetId, addr: CellAddr) {
        self.trace_reference(&FnReference {
            sheet_id: sheet_id.clone(),
            start: addr,
            end: addr,
        });
    }

    fn resolve_range_bounds(
        &self,
        sheet_id: &FnSheetId,
        start: CellAddr,
        end: CellAddr,
    ) -> Option<(CellAddr, CellAddr)> {
        let (rows, cols) = match sheet_id {
            FnSheetId::Local(id) => {
                if !self.resolver.sheet_exists(*id) {
                    return None;
                }
                self.resolver.sheet_dimensions(*id)
            }
            // External workbooks do not expose dimensions via the ValueResolver interface, so
            // treat the bounds as unknown and only resolve whole-row/whole-column sentinels using
            // Excel's default grid size.
            FnSheetId::External(_) => {
                (formula_model::EXCEL_MAX_ROWS, formula_model::EXCEL_MAX_COLS)
            }
        };
        let max_row = rows.saturating_sub(1);
        let max_col = cols.saturating_sub(1);

        let start = CellAddr {
            row: if start.row == CellAddr::SHEET_END {
                max_row
            } else {
                start.row
            },
            col: if start.col == CellAddr::SHEET_END {
                max_col
            } else {
                start.col
            },
        };
        let end = CellAddr {
            row: if end.row == CellAddr::SHEET_END {
                max_row
            } else {
                end.row
            },
            col: if end.col == CellAddr::SHEET_END {
                max_col
            } else {
                end.col
            },
        };

        if matches!(sheet_id, FnSheetId::Local(_))
            && (start.row >= rows || end.row >= rows || start.col >= cols || end.col >= cols)
        {
            return None;
        }
        Some((start, end))
    }

    fn function_result_to_eval_value(&self, value: Value) -> EvalValue {
        match value {
            Value::Reference(r) => {
                let Some((start, end)) = self.resolve_range_bounds(&r.sheet_id, r.start, r.end)
                else {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                };
                EvalValue::Reference(vec![ResolvedRange {
                    sheet_id: r.sheet_id,
                    start,
                    end,
                }])
            }
            Value::ReferenceUnion(ranges) => {
                let mut out = Vec::with_capacity(ranges.len());
                for r in ranges {
                    let Some((start, end)) = self.resolve_range_bounds(&r.sheet_id, r.start, r.end)
                    else {
                        return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                    };
                    out.push(ResolvedRange {
                        sheet_id: r.sheet_id,
                        start,
                        end,
                    });
                }
                EvalValue::Reference(out)
            }
            other => EvalValue::Scalar(other),
        }
    }

    /// Evaluate a compiled AST as a scalar formula result.
    pub fn eval_formula(&self, expr: &CompiledExpr) -> Value {
        match self.eval_value(expr) {
            EvalValue::Scalar(v) => v,
            EvalValue::Reference(range) => self.deref_reference_dynamic(range),
        }
    }

    fn eval_value(&self, expr: &CompiledExpr) -> EvalValue {
        match expr {
            Expr::Number(n) => EvalValue::Scalar(Value::Number(*n)),
            Expr::Text(s) => EvalValue::Scalar(Value::Text(s.clone())),
            Expr::Bool(b) => EvalValue::Scalar(Value::Bool(*b)),
            Expr::Blank => EvalValue::Scalar(Value::Blank),
            Expr::Error(e) => EvalValue::Scalar(Value::Error(*e)),
            Expr::ArrayLiteral { rows, cols, values } => {
                let mut out = Vec::with_capacity(rows.saturating_mul(*cols));
                for el in values.iter() {
                    let v = match self.eval_value(el) {
                        EvalValue::Scalar(v) => v,
                        EvalValue::Reference(ranges) => self.apply_implicit_intersection(&ranges),
                    };

                    let v = match v {
                        Value::Array(_) | Value::Spill { .. } => Value::Error(ErrorKind::Value),
                        other => other,
                    };
                    out.push(v);
                }
                EvalValue::Scalar(Value::Array(Array::new(*rows, *cols, out)))
            }
            Expr::CellRef(r) => match self.resolve_sheet_ids(&r.sheet) {
                Some(sheet_ids) => {
                    let Some(addr) = r.addr.resolve(self.ctx.current_cell) else {
                        return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                    };
                    let mut ranges = Vec::with_capacity(sheet_ids.len());
                    for sheet_id in sheet_ids {
                        if matches!(&sheet_id, FnSheetId::Local(id) if !self.resolver.sheet_exists(*id))
                        {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        }
                        let Some((start, end)) = self.resolve_range_bounds(&sheet_id, addr, addr)
                        else {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        };
                        ranges.push(ResolvedRange {
                            sheet_id,
                            start,
                            end,
                        });
                    }
                    EvalValue::Reference(ranges)
                }
                None => EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
            },
            Expr::RangeRef(r) => match self.resolve_sheet_ids(&r.sheet) {
                Some(sheet_ids) => {
                    let Some(start_addr) = r.start.resolve(self.ctx.current_cell) else {
                        return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                    };
                    let Some(end_addr) = r.end.resolve(self.ctx.current_cell) else {
                        return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                    };
                    let mut ranges = Vec::with_capacity(sheet_ids.len());
                    for sheet_id in sheet_ids {
                        if matches!(&sheet_id, FnSheetId::Local(id) if !self.resolver.sheet_exists(*id))
                        {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        }
                        let Some((start, end)) =
                            self.resolve_range_bounds(&sheet_id, start_addr, end_addr)
                        else {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        };
                        ranges.push(ResolvedRange {
                            sheet_id,
                            start,
                            end,
                        });
                    }
                    EvalValue::Reference(ranges)
                }
                None => EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
            },
            Expr::StructuredRef(sref_expr) => {
                // External workbook structured references (e.g. `[Book.xlsx]Sheet1!Table1[Col]`)
                // are resolved dynamically using provider-supplied table metadata.
                if let SheetReference::External(key) = &sref_expr.sheet {
                    if !key.starts_with('[') {
                        // `SheetReference::External` without a bracketed workbook prefix represents
                        // an invalid/missing sheet at compile time; preserve `#REF!` semantics.
                        return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                    }

                    let (workbook, explicit_sheet_key) = if let Some((workbook, _sheet)) =
                        crate::external_refs::parse_external_key(key)
                    {
                        (workbook, Some(key.as_str()))
                    } else if crate::external_refs::parse_external_span_key(key).is_some() {
                        // External 3D sheet spans are not valid structured-ref prefixes.
                        return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                    } else {
                        let Some(workbook) = crate::external_refs::parse_external_workbook_key(key)
                        else {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        };
                        (workbook, None)
                    };

                    let Some(table_name) = sref_expr.sref.table_name.as_deref() else {
                        return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                    };

                    // Excel's `[@ThisRow]` semantics depend on the formula being inside the table.
                    // For external workbooks we do not currently model the row context, so return
                    // `#REF!`.
                    if sref_expr.sref.items.iter().any(|item| {
                        matches!(item, crate::structured_refs::StructuredRefItem::ThisRow)
                    }) {
                        return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                    }

                    let Some((table_sheet, table)) =
                        self.resolver.external_workbook_table(workbook, table_name)
                    else {
                        return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                    };

                    let sheet_key = explicit_sheet_key
                        .map(|s| s.to_string())
                        .unwrap_or_else(|| {
                            crate::external_refs::format_external_key(workbook, &table_sheet)
                        });

                    let ranges = match crate::structured_refs::resolve_structured_ref_in_table(
                        &table,
                        self.ctx.current_cell,
                        &sref_expr.sref,
                    ) {
                        Ok(ranges) => ranges,
                        Err(_) => return EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
                    };

                    if ranges.is_empty() {
                        return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                    }

                    return EvalValue::Reference(
                        ranges
                            .into_iter()
                            .map(|(start, end)| ResolvedRange {
                                sheet_id: FnSheetId::External(sheet_key.clone()),
                                start,
                                end,
                            })
                            .collect(),
                    );
                }

                // Local structured references resolve via workbook table metadata.
                match self
                    .resolver
                    .resolve_structured_ref(self.ctx, &sref_expr.sref)
                {
                    Ok(ranges) if !ranges.is_empty() => {
                        if !ranges
                            .iter()
                            .all(|(sheet_id, _, _)| self.resolver.sheet_exists(*sheet_id))
                        {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Name));
                        }

                        if !ranges.iter().all(|(sheet_id, start, end)| {
                            self.reference_endpoints_in_bounds(
                                &FnSheetId::Local(*sheet_id),
                                *start,
                                *end,
                            )
                        }) {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        }

                        EvalValue::Reference(
                            ranges
                                .into_iter()
                                .map(|(sheet_id, start, end)| ResolvedRange {
                                    sheet_id: FnSheetId::Local(sheet_id),
                                    start,
                                    end,
                                })
                                .collect(),
                        )
                    }
                    Ok(_) => EvalValue::Scalar(Value::Error(ErrorKind::Ref)),
                    Err(e) => EvalValue::Scalar(Value::Error(e)),
                }
            }
            Expr::NameRef(nref) => self.eval_name_ref(nref),
            Expr::FieldAccess { base, field } => {
                let base = self.deref_eval_value_dynamic(self.eval_value(base));
                let field_key = field.as_str();
                if field_key.trim().is_empty() {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Value));
                }
                EvalValue::Scalar(elementwise_unary(&base, |elem| {
                    eval_field_access(elem, field_key)
                }))
            }
            Expr::SpillRange(inner) => {
                match self.eval_value(inner) {
                    EvalValue::Scalar(Value::Error(e)) => EvalValue::Scalar(Value::Error(e)),
                    EvalValue::Scalar(_) => EvalValue::Scalar(Value::Error(ErrorKind::Value)),
                    EvalValue::Reference(mut ranges) => {
                        // Spill-range references are only well-defined for a single-cell reference.
                        if ranges.len() != 1 {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Value));
                        }
                        let range = ranges.pop().expect("checked len() above");
                        if !range.is_single_cell() {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Value));
                        }

                        let FnSheetId::Local(sheet_id) = range.sheet_id else {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        };
                        let Some(origin) = self.resolver.spill_origin(sheet_id, range.start) else {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        };
                        let Some((start, end)) = self.resolver.spill_range(sheet_id, origin) else {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        };

                        EvalValue::Reference(vec![ResolvedRange {
                            sheet_id: FnSheetId::Local(sheet_id),
                            start,
                            end,
                        }])
                    }
                }
            }
            Expr::Unary { op, expr } => {
                let v = self.eval_value(expr);
                let v = self.deref_eval_value_dynamic(v);
                EvalValue::Scalar(elementwise_unary(&v, |elem| numeric_unary(self, *op, elem)))
            }
            Expr::Postfix { op, expr } => match op {
                PostfixOp::Percent => {
                    let v = self.deref_eval_value_dynamic(self.eval_value(expr));
                    EvalValue::Scalar(elementwise_unary(&v, |elem| numeric_percent(self, elem)))
                }
            },
            Expr::Binary { op, left, right } => match *op {
                BinaryOp::Range | BinaryOp::Union | BinaryOp::Intersect => {
                    self.eval_reference_binary(*op, left, right)
                }
                BinaryOp::Concat => {
                    let l = self.deref_eval_value_dynamic(self.eval_value(left));
                    let r = self.deref_eval_value_dynamic(self.eval_value(right));
                    let out = elementwise_binary(&l, &r, |a, b| concat_binary(self, a, b));
                    EvalValue::Scalar(out)
                }
                BinaryOp::Pow | BinaryOp::Add | BinaryOp::Sub | BinaryOp::Mul | BinaryOp::Div => {
                    let l = self.deref_eval_value_dynamic(self.eval_value(left));
                    let r = self.deref_eval_value_dynamic(self.eval_value(right));
                    let out = elementwise_binary(&l, &r, |a, b| numeric_binary(self, *op, a, b));
                    EvalValue::Scalar(out)
                }
            },
            Expr::Compare { op, left, right } => {
                let l = self.deref_eval_value_dynamic(self.eval_value(left));
                let r = self.deref_eval_value_dynamic(self.eval_value(right));
                let out = elementwise_binary(&l, &r, |a, b| excel_compare(a, b, *op));
                EvalValue::Scalar(out)
            }
            Expr::FunctionCall { name, args, .. } => {
                let value = self.eval_function_call(name, args);
                self.function_result_to_eval_value(value)
            }
            Expr::Call { callee, args } => {
                let call_name = match callee.as_ref() {
                    Expr::NameRef(nref) => nref.name.as_str(),
                    _ => ANON_LAMBDA_CALL_NAME,
                };

                let callee_value = match self.eval_value(callee) {
                    EvalValue::Scalar(v) => v,
                    EvalValue::Reference(ranges) => self.deref_reference_scalar(&ranges),
                };

                let value = self.call_value_as_function(call_name, callee_value, args);
                self.function_result_to_eval_value(value)
            }
            Expr::ImplicitIntersection(inner) => {
                let v = self.eval_value(inner);
                match v {
                    EvalValue::Scalar(v) => EvalValue::Scalar(v),
                    EvalValue::Reference(ranges) => {
                        EvalValue::Scalar(self.apply_implicit_intersection(&ranges))
                    }
                }
            }
        }
    }

    fn eval_function_call(&self, name: &str, args: &[CompiledExpr]) -> Value {
        if args.len() > crate::EXCEL_MAX_ARGS {
            return Value::Error(ErrorKind::Value);
        }
        if let Some(spec) = crate::functions::lookup_function(name) {
            if args.len() < spec.min_args || args.len() > spec.max_args {
                return Value::Error(ErrorKind::Value);
            }
            return (spec.implementation)(self, args);
        }

        if let Some(value) = self.lookup_lexical_value(name) {
            return self.call_value_as_function(name, value, args);
        }

        let nref = crate::eval::NameRef {
            sheet: SheetReference::Current,
            name: name.to_string(),
        };
        match self.eval_name_ref(&nref) {
            EvalValue::Scalar(v) => self.call_value_as_function(name, v, args),
            EvalValue::Reference(_) => Value::Error(ErrorKind::Value),
        }
    }

    fn call_value_as_function(
        &self,
        call_name: &str,
        value: Value,
        args: &[CompiledExpr],
    ) -> Value {
        if args.len() > crate::EXCEL_MAX_ARGS {
            return Value::Error(ErrorKind::Value);
        }

        match value {
            Value::Lambda(lambda) => self.call_lambda(call_name, lambda, args),
            Value::Error(e) => Value::Error(e),
            _ => Value::Error(ErrorKind::Value),
        }
    }

    fn call_lambda(
        &self,
        call_name: &str,
        lambda: crate::value::Lambda,
        args: &[CompiledExpr],
    ) -> Value {
        if args.len() > crate::EXCEL_MAX_ARGS {
            return Value::Error(ErrorKind::Value);
        }

        let depth = self.lambda_depth.get();
        if depth >= LAMBDA_RECURSION_LIMIT {
            return Value::Error(ErrorKind::Calc);
        }
        self.lambda_depth.set(depth + 1);

        struct DepthGuard {
            counter: Rc<Cell<u32>>,
        }

        impl Drop for DepthGuard {
            fn drop(&mut self) {
                let depth = self.counter.get();
                self.counter.set(depth.saturating_sub(1));
            }
        }

        let _depth_guard = DepthGuard {
            counter: Rc::clone(&self.lambda_depth),
        };

        if args.len() > lambda.params.len() {
            return Value::Error(ErrorKind::Value);
        }

        let mut evaluated_args = Vec::with_capacity(args.len());
        for arg in args {
            let v = match self.eval_value(arg) {
                EvalValue::Scalar(v) => v,
                EvalValue::Reference(mut ranges) => {
                    self.sort_resolved_ranges(&mut ranges);

                    match ranges.as_slice() {
                        [only] => Value::Reference(FnReference {
                            sheet_id: only.sheet_id.clone(),
                            start: only.start,
                            end: only.end,
                        }),
                        _ => Value::ReferenceUnion(
                            ranges
                                .into_iter()
                                .map(|r| FnReference {
                                    sheet_id: r.sheet_id,
                                    start: r.start,
                                    end: r.end,
                                })
                                .collect(),
                        ),
                    }
                }
            };
            evaluated_args.push(v);
        }

        let mut call_scope =
            HashMap::with_capacity(lambda.params.len().saturating_mul(2).saturating_add(1));
        call_scope.insert(casefold(call_name.trim()), Value::Lambda(lambda.clone()));
        for (idx, param) in lambda.params.iter().enumerate() {
            let value = evaluated_args.get(idx).cloned().unwrap_or(Value::Blank);
            let param_key = casefold(param.trim());
            call_scope.insert(param_key.clone(), value);

            if idx >= args.len() {
                call_scope.insert(
                    format!("{}{}", crate::eval::LAMBDA_OMITTED_PREFIX, param_key),
                    Value::Bool(true),
                );
            }
        }

        let mut scopes = Vec::new();
        if !lambda.env.is_empty() {
            scopes.push((*lambda.env).clone());
        }
        scopes.push(call_scope);

        let evaluator = self.with_lexical_scopes(scopes);
        match evaluator.eval_value(lambda.body.as_ref()) {
            EvalValue::Scalar(v) => v,
            EvalValue::Reference(mut ranges) => {
                // Ensure a stable order for deterministic function behavior (e.g. COUNT over a
                // multi-area union).
                evaluator.sort_resolved_ranges(&mut ranges);

                match ranges.as_slice() {
                    [only] => Value::Reference(FnReference {
                        sheet_id: only.sheet_id.clone(),
                        start: only.start,
                        end: only.end,
                    }),
                    _ => Value::ReferenceUnion(
                        ranges
                            .into_iter()
                            .map(|r| FnReference {
                                sheet_id: r.sheet_id,
                                start: r.start,
                                end: r.end,
                            })
                            .collect(),
                    ),
                }
            }
        }
    }

    fn deref_eval_value_dynamic(&self, value: EvalValue) -> Value {
        match value {
            EvalValue::Scalar(v) => v,
            EvalValue::Reference(range) => self.deref_reference_dynamic(range),
        }
    }

    fn eval_name_ref(&self, nref: &crate::eval::NameRef<usize>) -> EvalValue {
        let Some(sheet_id) = self.resolve_sheet_id(&nref.sheet) else {
            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
        };
        let FnSheetId::Local(sheet_id) = sheet_id else {
            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
        };
        if !self.resolver.sheet_exists(sheet_id) {
            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
        }

        if matches!(nref.sheet, SheetReference::Current) {
            if let Some(value) = self.lookup_lexical_value(&nref.name) {
                match value {
                    Value::Reference(r) => {
                        return EvalValue::Reference(vec![ResolvedRange {
                            sheet_id: r.sheet_id,
                            start: r.start,
                            end: r.end,
                        }]);
                    }
                    Value::ReferenceUnion(ranges) => {
                        return EvalValue::Reference(
                            ranges
                                .into_iter()
                                .map(|r| ResolvedRange {
                                    sheet_id: r.sheet_id,
                                    start: r.start,
                                    end: r.end,
                                })
                                .collect(),
                        );
                    }
                    other => return EvalValue::Scalar(other),
                }
            }
        }

        let Some(def) = self.resolver.resolve_name(sheet_id, &nref.name) else {
            return EvalValue::Scalar(Value::Error(ErrorKind::Name));
        };

        // Prevent infinite recursion from self-referential name chains.
        let key = (sheet_id, casefold(nref.name.trim()));
        {
            let mut stack = self.name_stack.borrow_mut();
            if stack.contains(&key) {
                return EvalValue::Scalar(Value::Error(ErrorKind::Name));
            }
            stack.push(key.clone());
        }

        struct NameGuard {
            stack: Rc<RefCell<Vec<(usize, String)>>>,
            key: (usize, String),
        }

        impl Drop for NameGuard {
            fn drop(&mut self) {
                let mut stack = self.stack.borrow_mut();
                let popped = stack.pop();
                debug_assert_eq!(popped.as_ref(), Some(&self.key));
            }
        }

        let _guard = NameGuard {
            stack: Rc::clone(&self.name_stack),
            key,
        };

        match def {
            ResolvedName::Constant(v) => EvalValue::Scalar(v),
            ResolvedName::Expr(expr) => {
                let evaluator = self.with_ctx(EvalContext {
                    current_sheet: sheet_id,
                    current_cell: self.ctx.current_cell,
                });
                evaluator.eval_value(&expr)
            }
        }
    }

    fn eval_scalar(&self, expr: &CompiledExpr) -> Value {
        match self.eval_value(expr) {
            EvalValue::Scalar(v) => v,
            EvalValue::Reference(ranges) => self.deref_reference_scalar(&ranges),
        }
    }

    fn resolve_sheet_id(&self, sheet: &SheetReference<usize>) -> Option<FnSheetId> {
        match sheet {
            SheetReference::Current => Some(FnSheetId::Local(self.ctx.current_sheet)),
            SheetReference::Sheet(id) => Some(FnSheetId::Local(*id)),
            SheetReference::SheetRange(a, b) => {
                if a == b {
                    Some(FnSheetId::Local(*a))
                } else {
                    None
                }
            }
            SheetReference::External(key) => {
                is_valid_external_single_sheet_key(key).then(|| FnSheetId::External(key.clone()))
            }
        }
    }

    fn resolve_sheet_ids(&self, sheet: &SheetReference<usize>) -> Option<Vec<FnSheetId>> {
        match sheet {
            SheetReference::Current => Some(vec![FnSheetId::Local(self.ctx.current_sheet)]),
            SheetReference::Sheet(id) => Some(vec![FnSheetId::Local(*id)]),
            SheetReference::SheetRange(a, b) => self
                .resolver
                .expand_sheet_span(*a, *b)
                .map(|ids| ids.into_iter().map(FnSheetId::Local).collect()),
            SheetReference::External(key) => {
                if is_valid_external_single_sheet_key(key) {
                    return Some(vec![FnSheetId::External(key.clone())]);
                }

                // External-workbook 3D spans are represented as a single key string (e.g.
                // `"[Book.xlsx]Sheet1:Sheet3"`). Expand these into per-sheet external keys using
                // workbook sheet order supplied by the resolver.
                let (workbook, start, end) = crate::external_refs::parse_external_span_key(key)?;
                let order = self.resolver.workbook_sheet_names(workbook)?;
                let keys = crate::external_refs::expand_external_sheet_span_from_order(
                    workbook, start, end, &order,
                )?;
                Some(keys.into_iter().map(FnSheetId::External).collect())
            }
        }
    }

    fn get_sheet_cell_value(&self, sheet_id: &FnSheetId, addr: CellAddr) -> Value {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.get_cell_value(*id, addr),
            FnSheetId::External(key) => self
                .resolver
                .get_external_value(key, addr)
                .unwrap_or(Value::Error(ErrorKind::Ref)),
        }
    }

    fn addr_in_sheet_bounds(&self, sheet_id: &FnSheetId, addr: CellAddr) -> bool {
        match sheet_id {
            FnSheetId::Local(id) => {
                if !self.resolver.sheet_exists(*id) {
                    return false;
                }
                let (rows, cols) = self.resolver.sheet_dimensions(*id);
                addr.row < rows && addr.col < cols
            }
            // External workbooks do not expose dimensions via the ValueResolver interface, so
            // treat bounds as unknown (and therefore valid) and rely on `get_external_value` to
            // surface `#REF!` when the reference cannot be resolved.
            FnSheetId::External(_) => true,
        }
    }

    fn reference_endpoints_in_bounds(
        &self,
        sheet_id: &FnSheetId,
        start: CellAddr,
        end: CellAddr,
    ) -> bool {
        self.addr_in_sheet_bounds(sheet_id, start) && self.addr_in_sheet_bounds(sheet_id, end)
    }

    fn deref_reference_scalar(&self, ranges: &[ResolvedRange]) -> Value {
        match ranges {
            [only] => {
                if !self.reference_endpoints_in_bounds(&only.sheet_id, only.start, only.end) {
                    return Value::Error(ErrorKind::Ref);
                }
                if only.is_single_cell() {
                    self.trace_cell(&only.sheet_id, only.start);
                    return self.get_sheet_cell_value(&only.sheet_id, only.start);
                }
                // Multi-cell references used as scalars behave like a spill attempt.
                Value::Error(ErrorKind::Spill)
            }
            _ => Value::Error(ErrorKind::Value),
        }
    }

    fn deref_reference_dynamic(&self, ranges: Vec<ResolvedRange>) -> Value {
        match ranges.as_slice() {
            [] => Value::Error(ErrorKind::Ref),
            [only] => self.deref_reference_dynamic_single(only),
            // Discontiguous unions cannot be represented as a single rectangular spill.
            _ => Value::Error(ErrorKind::Value),
        }
    }

    fn deref_reference_dynamic_single(&self, range: &ResolvedRange) -> Value {
        if !self.reference_endpoints_in_bounds(&range.sheet_id, range.start, range.end) {
            return Value::Error(ErrorKind::Ref);
        }
        if range.is_single_cell() {
            self.trace_cell(&range.sheet_id, range.start);
            return self.get_sheet_cell_value(&range.sheet_id, range.start);
        }
        let range = range.normalized();
        let reference = FnReference {
            sheet_id: range.sheet_id.clone(),
            start: range.start,
            end: range.end,
        };
        self.trace_reference(&reference);
        let rows = (range.end.row - range.start.row + 1) as usize;
        let cols = (range.end.col - range.start.col + 1) as usize;

        let total_cells = match rows.checked_mul(cols) {
            Some(v) => v,
            None => return Value::Error(ErrorKind::Spill),
        };
        if total_cells > MAX_MATERIALIZED_ARRAY_CELLS {
            return Value::Error(ErrorKind::Spill);
        }

        let mut values: Vec<Value> = Vec::new();
        if values.try_reserve_exact(total_cells).is_err() {
            return Value::Error(ErrorKind::Num);
        }
        for row in range.start.row..=range.end.row {
            for col in range.start.col..=range.end.col {
                values.push(self.get_sheet_cell_value(&range.sheet_id, CellAddr { row, col }));
            }
        }
        Value::Array(Array::new(rows, cols, values))
    }

    fn apply_implicit_intersection(&self, ranges: &[ResolvedRange]) -> Value {
        match ranges {
            [] => Value::Error(ErrorKind::Value),
            [only] => self.apply_implicit_intersection_single(only),
            many => {
                // If multiple areas intersect, Excel's implicit intersection is ambiguous. We
                // approximate by succeeding only when exactly one area intersects.
                let mut hits = Vec::new();
                for r in many {
                    let v = self.apply_implicit_intersection_single(r);
                    if !matches!(v, Value::Error(ErrorKind::Value)) {
                        hits.push(v);
                    }
                }
                match hits.as_slice() {
                    [only] => only.clone(),
                    _ => Value::Error(ErrorKind::Value),
                }
            }
        }
    }

    fn apply_implicit_intersection_single(&self, range: &ResolvedRange) -> Value {
        if !self.reference_endpoints_in_bounds(&range.sheet_id, range.start, range.end) {
            return Value::Error(ErrorKind::Ref);
        }
        if range.is_single_cell() {
            self.trace_cell(&range.sheet_id, range.start);
            return self.get_sheet_cell_value(&range.sheet_id, range.start);
        }

        let range = range.normalized();
        let cur = self.ctx.current_cell;

        // 1D ranges intersect on the matching row/column.
        if range.start.col == range.end.col {
            if cur.row >= range.start.row && cur.row <= range.end.row {
                let addr = CellAddr {
                    row: cur.row,
                    col: range.start.col,
                };
                self.trace_cell(&range.sheet_id, addr);
                return self.get_sheet_cell_value(&range.sheet_id, addr);
            }
            return Value::Error(ErrorKind::Value);
        }
        if range.start.row == range.end.row {
            if cur.col >= range.start.col && cur.col <= range.end.col {
                let addr = CellAddr {
                    row: range.start.row,
                    col: cur.col,
                };
                self.trace_cell(&range.sheet_id, addr);
                return self.get_sheet_cell_value(&range.sheet_id, addr);
            }
            return Value::Error(ErrorKind::Value);
        }

        // 2D ranges intersect only if the current cell is within the rectangle.
        if cur.row >= range.start.row
            && cur.row <= range.end.row
            && cur.col >= range.start.col
            && cur.col <= range.end.col
        {
            self.trace_cell(&range.sheet_id, cur);
            return self.get_sheet_cell_value(&range.sheet_id, cur);
        }

        Value::Error(ErrorKind::Value)
    }

    fn eval_reference_binary(
        &self,
        op: BinaryOp,
        left: &CompiledExpr,
        right: &CompiledExpr,
    ) -> EvalValue {
        let left = match self.eval_reference_operand(left) {
            Ok(r) => r,
            Err(v) => return EvalValue::Scalar(v),
        };
        let right = match self.eval_reference_operand(right) {
            Ok(r) => r,
            Err(v) => return EvalValue::Scalar(v),
        };

        match op {
            BinaryOp::Union => {
                let Some(sheet_id) = left.first().map(|r| &r.sheet_id) else {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                };
                if left.iter().any(|r| &r.sheet_id != sheet_id)
                    || right.iter().any(|r| &r.sheet_id != sheet_id)
                {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                }

                let mut out = left;
                out.extend(right);
                EvalValue::Reference(out)
            }
            BinaryOp::Intersect => {
                let mut out = Vec::new();
                for a in &left {
                    for b in &right {
                        if a.sheet_id != b.sheet_id {
                            return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                        }
                        if let Some(r) = intersect_ranges(a, b) {
                            out.push(r);
                        }
                    }
                }
                if out.is_empty() {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Null));
                }
                EvalValue::Reference(out)
            }
            BinaryOp::Range => {
                let (Some(a), Some(b)) = (left.first(), right.first()) else {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                };
                if left.len() != 1 || right.len() != 1 {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Value));
                }
                if a.sheet_id != b.sheet_id {
                    return EvalValue::Scalar(Value::Error(ErrorKind::Ref));
                }

                let a = a.normalized();
                let b = b.normalized();

                let start = CellAddr {
                    row: a.start.row.min(b.start.row),
                    col: a.start.col.min(b.start.col),
                };
                let end = CellAddr {
                    row: a.end.row.max(b.end.row),
                    col: a.end.col.max(b.end.col),
                };

                EvalValue::Reference(vec![ResolvedRange {
                    sheet_id: a.sheet_id,
                    start,
                    end,
                }])
            }
            _ => EvalValue::Scalar(Value::Error(ErrorKind::Value)),
        }
    }

    fn eval_reference_operand(&self, expr: &CompiledExpr) -> Result<Vec<ResolvedRange>, Value> {
        match self.eval_value(expr) {
            EvalValue::Reference(r) => Ok(r),
            EvalValue::Scalar(Value::Error(e)) => Err(Value::Error(e)),
            EvalValue::Scalar(_) => Err(Value::Error(ErrorKind::Value)),
        }
    }

    // Built-in functions are implemented in `crate::functions` and dispatched via
    // `crate::functions::call_function`.
}

fn intersect_ranges(a: &ResolvedRange, b: &ResolvedRange) -> Option<ResolvedRange> {
    if a.sheet_id != b.sheet_id {
        return None;
    }
    let a = a.normalized();
    let b = b.normalized();

    let start_row = a.start.row.max(b.start.row);
    let end_row = a.end.row.min(b.end.row);
    if start_row > end_row {
        return None;
    }
    let start_col = a.start.col.max(b.start.col);
    let end_col = a.end.col.min(b.end.col);
    if start_col > end_col {
        return None;
    }

    Some(ResolvedRange {
        sheet_id: a.sheet_id,
        start: CellAddr {
            row: start_row,
            col: start_col,
        },
        end: CellAddr {
            row: end_row,
            col: end_col,
        },
    })
}

pub(crate) fn is_valid_external_single_sheet_key(key: &str) -> bool {
    crate::external_refs::parse_external_key(key).is_some()
}

pub(crate) fn split_external_sheet_key_parts(key: &str) -> Option<(&str, &str)> {
    crate::external_refs::split_external_sheet_key_parts(key)
}

pub(crate) fn split_external_sheet_span_key(key: &str) -> Option<(&str, &str, &str)> {
    crate::external_refs::parse_external_span_key(key)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn split_external_sheet_key_parts_parses_workbook_and_sheet() {
        let key = "[Book.xlsx]Sheet1";
        let (workbook, sheet) = split_external_sheet_key_parts(key).unwrap();
        assert_eq!(workbook, "Book.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn split_external_sheet_span_key_parses_workbook_and_sheet_span() {
        let key = "[Book.xlsx]Sheet1:Sheet3";
        let (workbook, start, end) = split_external_sheet_span_key(key).unwrap();
        assert_eq!(workbook, "Book.xlsx");
        assert_eq!(start, "Sheet1");
        assert_eq!(end, "Sheet3");
    }

    #[test]
    fn split_external_sheet_span_key_uses_last_closing_bracket_for_workbook_id() {
        let key = "[C:\\[foo]\\Book.xlsx]Sheet1:Sheet3";
        let (workbook, start, end) = split_external_sheet_span_key(key).unwrap();
        assert_eq!(workbook, "C:\\[foo]\\Book.xlsx");
        assert_eq!(start, "Sheet1");
        assert_eq!(end, "Sheet3");
    }

    #[test]
    fn split_external_sheet_key_parts_uses_last_closing_bracket_for_workbook_id() {
        // Workbook ids can contain `[` / `]` in a path prefix, so we must locate the *last* `]`.
        let key = "[C:\\[foo]\\Book.xlsx]Sheet1";
        let (workbook, sheet) = split_external_sheet_key_parts(key).unwrap();
        assert_eq!(workbook, "C:\\[foo]\\Book.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn split_external_sheet_key_parts_rejects_invalid_inputs() {
        for key in [
            "Book.xlsx]Sheet1", // missing leading '['
            "[Book.xlsxSheet1", // missing closing ']'
            "Sheet1",           // missing workbook prefix
            "[]Sheet1",         // empty workbook
            "[Book.xlsx]",      // empty sheet
        ] {
            assert!(
                split_external_sheet_key_parts(key).is_none(),
                "expected None for key {key:?}"
            );
        }
    }

    #[test]
    fn split_external_sheet_span_key_rejects_missing_endpoints() {
        for key in [
            "[Book.xlsx]Sheet1",
            "[Book.xlsx]Sheet1:",
            "[Book.xlsx]:Sheet2",
        ] {
            assert!(
                split_external_sheet_span_key(key).is_none(),
                "expected None for key {key:?}"
            );
        }
    }

    #[test]
    fn is_valid_external_single_sheet_key_accepts_single_sheet_rejects_span() {
        assert!(is_valid_external_single_sheet_key("[Book.xlsx]Sheet1"));
        assert!(!is_valid_external_single_sheet_key(
            "[Book.xlsx]Sheet1:Sheet3"
        ));
    }
}

impl<'a, R: ValueResolver> FunctionContext for Evaluator<'a, R> {
    fn eval_arg(&self, expr: &CompiledExpr) -> FnArgValue {
        match self.eval_value(expr) {
            EvalValue::Scalar(v) => FnArgValue::Scalar(v),
            EvalValue::Reference(mut ranges) => {
                // Ensure a stable order for deterministic function behavior (e.g. COUNT over a
                // multi-area union).
                self.sort_resolved_ranges(&mut ranges);
                match ranges.as_slice() {
                    [only] => FnArgValue::Reference(FnReference {
                        sheet_id: only.sheet_id.clone(),
                        start: only.start,
                        end: only.end,
                    }),
                    _ => FnArgValue::ReferenceUnion(
                        ranges
                            .into_iter()
                            .map(|r| FnReference {
                                sheet_id: r.sheet_id,
                                start: r.start,
                                end: r.end,
                            })
                            .collect(),
                    ),
                }
            }
        }
    }

    fn eval_scalar(&self, expr: &CompiledExpr) -> Value {
        Evaluator::eval_scalar(self, expr)
    }

    fn eval_formula(&self, expr: &CompiledExpr) -> Value {
        Evaluator::eval_formula(self, expr)
    }

    fn eval_formula_with_bindings(
        &self,
        expr: &CompiledExpr,
        bindings: &HashMap<String, Value>,
    ) -> Value {
        if bindings.is_empty() {
            return self.eval_formula(expr);
        }

        let mut scope = HashMap::with_capacity(bindings.len());
        for (k, v) in bindings {
            scope.insert(casefold(k.trim()), v.clone());
        }
        let _guard = self.push_lexical_scope(scope);
        self.eval_formula(expr)
    }

    fn capture_lexical_env(&self) -> HashMap<String, Value> {
        self.capture_lexical_env_map()
    }

    fn apply_implicit_intersection(&self, reference: &FnReference) -> Value {
        Evaluator::apply_implicit_intersection(
            self,
            &[ResolvedRange {
                sheet_id: reference.sheet_id.clone(),
                start: reference.start,
                end: reference.end,
            }],
        )
    }

    fn get_cell_value(&self, sheet_id: &FnSheetId, addr: CellAddr) -> Value {
        self.get_sheet_cell_value(sheet_id, addr)
    }

    fn get_cell_phonetic(&self, sheet_id: &FnSheetId, addr: CellAddr) -> Option<&str> {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.get_cell_phonetic(*id, addr),
            FnSheetId::External(_) => None,
        }
    }

    fn cell_horizontal_alignment(
        &self,
        sheet_id: &FnSheetId,
        addr: CellAddr,
    ) -> Option<HorizontalAlignment> {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.cell_horizontal_alignment(*id, addr),
            FnSheetId::External(_) => None,
        }
    }

    fn sheet_dimensions(&self, sheet_id: &FnSheetId) -> (u32, u32) {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.sheet_dimensions(*id),
            FnSheetId::External(_) => {
                (formula_model::EXCEL_MAX_ROWS, formula_model::EXCEL_MAX_COLS)
            }
        }
    }

    fn sheet_default_col_width(&self, sheet_id: &FnSheetId) -> Option<f32> {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.sheet_default_col_width(*id),
            FnSheetId::External(_) => None,
        }
    }
    fn sheet_origin_cell(&self, sheet_id: usize) -> Option<CellAddr> {
        self.resolver.sheet_origin_cell(sheet_id)
    }
    fn iter_reference_cells<'b>(
        &'b self,
        reference: &'b FnReference,
    ) -> Box<dyn Iterator<Item = CellAddr> + 'b> {
        self.trace_reference(reference);
        match &reference.sheet_id {
            FnSheetId::Local(sheet_id) => {
                if let Some(iter) = self.resolver.iter_sheet_cells(*sheet_id) {
                    Box::new(iter.filter(move |addr| reference.contains(*addr)))
                } else {
                    Box::new(reference.iter_cells())
                }
            }
            FnSheetId::External(_) => Box::new(reference.iter_cells()),
        }
    }

    fn record_reference(&self, reference: &FnReference) {
        self.trace_reference(reference);
    }

    fn locale_config(&self) -> LocaleConfig {
        self.locale.clone()
    }

    fn now_utc(&self) -> chrono::DateTime<chrono::Utc> {
        self.recalc_ctx.now_utc
    }

    fn calculation_mode(&self) -> CalculationMode {
        self.recalc_ctx.calculation_mode
    }

    fn push_local_scope(&self) {
        self.lexical_scopes.borrow_mut().push(HashMap::new());
    }

    fn pop_local_scope(&self) {
        self.lexical_scopes.borrow_mut().pop();
    }

    fn set_local(&self, name: &str, value: FnArgValue) {
        let key = casefold(name.trim());
        let mut scopes = self.lexical_scopes.borrow_mut();
        if scopes.is_empty() {
            scopes.push(HashMap::new());
        }
        if let Some(scope) = scopes.last_mut() {
            let value = match value {
                FnArgValue::Scalar(v) => v,
                FnArgValue::Reference(r) => Value::Reference(r),
                FnArgValue::ReferenceUnion(ranges) => Value::ReferenceUnion(ranges),
            };
            scope.insert(key, value);
        }
    }

    fn make_lambda(&self, params: Vec<String>, body: CompiledExpr) -> Value {
        let params: Vec<String> = params.into_iter().map(|p| casefold(p.trim())).collect();

        let mut env = self.capture_lexical_env_map();
        env.retain(|k, _| !k.starts_with(crate::eval::LAMBDA_OMITTED_PREFIX));

        Value::Lambda(Lambda {
            params: params.into(),
            body: Arc::new(body),
            env: Arc::new(env),
        })
    }

    fn eval_lambda(&self, lambda: &Lambda, args: Vec<FnArgValue>) -> Value {
        if args.len() > crate::EXCEL_MAX_ARGS {
            return Value::Error(ErrorKind::Value);
        }

        if args.len() > lambda.params.len() {
            return Value::Error(ErrorKind::Value);
        }

        let depth = self.lambda_depth.get();
        if depth >= LAMBDA_RECURSION_LIMIT {
            return Value::Error(ErrorKind::Calc);
        }
        self.lambda_depth.set(depth + 1);

        struct DepthGuard {
            counter: Rc<Cell<u32>>,
        }

        impl Drop for DepthGuard {
            fn drop(&mut self) {
                let depth = self.counter.get();
                self.counter.set(depth.saturating_sub(1));
            }
        }

        let _depth_guard = DepthGuard {
            counter: Rc::clone(&self.lambda_depth),
        };

        let mut call_scope =
            HashMap::with_capacity(lambda.params.len().saturating_mul(2).saturating_add(1));
        call_scope.insert(
            ANON_LAMBDA_CALL_NAME.to_string(),
            Value::Lambda(lambda.clone()),
        );

        for (idx, param) in lambda.params.iter().enumerate() {
            let value = args
                .get(idx)
                .cloned()
                .unwrap_or(FnArgValue::Scalar(Value::Blank));
            let value = match value {
                FnArgValue::Scalar(v) => v,
                FnArgValue::Reference(r) => Value::Reference(r),
                FnArgValue::ReferenceUnion(ranges) => Value::ReferenceUnion(ranges),
            };
            let param_key = casefold(param.trim());
            call_scope.insert(param_key.clone(), value);

            if idx >= args.len() {
                call_scope.insert(
                    format!("{}{}", crate::eval::LAMBDA_OMITTED_PREFIX, param_key),
                    Value::Bool(true),
                );
            }
        }

        let mut scopes = Vec::new();
        if !lambda.env.is_empty() {
            scopes.push((*lambda.env).clone());
        }
        scopes.push(call_scope);

        let evaluator = self.with_lexical_scopes(scopes);
        match evaluator.eval_value(lambda.body.as_ref()) {
            EvalValue::Scalar(v) => v,
            EvalValue::Reference(mut ranges) => {
                // Ensure a stable order for deterministic function behavior (e.g. COUNT over a
                // multi-area union).
                evaluator.sort_resolved_ranges(&mut ranges);

                match ranges.as_slice() {
                    [only] => Value::Reference(FnReference {
                        sheet_id: only.sheet_id.clone(),
                        start: only.start,
                        end: only.end,
                    }),
                    _ => Value::ReferenceUnion(
                        ranges
                            .into_iter()
                            .map(|r| FnReference {
                                sheet_id: r.sheet_id,
                                start: r.start,
                                end: r.end,
                            })
                            .collect(),
                    ),
                }
            }
        }
    }

    fn volatile_rand_u64(&self) -> u64 {
        let draw = self.rng_counter.get();
        self.rng_counter.set(draw.wrapping_add(1));

        let mut seed = self.recalc_ctx.recalc_id;
        seed ^= (self.ctx.current_sheet as u64).wrapping_mul(0x9e3779b97f4a7c15);
        seed ^= (self.ctx.current_cell.row as u64).wrapping_mul(0xbf58476d1ce4e5b9);
        seed ^= (self.ctx.current_cell.col as u64).wrapping_mul(0x94d049bb133111eb);
        seed ^= draw.wrapping_mul(0x3c79ac492ba7b653);
        splitmix64(seed)
    }

    fn date_system(&self) -> ExcelDateSystem {
        self.date_system
    }

    fn current_sheet_id(&self) -> usize {
        self.ctx.current_sheet
    }

    fn current_cell_addr(&self) -> CellAddr {
        self.ctx.current_cell
    }

    fn sheet_order_index(&self, sheet_id: usize) -> Option<usize> {
        self.resolver.sheet_order_index(sheet_id)
    }

    fn sheet_name(&self, sheet_id: usize) -> Option<&str> {
        self.resolver.sheet_name(sheet_id)
    }

    fn sheet_count(&self) -> usize {
        self.resolver.sheet_count()
    }

    fn info_system(&self) -> Option<&str> {
        self.resolver.info_system()
    }

    fn info_directory(&self) -> Option<&str> {
        self.resolver.info_directory()
    }

    fn info_osversion(&self) -> Option<&str> {
        self.resolver.info_osversion()
    }

    fn info_release(&self) -> Option<&str> {
        self.resolver.info_release()
    }

    fn info_version(&self) -> Option<&str> {
        self.resolver.info_version()
    }

    fn info_memavail(&self) -> Option<f64> {
        self.resolver.info_memavail()
    }

    fn info_totmem(&self) -> Option<f64> {
        self.resolver.info_totmem()
    }

    fn info_origin(&self) -> Option<&str> {
        self.resolver.info_origin(self.ctx.current_sheet)
    }

    fn get_cell_formula(&self, sheet_id: &FnSheetId, addr: CellAddr) -> Option<&str> {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.get_cell_formula(*id, addr),
            FnSheetId::External(_) => None,
        }
    }

    fn style_table(&self) -> Option<&formula_model::StyleTable> {
        self.resolver.style_table()
    }

    fn sheet_default_style_id(&self, sheet_id: &FnSheetId) -> Option<u32> {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.sheet_default_style_id(*id),
            FnSheetId::External(_) => None,
        }
    }

    fn cell_style_id(&self, sheet_id: &FnSheetId, addr: CellAddr) -> u32 {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.cell_style_id(*id, addr),
            FnSheetId::External(_) => 0,
        }
    }

    fn format_run_style_id(&self, sheet_id: &FnSheetId, addr: CellAddr) -> u32 {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.format_run_style_id(*id, addr),
            FnSheetId::External(_) => 0,
        }
    }

    fn row_style_id(&self, sheet_id: &FnSheetId, row: u32) -> Option<u32> {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.row_style_id(*id, row),
            FnSheetId::External(_) => None,
        }
    }

    fn get_cell_number_format(&self, sheet_id: &FnSheetId, addr: CellAddr) -> Option<&str> {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.get_cell_number_format(*id, addr),
            FnSheetId::External(_) => None,
        }
    }

    fn col_properties(
        &self,
        sheet_id: &FnSheetId,
        col: u32,
    ) -> Option<formula_model::ColProperties> {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.col_properties(*id, col),
            FnSheetId::External(_) => None,
        }
    }

    fn range_run_style_id(&self, sheet_id: &FnSheetId, addr: CellAddr) -> u32 {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.range_run_style_id(*id, addr),
            FnSheetId::External(_) => 0,
        }
    }

    fn workbook_directory(&self) -> Option<&str> {
        self.resolver.workbook_directory()
    }

    fn workbook_filename(&self) -> Option<&str> {
        self.resolver.workbook_filename()
    }
    fn pivot_registry(&self) -> Option<&crate::pivot_registry::PivotRegistry> {
        self.resolver.pivot_registry()
    }

    fn effective_cell_style(
        &self,
        sheet_id: &FnSheetId,
        addr: CellAddr,
    ) -> crate::style_patch::EffectiveStyle {
        match sheet_id {
            FnSheetId::Local(id) => self.resolver.effective_cell_style(*id, addr),
            FnSheetId::External(_) => crate::style_patch::EffectiveStyle::default(),
        }
    }
    fn resolve_sheet_name(&self, name: &str) -> Option<usize> {
        self.resolver.sheet_id(name)
    }

    fn external_sheet_order(&self, workbook: &str) -> Option<Vec<String>> {
        self.resolver.external_sheet_order(workbook)
    }

    fn workbook_sheet_names(&self, workbook: &str) -> Option<Arc<[String]>> {
        self.resolver.workbook_sheet_names(workbook)
    }

    fn external_data_provider(&self) -> Option<&dyn crate::ExternalDataProvider> {
        self.resolver.external_data_provider()
    }

    fn number_locale(&self) -> NumberLocale {
        self.recalc_ctx.number_locale
    }

    fn value_locale(&self) -> ValueLocaleConfig {
        self.value_locale
    }

    fn text_codepage(&self) -> u16 {
        self.text_codepage
    }
}

fn excel_compare(left: &Value, right: &Value, op: CompareOp) -> Value {
    let ord = match excel_order(left, right) {
        Ok(ord) => ord,
        Err(e) => return Value::Error(e),
    };

    let result = match op {
        CompareOp::Eq => ord == Ordering::Equal,
        CompareOp::Ne => ord != Ordering::Equal,
        CompareOp::Lt => ord == Ordering::Less,
        CompareOp::Le => ord != Ordering::Greater,
        CompareOp::Gt => ord == Ordering::Greater,
        CompareOp::Ge => ord != Ordering::Less,
    };

    Value::Bool(result)
}

fn excel_order(left: &Value, right: &Value) -> Result<Ordering, ErrorKind> {
    if let Value::Error(e) = left {
        return Err(*e);
    }
    if let Value::Error(e) = right {
        return Err(*e);
    }

    // Treat rich values as text for comparison semantics.
    let left = match left.clone() {
        Value::Entity(v) => Value::Text(v.display),
        Value::Record(v) => Value::Text(v.display),
        other => other,
    };
    let right = match right.clone() {
        Value::Entity(v) => Value::Text(v.display),
        Value::Record(v) => Value::Text(v.display),
        other => other,
    };
    if matches!(
        &left,
        Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. }
            | Value::Record(_)
            | Value::Entity(_)
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
    ) || matches!(
        &right,
        Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. }
            | Value::Record(_)
            | Value::Entity(_)
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
    ) {
        return Err(ErrorKind::Value);
    }

    // Blank coerces to the other type for comparisons.
    let (l, r) = match (left, right) {
        (Value::Blank, Value::Number(b)) => (Value::Number(0.0), Value::Number(b)),
        (Value::Number(a), Value::Blank) => (Value::Number(a), Value::Number(0.0)),
        (Value::Blank, Value::Bool(b)) => (Value::Bool(false), Value::Bool(b)),
        (Value::Bool(a), Value::Blank) => (Value::Bool(a), Value::Bool(false)),
        (Value::Blank, Value::Text(b)) => (Value::Text(String::new()), Value::Text(b)),
        (Value::Text(a), Value::Blank) => (Value::Text(a), Value::Text(String::new())),
        (l, r) => (l, r),
    };

    fn text_like_str(v: &Value) -> Option<&str> {
        match v {
            Value::Text(s) => Some(s),
            _ => None,
        }
    }

    Ok(match (&l, &r) {
        (Value::Number(a), Value::Number(b)) => a.partial_cmp(b).unwrap_or(Ordering::Equal),
        (a, b) if text_like_str(a).is_some() && text_like_str(b).is_some() => {
            cmp_case_insensitive(text_like_str(a).unwrap(), text_like_str(b).unwrap())
        }
        (Value::Bool(a), Value::Bool(b)) => a.cmp(b),
        // Type precedence (approximate Excel): numbers < text < booleans.
        (
            Value::Number(_),
            Value::Text(_) | Value::Entity(_) | Value::Record(_) | Value::Bool(_),
        ) => Ordering::Less,
        (Value::Text(_) | Value::Entity(_) | Value::Record(_), Value::Bool(_)) => Ordering::Less,
        (Value::Text(_) | Value::Entity(_) | Value::Record(_), Value::Number(_)) => {
            Ordering::Greater
        }
        (
            Value::Bool(_),
            Value::Number(_) | Value::Text(_) | Value::Entity(_) | Value::Record(_),
        ) => Ordering::Greater,
        // Blank should have been coerced above.
        (Value::Blank, Value::Blank) => Ordering::Equal,
        (Value::Blank, _) => Ordering::Less,
        (_, Value::Blank) => Ordering::Greater,
        // Errors are handled above.
        (Value::Error(_), _) | (_, Value::Error(_)) => Ordering::Equal,
        (Value::Entity(_), _)
        | (_, Value::Entity(_))
        | (Value::Record(_), _)
        | (_, Value::Record(_)) => Ordering::Equal,
        // Arrays/spill markers/lambdas/references are rejected above.
        (Value::Array(_), _)
        | (_, Value::Array(_))
        | (Value::Lambda(_), _)
        | (_, Value::Lambda(_))
        | (Value::Spill { .. }, _)
        | (_, Value::Spill { .. })
        | (Value::Reference(_), _)
        | (_, Value::Reference(_))
        | (Value::ReferenceUnion(_), _)
        | (_, Value::ReferenceUnion(_)) => Ordering::Equal,
        _ => Ordering::Equal,
    })
}

fn numeric_unary(ctx: &dyn FunctionContext, op: UnaryOp, value: &Value) -> Value {
    match value {
        Value::Error(e) => return Value::Error(*e),
        other => {
            let n = match other.coerce_to_number_with_ctx(ctx) {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            let out = match op {
                UnaryOp::Plus => n,
                UnaryOp::Minus => -n,
            };
            Value::Number(out)
        }
    }
}

fn numeric_percent(ctx: &dyn FunctionContext, value: &Value) -> Value {
    match value {
        Value::Error(e) => return Value::Error(*e),
        other => {
            let n = match other.coerce_to_number_with_ctx(ctx) {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            Value::Number(n / 100.0)
        }
    }
}

fn concat_binary(ctx: &dyn FunctionContext, left: &Value, right: &Value) -> Value {
    if let Value::Error(e) = left {
        return Value::Error(*e);
    }
    if let Value::Error(e) = right {
        return Value::Error(*e);
    }

    let ls = match left.coerce_to_string_with_ctx(ctx) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let rs = match right.coerce_to_string_with_ctx(ctx) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    Value::Text(format!("{ls}{rs}"))
}

fn numeric_binary(ctx: &dyn FunctionContext, op: BinaryOp, left: &Value, right: &Value) -> Value {
    if let Value::Error(e) = left {
        return Value::Error(*e);
    }
    if let Value::Error(e) = right {
        return Value::Error(*e);
    }

    let ln = match left.coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let rn = match right.coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    match op {
        BinaryOp::Add => Value::Number(ln + rn),
        BinaryOp::Sub => Value::Number(ln - rn),
        BinaryOp::Mul => Value::Number(ln * rn),
        BinaryOp::Div => {
            if rn == 0.0 {
                Value::Error(ErrorKind::Div0)
            } else {
                Value::Number(ln / rn)
            }
        }
        BinaryOp::Pow => match crate::functions::math::power(ln, rn) {
            Ok(n) => Value::Number(n),
            Err(e) => Value::Error(match e {
                ExcelError::Div0 => ErrorKind::Div0,
                ExcelError::Value => ErrorKind::Value,
                ExcelError::Num => ErrorKind::Num,
            }),
        },
        _ => Value::Error(ErrorKind::Value),
    }
}

fn elementwise_unary(value: &Value, f: impl Fn(&Value) -> Value) -> Value {
    match value {
        Value::Array(arr) => {
            Value::Array(Array::new(arr.rows, arr.cols, arr.iter().map(f).collect()))
        }
        other => f(other),
    }
}

fn eval_field_access(value: &Value, field_key: &str) -> Value {
    match value {
        Value::Error(e) => Value::Error(*e),
        Value::Entity(entity) => entity
            .get_field_case_insensitive(field_key)
            .unwrap_or(Value::Error(ErrorKind::Field)),
        Value::Record(record) => record
            .get_field_case_insensitive(field_key)
            .unwrap_or(Value::Error(ErrorKind::Field)),
        // Field access on a non-rich value yields `#VALUE!` (wrong argument type). `#FIELD!` is
        // reserved for missing fields on rich values.
        _ => Value::Error(ErrorKind::Value),
    }
}

fn elementwise_binary(left: &Value, right: &Value, f: impl Fn(&Value, &Value) -> Value) -> Value {
    match (left, right) {
        (Value::Array(left_arr), Value::Array(right_arr)) => {
            let out_rows = if left_arr.rows == right_arr.rows {
                left_arr.rows
            } else if left_arr.rows == 1 {
                right_arr.rows
            } else if right_arr.rows == 1 {
                left_arr.rows
            } else {
                return Value::Error(ErrorKind::Value);
            };

            let out_cols = if left_arr.cols == right_arr.cols {
                left_arr.cols
            } else if left_arr.cols == 1 {
                right_arr.cols
            } else if right_arr.cols == 1 {
                left_arr.cols
            } else {
                return Value::Error(ErrorKind::Value);
            };

            let mut out = Vec::with_capacity(out_rows.saturating_mul(out_cols));
            for row in 0..out_rows {
                let l_row = if left_arr.rows == 1 { 0 } else { row };
                let r_row = if right_arr.rows == 1 { 0 } else { row };
                for col in 0..out_cols {
                    let l_col = if left_arr.cols == 1 { 0 } else { col };
                    let r_col = if right_arr.cols == 1 { 0 } else { col };
                    let l = left_arr.get(l_row, l_col).unwrap_or(&Value::Blank);
                    let r = right_arr.get(r_row, r_col).unwrap_or(&Value::Blank);
                    out.push(f(l, r));
                }
            }

            Value::Array(Array::new(out_rows, out_cols, out))
        }
        (Value::Array(left_arr), right_scalar) => Value::Array(Array::new(
            left_arr.rows,
            left_arr.cols,
            left_arr.values.iter().map(|a| f(a, right_scalar)).collect(),
        )),
        (left_scalar, Value::Array(right_arr)) => Value::Array(Array::new(
            right_arr.rows,
            right_arr.cols,
            right_arr.values.iter().map(|b| f(left_scalar, b)).collect(),
        )),
        (left_scalar, right_scalar) => f(left_scalar, right_scalar),
    }
}

fn splitmix64(mut state: u64) -> u64 {
    // A simple, fast mixer with good statistical properties (used as a deterministic
    // PRNG building block). The transform is bijective over u64, making it a good fit
    // for per-cell deterministic RNG.
    state = state.wrapping_add(0x9e3779b97f4a7c15);
    state = (state ^ (state >> 30)).wrapping_mul(0xbf58476d1ce4e5b9);
    state = (state ^ (state >> 27)).wrapping_mul(0x94d049bb133111eb);
    state ^ (state >> 31)
}
