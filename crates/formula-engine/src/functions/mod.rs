use std::collections::HashMap;
use std::sync::{Arc, OnceLock};

use crate::date::ExcelDateSystem;
use crate::eval::{CellAddr, CompiledExpr};
use crate::locale::ValueLocaleConfig;
use crate::value::{ErrorKind, Lambda, Value};
use crate::LocaleConfig;
use formula_model::HorizontalAlignment;
use formula_model::{EXCEL_MAX_COLS, EXCEL_MAX_ROWS};

pub(crate) mod array_lift;
pub mod database;
pub mod date_time;
pub mod engineering;
pub mod financial;
pub mod information;
pub mod lookup;
pub mod math;
pub mod statistical;
pub mod text;
pub(crate) mod wildcard;

/// Identifies the source worksheet for a reference argument.
///
/// Most references target a worksheet inside the current workbook (`Local`). Excel formulas can
/// also reference cells/ranges in external workbooks, e.g. `=[Book.xlsx]Sheet1!A1`.
///
/// External sheet keys are preserved in the canonical bracketed form produced by the parser
/// (`"[Book.xlsx]Sheet1"`), which avoids allocating synthetic local sheet ids.
#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum SheetId {
    Local(usize),
    External(String),
}

// Built-in Excel-compatible functions registered with the inventory-backed
// registry live in dedicated modules to avoid merge conflicts.
mod builtins_array;
mod builtins_cube;
mod builtins_database;
mod builtins_date_time;
mod builtins_dynamic_array_textsplit;
mod builtins_dynamic_arrays;
mod builtins_engineering;
mod builtins_engineering_complex;
mod builtins_engineering_convert;
mod builtins_engineering_special;
mod builtins_image;
mod builtins_information;
mod builtins_information_workbook;
mod builtins_information_worksheet;
mod builtins_lambda;
mod builtins_logical;
mod builtins_logical_constants;
mod builtins_logical_extended;
mod builtins_lookup;
mod builtins_math;
mod builtins_math_extended;
mod builtins_math_matrix;
mod builtins_math_more;
mod builtins_reference;
mod builtins_rich_values;
mod builtins_roman;
mod builtins_select;
mod builtins_statistical;
mod builtins_statistical_distributions;
mod builtins_statistical_ets;
mod builtins_statistical_moments;
mod builtins_statistical_more;
mod builtins_statistical_regression;
mod builtins_text;
mod builtins_text_dbcs;
mod builtins_thai;

// On wasm targets, `inventory` registrations can be dropped by the linker if the codegen unit
// contains no otherwise-referenced symbols (common when everything is registered via
// `inventory::submit!`). This leads to missing built-ins at runtime (e.g. `SORT()` returning
// `#NAME?`).
//
// Force-link every module that contains `inventory::submit!` registrations by referencing each
// module's `__force_link()` function. We intentionally "use" these function pointers via
// `black_box` so the optimizer cannot prove the calls are no-ops and remove the references during
// LTO.
#[cfg(target_arch = "wasm32")]
fn force_link_inventory_modules() {
    let builtins: &[fn()] = &[
        builtins_array::__force_link,
        builtins_cube::__force_link,
        builtins_database::__force_link,
        builtins_date_time::__force_link,
        builtins_dynamic_arrays::__force_link,
        builtins_dynamic_array_textsplit::__force_link,
        builtins_engineering::__force_link,
        builtins_engineering_complex::__force_link,
        builtins_engineering_convert::__force_link,
        builtins_engineering_special::__force_link,
        builtins_image::__force_link,
        builtins_information::__force_link,
        builtins_information_workbook::__force_link,
        builtins_information_worksheet::__force_link,
        builtins_lambda::__force_link,
        builtins_logical::__force_link,
        builtins_logical_constants::__force_link,
        builtins_logical_extended::__force_link,
        builtins_lookup::__force_link,
        builtins_math::__force_link,
        builtins_math_matrix::__force_link,
        builtins_math_extended::__force_link,
        builtins_math_more::__force_link,
        builtins_roman::__force_link,
        builtins_select::__force_link,
        builtins_reference::__force_link,
        builtins_statistical::__force_link,
        builtins_statistical_distributions::__force_link,
        builtins_statistical_ets::__force_link,
        builtins_statistical_moments::__force_link,
        builtins_statistical_more::__force_link,
        builtins_statistical_regression::__force_link,
        builtins_text::__force_link,
        builtins_text_dbcs::__force_link,
        builtins_thai::__force_link,
        financial::__force_link,
    ];

    for f in builtins {
        let f = std::hint::black_box(*f);
        f();
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Volatility {
    NonVolatile,
    Volatile,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ThreadSafety {
    ThreadSafe,
    NotThreadSafe,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArraySupport {
    ScalarOnly,
    SupportsArrays,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ValueType {
    Any,
    Number,
    Text,
    Bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Reference {
    pub sheet_id: SheetId,
    pub start: CellAddr,
    pub end: CellAddr,
}

impl Reference {
    pub fn normalized(&self) -> Self {
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

    pub fn is_single_cell(&self) -> bool {
        self.start == self.end
    }

    pub fn size(&self) -> u64 {
        let norm = self.normalized();
        let rows = norm.end.row as u64 - norm.start.row as u64 + 1;
        let cols = norm.end.col as u64 - norm.start.col as u64 + 1;
        rows.saturating_mul(cols)
    }

    pub fn contains(&self, addr: CellAddr) -> bool {
        let norm = self.normalized();
        addr.row >= norm.start.row
            && addr.row <= norm.end.row
            && addr.col >= norm.start.col
            && addr.col <= norm.end.col
    }

    pub fn iter_cells(&self) -> impl Iterator<Item = CellAddr> {
        let norm = self.normalized();
        let rows = norm.start.row..=norm.end.row;
        let cols = norm.start.col..=norm.end.col;
        rows.flat_map(move |row| cols.clone().map(move |col| CellAddr { row, col }))
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ArgValue {
    Scalar(Value),
    Reference(Reference),
    /// A multi-area reference produced by the union/intersection operators.
    ReferenceUnion(Vec<Reference>),
}

pub trait FunctionContext {
    fn eval_arg(&self, expr: &CompiledExpr) -> ArgValue;
    fn eval_scalar(&self, expr: &CompiledExpr) -> Value;
    fn eval_formula(&self, expr: &CompiledExpr) -> Value;
    fn eval_formula_with_bindings(
        &self,
        expr: &CompiledExpr,
        bindings: &HashMap<String, Value>,
    ) -> Value;
    fn capture_lexical_env(&self) -> HashMap<String, Value>;
    fn apply_implicit_intersection(&self, reference: &Reference) -> Value;
    fn get_cell_value(&self, sheet_id: &SheetId, addr: CellAddr) -> Value;
    /// Returns the cell's stored phonetic guide (furigana) text, if available.
    ///
    /// This is used by the `PHONETIC` worksheet function. Most backends do not model phonetic
    /// guides, so the default implementation returns `None`.
    fn get_cell_phonetic(&self, _sheet_id: &SheetId, _addr: CellAddr) -> Option<&str> {
        None
    }
    /// Returns the effective horizontal alignment for the given cell, if available.
    ///
    /// This is used by worksheet information functions like `CELL("prefix")`.
    fn cell_horizontal_alignment(
        &self,
        _sheet_id: &SheetId,
        _addr: CellAddr,
    ) -> Option<HorizontalAlignment> {
        None
    }
    fn iter_reference_cells<'a>(
        &'a self,
        reference: &'a Reference,
    ) -> Box<dyn Iterator<Item = CellAddr> + 'a>;
    /// Records that `reference` was dereferenced during evaluation.
    ///
    /// Implementations may use this to build dynamic dependency sets. The default is a no-op so
    /// callers that do not care about dependency tracing can ignore it.
    fn record_reference(&self, _reference: &Reference) {}
    fn now_utc(&self) -> chrono::DateTime<chrono::Utc>;
    /// Workbook calculation mode (automatic vs manual).
    ///
    /// This is primarily surfaced for worksheet information functions like `INFO("recalc")`.
    fn calculation_mode(&self) -> crate::calc_settings::CalculationMode {
        crate::calc_settings::CalculationMode::Automatic
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
    /// Returns the upper-left visible cell on the current worksheet, as an absolute A1 reference
    /// (e.g. `"$A$1"`).
    fn info_origin(&self) -> Option<&str> {
        None
    }

    fn date_system(&self) -> ExcelDateSystem;
    fn current_sheet_id(&self) -> usize;
    fn current_cell_addr(&self) -> CellAddr;
    /// Resolve a sheet id back to its display name.
    ///
    /// This is used by worksheet information functions like `CELL("address")`.
    fn sheet_name(&self, _sheet_id: usize) -> Option<&str> {
        None
    }
    /// Return the number of worksheets in the workbook.
    ///
    /// This is used by worksheet information functions like `INFO("numfile")`.
    fn sheet_count(&self) -> usize {
        1
    }
    /// Returns the current worksheet tab order index for `sheet_id` (0-based).
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
        Some(sheet_id)
    }
    /// Convenience wrapper around [`FunctionContext::sheet_name`] for the current sheet.
    fn current_sheet_name(&self) -> Option<&str> {
        self.sheet_name(self.current_sheet_id())
    }
    /// Returns the logical worksheet dimensions (row/column count) for `sheet_id`.
    ///
    /// This is used by functions that need to validate or reason about worksheet bounds (e.g.
    /// `OFFSET`, or special-casing whole-row/whole-column references in `ROW`/`COLUMN`).
    ///
    /// Implementations that do not track sheet dimensions can fall back to Excel's defaults.
    fn sheet_dimensions(&self, _sheet_id: &SheetId) -> (u32, u32) {
        (EXCEL_MAX_ROWS, EXCEL_MAX_COLS)
    }
    /// Returns the sheet default column width in Excel "character" units.
    ///
    /// This corresponds to the worksheet's `<sheetFormatPr defaultColWidth="...">` metadata.
    fn sheet_default_col_width(&self, _sheet_id: &SheetId) -> Option<f32> {
        None
    }

    /// Returns the top-left visible cell ("origin") for a worksheet view, if provided by the host.
    ///
    /// This is used by Excel-compatibility functions like `INFO("origin")`. The formula engine is
    /// deterministic and does not inspect UI state directly; hosts should provide view metadata
    /// explicitly via engine APIs.
    fn sheet_origin_cell(&self, _sheet_id: usize) -> Option<CellAddr> {
        None
    }
    /// Returns the stored formula text for a cell (including the leading `=`), if available.
    ///
    /// This is used by worksheet information functions like `CELL("contents")`.
    fn get_cell_formula(&self, _sheet_id: &SheetId, _addr: CellAddr) -> Option<&str> {
        None
    }

    /// Returns the number format string for a cell, if available.
    ///
    /// This is used by worksheet information functions like `CELL("format")` /
    /// `CELL("color")` / `CELL("parentheses")`.
    ///
    /// For external workbook references (`SheetId::External`), returning `None` is acceptable
    /// because the cell format is not available via the engine's external reference interfaces.
    fn get_cell_number_format(&self, _sheet_id: &SheetId, _addr: CellAddr) -> Option<&str> {
        None
    }
    /// Returns the workbook's style table, if available.
    fn style_table(&self) -> Option<&formula_model::StyleTable> {
        None
    }

    /// Return the default style id for an entire worksheet, if present.
    fn sheet_default_style_id(&self, _sheet_id: &SheetId) -> Option<u32> {
        None
    }

    /// Return the style id for a specific cell.
    ///
    /// Style id `0` is always the default (empty) style.
    fn cell_style_id(&self, _sheet_id: &SheetId, _addr: CellAddr) -> u32 {
        0
    }

    /// Return the style id from the compressed range-run formatting layer for a cell, if present.
    ///
    /// This corresponds to the `formatRunsByCol` / `setFormatRunsByCol` representation used by
    /// DocumentController hydration/deltas.
    ///
    /// Style id `0` indicates no run-style override (default formatting).
    ///
    /// Style precedence matches the DocumentController layering:
    /// `sheet < col < row < range-run < cell`.
    fn format_run_style_id(&self, _sheet_id: &SheetId, _addr: CellAddr) -> u32 {
        0
    }

    /// Return the default style id for an entire row, if present.
    fn row_style_id(&self, _sheet_id: &SheetId, _row: u32) -> Option<u32> {
        None
    }

    /// Return per-column properties (width/hidden/default style), if present.
    ///
    /// `ColProperties.width` is expressed in Excel "character" units (OOXML `col/@width`), not pixels.
    fn col_properties(
        &self,
        _sheet_id: &SheetId,
        _col: u32,
    ) -> Option<formula_model::ColProperties> {
        None
    }

    /// Return the style id from the range-run formatting layer for a cell, if present.
    ///
    /// This corresponds to DocumentController's `formatRunsByCol` layer (large range formatting
    /// rectangles compressed into per-column runs).
    ///
    /// Style id `0` indicates "no run applies".
    fn range_run_style_id(&self, _sheet_id: &SheetId, _addr: CellAddr) -> u32 {
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

    /// Returns the in-memory pivot registry (if available) for resolving `GETPIVOTDATA`.
    ///
    /// Implementations that do not support pivots can return `None`.
    fn pivot_registry(&self) -> Option<&crate::pivot_registry::PivotRegistry> {
        None
    }

    /// Resolve effective formatting/protection metadata for a cell.
    ///
    /// This is used by worksheet information functions like `CELL("prefix")` and `CELL("protect")`.
    ///
    /// Implementations that do not track formatting can return the default style values.
    fn effective_cell_style(
        &self,
        _sheet_id: &SheetId,
        _addr: CellAddr,
    ) -> crate::style_patch::EffectiveStyle {
        crate::style_patch::EffectiveStyle::default()
    }

    /// Resolve a worksheet name to an internal sheet id for runtime-parsed sheet references.
    ///
    /// Implementations should match Excel's Unicode-aware, NFKC + case-insensitive comparison
    /// semantics (see [`formula_model::sheet_name_eq_case_insensitive`]).
    fn resolve_sheet_name(&self, _name: &str) -> Option<usize> {
        None
    }

    /// Return the sheet order for an external workbook.
    ///
    /// This is used by workbook information functions like `SHEET` when passed an external
    /// reference such as `=[Book.xlsx]Sheet1!A1`. Implementations should return the workbook's
    /// sheet names (without the `[Book.xlsx]` prefix) in tab order.
    ///
    /// Returning `None` indicates that the sheet order is unavailable.
    fn external_sheet_order(&self, _workbook: &str) -> Option<Vec<String>> {
        None
    }

    /// Optional external data provider used by RTD / CUBE* functions.
    fn external_data_provider(&self) -> Option<&dyn crate::ExternalDataProvider> {
        None
    }

    /// Optional API for discovering the sheet order of an external workbook.
    ///
    /// This enables workbook/worksheet information functions to map an external sheet name to its
    /// 1-based position within the external workbook (e.g. `SHEET([Book.xlsx]Sheet2!A1)`).
    fn workbook_sheet_names(&self, workbook: &str) -> Option<Arc<[String]>> {
        self.external_sheet_order(workbook).map(Arc::from)
    }

    /// Locale configuration used when parsing locale-sensitive strings at runtime.
    ///
    /// This is primarily used for parsing numbers that appear inside string literals, such as
    /// criteria arguments (`">1,5"` in `de-DE`).
    fn locale_config(&self) -> LocaleConfig {
        LocaleConfig::en_us()
    }

    /// Locale used for implicit numeric coercion (text -> number).
    ///
    /// This is plumbed through evaluation so we can eventually respect workbook locale for
    /// implicit coercions and for VALUE/NUMBERVALUE.
    fn number_locale(&self) -> crate::value::NumberLocale {
        let separators = self.value_locale().separators;
        crate::value::NumberLocale::new(separators.decimal_sep, Some(separators.thousands_sep))
    }

    fn value_locale(&self) -> ValueLocaleConfig {
        ValueLocaleConfig::default()
    }

    /// Workbook text codepage (Windows code page number).
    ///
    /// This is used for legacy DBCS semantics (e.g. `ASC` / `DBCS`, and eventually `*B`
    /// byte-count functions) which depend on the active workbook locale / codepage.
    ///
    /// The engine defaults to an en-US workbook locale, which corresponds to Windows-1252.
    fn text_codepage(&self) -> u16 {
        1252
    }

    fn push_local_scope(&self);
    fn pop_local_scope(&self);
    fn set_local(&self, name: &str, value: ArgValue);

    fn make_lambda(&self, params: Vec<String>, body: CompiledExpr) -> Value;
    fn eval_lambda(&self, lambda: &Lambda, args: Vec<ArgValue>) -> Value;

    /// Deterministic per-recalc random bits, scoped to the current cell evaluation.
    ///
    /// This is used by volatile worksheet RNG functions (e.g. RAND, RANDBETWEEN) so that
    /// results are stable within a single recalculation and independent of scheduling
    /// order (single-threaded vs multi-threaded).
    fn volatile_rand_u64(&self) -> u64;

    fn volatile_rand(&self) -> f64 {
        let bits = self.volatile_rand_u64() >> 11; // 53 bits.
        (bits as f64) / ((1u64 << 53) as f64)
    }
}

pub type FunctionImpl = fn(&dyn FunctionContext, &[CompiledExpr]) -> Value;

#[derive(Clone, Copy)]
pub struct FunctionSpec {
    pub name: &'static str,
    pub min_args: usize,
    pub max_args: usize,
    pub volatility: Volatility,
    pub thread_safety: ThreadSafety,
    pub array_support: ArraySupport,
    pub return_type: ValueType,
    pub arg_types: &'static [ValueType],
    pub implementation: FunctionImpl,
}

inventory::collect!(FunctionSpec);

/// Iterate all [`FunctionSpec`] registrations collected via [`inventory`].
///
/// This is primarily intended for cross-crate test coverage (e.g. ensuring XLSB
/// BIFF function-id mappings cover every function that the engine can evaluate).
pub fn iter_function_specs() -> impl Iterator<Item = &'static FunctionSpec> {
    inventory::iter::<FunctionSpec>.into_iter()
}

fn registry() -> &'static HashMap<String, &'static FunctionSpec> {
    static REGISTRY: OnceLock<HashMap<String, &'static FunctionSpec>> = OnceLock::new();
    REGISTRY.get_or_init(|| {
        #[cfg(target_arch = "wasm32")]
        force_link_inventory_modules();

        let mut map = HashMap::new();
        for spec in inventory::iter::<FunctionSpec> {
            map.insert(spec.name.to_ascii_uppercase(), spec);
        }

        // Internal synthetic functions used by expression lowering.
        map.insert(
            builtins_rich_values::FIELDACCESS_SPEC
                .name
                .to_ascii_uppercase(),
            &builtins_rich_values::FIELDACCESS_SPEC,
        );
        map
    })
}

pub fn lookup_function(name: &str) -> Option<&'static FunctionSpec> {
    let upper = name.to_ascii_uppercase();
    if let Some(spec) = registry().get(&upper).copied() {
        return Some(spec);
    }

    // Excel stores newer functions in files with an `_xlfn.` prefix (e.g. `_xlfn.XLOOKUP`).
    // For evaluation we treat these as aliases of the unprefixed built-in.
    if let Some(stripped) = upper.strip_prefix("_XLFN.") {
        return registry().get(stripped).copied();
    }

    None
}

pub fn call_function(ctx: &dyn FunctionContext, name: &str, args: &[CompiledExpr]) -> Value {
    let spec = match lookup_function(name) {
        Some(spec) => spec,
        None => return Value::Error(ErrorKind::Name),
    };

    if args.len() < spec.min_args || args.len() > spec.max_args {
        return Value::Error(ErrorKind::Value);
    }

    (spec.implementation)(ctx, args)
}

/// Generate an unbiased uniform integer in `[0, span)`, using the deterministic per-recalc RNG.
///
/// Excel's integer RNG helpers (RANDBETWEEN/RANDARRAY whole_number=TRUE) should be uniform across
/// the integer interval. Using a simple `% span` introduces modulo bias when `span` is not a power
/// of two. We avoid that by rejection sampling.
pub(crate) fn volatile_rand_u64_below(ctx: &dyn FunctionContext, span: u64) -> u64 {
    if span <= 1 {
        return 0;
    }

    // Accept only values within a prefix of the u64 range whose length is a multiple of `span`.
    // The resulting distribution of `value % span` is uniform.
    let zone = (u64::MAX / span) * span;
    loop {
        let v = ctx.volatile_rand_u64();
        if v < zone {
            return v % span;
        }
    }
}

pub(crate) fn eval_scalar_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Value {
    match ctx.eval_arg(expr) {
        ArgValue::Scalar(v) => v,
        ArgValue::Reference(r) => ctx.apply_implicit_intersection(&r),
        ArgValue::ReferenceUnion(ranges) => apply_implicit_intersection_union(ctx, &ranges),
    }
}

fn apply_implicit_intersection_union(ctx: &dyn FunctionContext, ranges: &[Reference]) -> Value {
    // Excel's implicit intersection on a multi-area reference is ambiguous; we approximate by
    // succeeding only when exactly one area intersects.
    let mut hits = Vec::new();
    for r in ranges {
        let v = ctx.apply_implicit_intersection(r);
        if !matches!(v, Value::Error(ErrorKind::Value)) {
            hits.push(v);
        }
    }
    match hits.as_slice() {
        [only] => only.clone(),
        _ => Value::Error(ErrorKind::Value),
    }
}
