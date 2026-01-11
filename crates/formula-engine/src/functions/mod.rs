use std::collections::HashMap;
use std::sync::OnceLock;

use crate::date::ExcelDateSystem;
use crate::eval::{CellAddr, CompiledExpr};
use crate::value::{ErrorKind, Value};

pub mod date_time;
pub mod financial;
pub mod information;
pub mod lookup;
pub mod math;
pub mod statistical;
pub mod text;
pub(crate) mod array_lift;
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
mod builtins_date_time;
mod builtins_dynamic_arrays;
mod builtins_information;
mod builtins_lambda;
mod builtins_logical;
mod builtins_logical_extended;
mod builtins_lookup;
mod builtins_math;
mod builtins_math_extended;
mod builtins_select;
mod builtins_reference;
mod builtins_statistical;
mod builtins_text;

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

#[derive(Debug, Clone)]
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
    fn eval_formula_with_bindings(&self, expr: &CompiledExpr, bindings: &HashMap<String, Value>) -> Value;
    fn capture_lexical_env(&self) -> HashMap<String, Value>;
    fn apply_implicit_intersection(&self, reference: &Reference) -> Value;
    fn get_cell_value(&self, sheet_id: &SheetId, addr: CellAddr) -> Value;
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
    fn date_system(&self) -> ExcelDateSystem;
    fn current_sheet_id(&self) -> usize;
    fn current_cell_addr(&self) -> CellAddr;
    fn resolve_sheet_name(&self, _name: &str) -> Option<usize> {
        None
    }

    /// Locale used for implicit numeric coercion (text -> number).
    ///
    /// This is plumbed through evaluation so we can eventually respect workbook locale for
    /// implicit coercions and for VALUE/NUMBERVALUE.
    fn number_locale(&self) -> crate::value::NumberLocale {
        crate::value::NumberLocale::en_us()
    }

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
        let mut map = HashMap::new();
        for spec in inventory::iter::<FunctionSpec> {
            map.insert(spec.name.to_ascii_uppercase(), spec);
        }
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
