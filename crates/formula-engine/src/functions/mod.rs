use std::collections::HashMap;
use std::sync::OnceLock;

use crate::eval::{CellAddr, CompiledExpr};
use crate::value::{ErrorKind, Value};

pub mod financial;
pub mod information;
pub mod lookup;
pub mod math;
pub mod text;
pub mod date_time;

// Built-in Excel-compatible functions registered with the inventory-backed
// registry live in dedicated modules to avoid merge conflicts.
mod builtins_date_time;
mod builtins_logical;
mod builtins_lookup;
mod builtins_math;
mod builtins_text;
mod builtins_array;

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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Reference {
    pub sheet_id: usize,
    pub start: CellAddr,
    pub end: CellAddr,
}

impl Reference {
    pub fn normalized(self) -> Self {
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
            sheet_id: self.sheet_id,
            start: CellAddr { row: r1, col: c1 },
            end: CellAddr { row: r2, col: c2 },
        }
    }

    pub fn is_single_cell(self) -> bool {
        self.start == self.end
    }

    pub fn iter_cells(self) -> impl Iterator<Item = CellAddr> {
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
}

pub trait FunctionContext {
    fn eval_arg(&self, expr: &CompiledExpr) -> ArgValue;
    fn eval_scalar(&self, expr: &CompiledExpr) -> Value;
    fn apply_implicit_intersection(&self, reference: Reference) -> Value;
    fn get_cell_value(&self, sheet_id: usize, addr: CellAddr) -> Value;
    fn now_utc(&self) -> chrono::DateTime<chrono::Utc>;
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
        ArgValue::Reference(r) => ctx.apply_implicit_intersection(r),
    }
}
