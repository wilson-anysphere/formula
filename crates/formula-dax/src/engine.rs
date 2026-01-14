//! DAX evaluation engine.
//!
//! Relationship filtering in Tabular models is global: a filter on one table can restrict rows in
//! another table through active relationships, and bidirectional relationships can create cycles.
//! `formula-dax` models this by resolving a [`FilterContext`] into per-table row sets and then
//! repeatedly propagating constraints across relationships until reaching a fixed point (see
//! `resolve_row_sets` / `propagate_filter`).
//!
//! For many-to-many relationships ([`Cardinality::ManyToMany`]), propagation uses the **distinct
//! set of visible key values** on the source side (conceptually similar to
//! `TREATAS(VALUES(source[key]), target[key])`) instead of requiring a unique lookup row.
use crate::backend::TableBackend;
use crate::model::{
    normalize_ident, Cardinality, CrossFilterDirection, DataModel, RelationshipInfo,
    RelationshipPathDirection, RowSet, ToIndex,
};
use crate::parser::{BinaryOp, Expr, UnaryOp};
use crate::value::Value;
use formula_columnar::BitVec;
use ordered_float::OrderedFloat;
use std::borrow::Cow;
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet, VecDeque};
use std::sync::atomic::{AtomicBool, Ordering as AtomicOrdering};
use std::sync::Arc;
use std::sync::OnceLock;

pub type DaxResult<T> = Result<T, DaxError>;

#[derive(Debug, thiserror::Error)]
pub enum DaxError {
    #[error("parse error: {0}")]
    Parse(String),

    #[error("unknown table: {0}")]
    UnknownTable(String),

    #[error("unknown measure: {0}")]
    UnknownMeasure(String),

    #[error("unknown column {table}[{column}]")]
    UnknownColumn { table: String, column: String },

    #[error("duplicate table: {table}")]
    DuplicateTable { table: String },

    #[error("duplicate column {table}[{column}]")]
    DuplicateColumn { table: String, column: String },

    #[error("duplicate measure: {measure}")]
    DuplicateMeasure { measure: String },

    #[error("schema mismatch for {table}: expected {expected} values, got {actual}")]
    SchemaMismatch {
        table: String,
        expected: usize,
        actual: usize,
    },

    #[error("calculated column length mismatch for {table}[{column}]: expected {expected} values, got {actual}")]
    ColumnLengthMismatch {
        table: String,
        column: String,
        expected: usize,
        actual: usize,
    },

    #[error("unsupported relationship cardinality {cardinality:?} in {relationship}")]
    UnsupportedCardinality {
        relationship: String,
        cardinality: Cardinality,
    },

    #[error("non-unique key in {table}[{column}]: {value}")]
    NonUniqueKey {
        table: String,
        column: String,
        value: Value,
    },

    #[error(
        "referential integrity violation in relationship {relationship}: value {value} in {from_table}[{from_column}] has no match in {to_table}[{to_column}]"
    )]
    ReferentialIntegrityViolation {
        relationship: String,
        from_table: String,
        from_column: String,
        to_table: String,
        to_column: String,
        value: Value,
    },

    #[error(
        "relationship {relationship} join columns have incompatible types: {from_table}[{from_column}] ({from_type}) vs {to_table}[{to_column}] ({to_type})"
    )]
    RelationshipJoinColumnTypeMismatch {
        relationship: String,
        from_table: String,
        from_column: String,
        from_type: String,
        to_table: String,
        to_column: String,
        to_type: String,
    },

    #[error("type error: {0}")]
    Type(String),

    #[error("evaluation error: {0}")]
    Eval(String),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RelationshipOverride {
    Active(CrossFilterDirection),
    OneWayReverse,
    Disabled,
}

#[derive(Clone, Debug)]
enum RowFilter {
    /// All physical rows of the table are included.
    ///
    /// This is primarily used to avoid allocating a huge `HashSet<usize>` for filters like
    /// `ALLNOBLANKROW(Table)` when used as a CALCULATE filter argument. Even though this does not
    /// restrict physical rows, the *presence* of a row filter still suppresses the
    /// relationship-generated blank/unknown member (see `blank_row_allowed`).
    All,
    /// An explicit set of physical row indices.
    Rows(HashSet<usize>),
    /// A bitmap of allowed physical row indices.
    Mask(Arc<BitVec>),
}

#[derive(Clone, Debug, Default)]
pub struct FilterContext {
    column_filters: HashMap<(String, String), HashSet<Value>>,
    row_filters: HashMap<String, RowFilter>,
    active_relationship_overrides: HashSet<usize>,
    cross_filter_overrides: HashMap<usize, RelationshipOverride>,
    suppress_implicit_measure_context_transition: bool,
    /// Set of (table, column) references that are considered "in scope" for the current
    /// evaluation.
    ///
    /// This is **pivot-driven metadata** used to implement a minimal version of `ISINSCOPE`.
    /// Unlike full DAX engines, `formula-dax` does not currently infer scope from arbitrary DAX
    /// queries; instead, the pivot engine explicitly populates this set with its axis columns.
    ///
    /// Non-goals / limitations (MVP):
    /// - We do not attempt to infer scope from arbitrary DAX query shapes.
    /// - Related filter-introspection functions like `ISFILTERED` / `HASONEFILTER` are not yet
    ///   implemented.
    pub(crate) in_scope_columns: HashSet<(String, String)>,
}

impl FilterContext {
    pub fn empty() -> Self {
        Self::default()
    }

    pub fn is_empty(&self) -> bool {
        self.column_filters.is_empty() && self.row_filters.is_empty()
    }

    pub fn with_column_equals(mut self, table: &str, column: &str, value: Value) -> Self {
        self.set_column_equals(table, column, value);
        self
    }

    pub fn with_column_in(
        mut self,
        table: &str,
        column: &str,
        values: impl IntoIterator<Item = Value>,
    ) -> Self {
        self.set_column_in(table, column, values);
        self
    }

    pub fn set_column_equals(&mut self, table: &str, column: &str, value: Value) {
        let table = normalize_ident(table);
        let column = normalize_ident(column);
        self.column_filters
            .insert((table, column), HashSet::from([value]));
    }

    pub fn set_column_in(
        &mut self,
        table: &str,
        column: &str,
        values: impl IntoIterator<Item = Value>,
    ) {
        let table = normalize_ident(table);
        let column = normalize_ident(column);
        self.column_filters
            .insert((table, column), values.into_iter().collect());
    }

    pub fn clear_column_filter_public(&mut self, table: &str, column: &str) {
        self.clear_column_filter(table, column);
    }

    pub fn clear_table_filters_public(&mut self, table: &str) {
        self.clear_table_filters(table);
    }

    pub(crate) fn relationship_overrides(&self) -> &HashSet<usize> {
        &self.active_relationship_overrides
    }

    pub(crate) fn is_relationship_disabled(&self, relationship_idx: usize) -> bool {
        matches!(
            self.cross_filter_overrides.get(&relationship_idx).copied(),
            Some(RelationshipOverride::Disabled)
        )
    }

    fn activate_relationship(&mut self, relationship_idx: usize) {
        self.active_relationship_overrides.insert(relationship_idx);
    }

    fn clear_table_filters(&mut self, table: &str) {
        let table = normalize_ident(table);
        self.column_filters.retain(|(t, _), _| t != &table);
        self.row_filters.remove(&table);
    }

    fn clear_column_filter(&mut self, table: &str, column: &str) {
        let key = (normalize_ident(table), normalize_ident(column));
        self.column_filters.remove(&key);
    }

    fn set_row_filter(&mut self, table: &str, rows: RowFilter) {
        let table = normalize_ident(table);
        self.row_filters.insert(table, rows);
    }
}

#[derive(Clone, Debug)]
enum RowContextFrame {
    Physical {
        table: String,
        row: usize,
        /// If set, restrict the row context to only these column indices. This is used for
        /// single-column table expressions like `VALUES(Table[Column])` where DAX exposes only the
        /// column values, not the full underlying physical row.
        visible_cols: Option<Vec<usize>>,
    },
    Virtual {
        /// Explicit (table,column) -> value bindings for a virtual table row (e.g. a `SUMMARIZE`
        /// grouping key). Only these columns are visible in row context, and context transition
        /// should apply filters only for these bindings.
        bindings: Vec<((String, String), Value)>,
    },
}

#[derive(Clone, Debug, Default)]
pub struct RowContext {
    stack: Vec<RowContextFrame>,
}

impl RowContext {
    /// Push a full row context frame for `table`/`row` (all columns visible).
    pub fn push(&mut self, table: &str, row: usize) {
        self.push_physical(table, row, None);
    }

    pub(crate) fn push_physical(
        &mut self,
        table: &str,
        row: usize,
        visible_cols: Option<Vec<usize>>,
    ) {
        self.stack.push(RowContextFrame::Physical {
            table: normalize_ident(table),
            row,
            visible_cols,
        });
    }

    pub(crate) fn push_virtual(&mut self, bindings: Vec<((String, String), Value)>) {
        let bindings = bindings
            .into_iter()
            .map(|((t, c), v)| ((normalize_ident(&t), normalize_ident(&c)), v))
            .collect();
        self.stack.push(RowContextFrame::Virtual { bindings });
    }

    pub fn pop(&mut self) {
        self.stack.pop();
    }

    /// Update the row index for the innermost (top-of-stack) *physical* row context.
    ///
    /// This is useful in hot loops (e.g. calculated-column evaluation) where we want to reuse a
    /// single [`RowContext`] and avoid allocating a new table name string for each row.
    pub fn set_current_row(&mut self, row: usize) {
        if let Some(RowContextFrame::Physical {
            row: current_row, ..
        }) = self.stack.last_mut()
        {
            *current_row = row;
        }
    }

    fn is_empty(&self) -> bool {
        self.stack.is_empty()
    }

    /// The "current table" is the most recent *physical* row context table.
    ///
    /// Virtual row contexts (e.g. from `SUMMARIZE`) do not have a single current table name.
    fn current_table(&self) -> Option<&str> {
        self.stack.iter().rev().find_map(|frame| match frame {
            RowContextFrame::Physical { table, .. } => Some(table.as_str()),
            RowContextFrame::Virtual { .. } => None,
        })
    }

    fn physical_row_for(&self, table: &str) -> Option<(usize, Option<&[usize]>)> {
        let table = normalize_ident(table);
        self.stack.iter().rev().find_map(|frame| match frame {
            RowContextFrame::Physical {
                table: t,
                row,
                visible_cols,
            } if t == &table => Some((*row, visible_cols.as_deref())),
            _ => None,
        })
    }

    fn physical_row_for_level(
        &self,
        table: &str,
        level_from_inner: usize,
    ) -> Option<(usize, Option<&[usize]>)> {
        let table = normalize_ident(table);
        self.stack
            .iter()
            .rev()
            .filter_map(|frame| match frame {
                RowContextFrame::Physical {
                    table: t,
                    row,
                    visible_cols,
                } if t == &table => Some((*row, visible_cols.as_deref())),
                _ => None,
            })
            .nth(level_from_inner)
    }
    fn physical_row_for_outermost(&self, table: &str) -> Option<(usize, Option<&[usize]>)> {
        let table = normalize_ident(table);
        self.stack.iter().find_map(|frame| match frame {
            RowContextFrame::Physical {
                table: t,
                row,
                visible_cols,
            } if t == &table => Some((*row, visible_cols.as_deref())),
            _ => None,
        })
    }

    fn virtual_binding(&self, table: &str, column: &str) -> Option<&Value> {
        let table = normalize_ident(table);
        let column = normalize_ident(column);
        for frame in self.stack.iter().rev() {
            let RowContextFrame::Virtual { bindings } = frame else {
                continue;
            };
            for ((t, c), v) in bindings {
                if t == &table && c == &column {
                    return Some(v);
                }
            }
        }
        None
    }
}

#[derive(Clone, Debug)]
enum VarValue {
    Scalar(Value),
    Table(TableResult),
    /// A one-column virtual table value (currently produced by `{...}` table constructors in `VAR`
    /// bindings).
    OneColumnTable(Vec<Value>),
}

#[derive(Clone, Debug, Default)]
struct VarEnv {
    scopes: Vec<HashMap<String, VarValue>>,
}

impl VarEnv {
    fn normalize_name(name: &str) -> String {
        normalize_ident(name)
    }

    fn lookup(&self, name: &str) -> Option<&VarValue> {
        let key = Self::normalize_name(name);
        self.scopes.iter().rev().find_map(|scope| scope.get(&key))
    }

    fn push_scope(&mut self) {
        self.scopes.push(HashMap::new());
    }

    fn pop_scope(&mut self) {
        self.scopes.pop();
    }

    fn define(&mut self, name: &str, value: VarValue) {
        if self.scopes.is_empty() {
            self.push_scope();
        }
        let scope = self.scopes.last_mut().expect("just pushed if empty");
        scope.insert(Self::normalize_name(name), value);
    }
}

#[derive(Clone, Debug, Default)]
pub struct DaxEngine;

impl DaxEngine {
    pub fn new() -> Self {
        Self
    }

    /// Apply `CALCULATE`-style filter arguments to an existing filter context, returning the
    /// resulting [`FilterContext`].
    ///
    /// This is primarily useful for APIs that accept a [`FilterContext`] (like [`crate::pivot`])
    /// but need to support DAX filter expressions that can't be expressed with
    /// [`FilterContext::with_column_equals`], such as `Table[Column] <> BLANK()`.
    pub fn apply_calculate_filters(
        &self,
        model: &DataModel,
        filter: &FilterContext,
        filter_args: &[&str],
    ) -> DaxResult<FilterContext> {
        let mut parsed_args = Vec::with_capacity(filter_args.len());
        for arg in filter_args {
            parsed_args.push(crate::parser::parse(arg)?);
        }
        let row_ctx = RowContext::default();
        let mut env = VarEnv::default();
        self.build_calculate_filter(model, filter, &row_ctx, &parsed_args, &mut env)
    }

    pub fn evaluate(
        &self,
        model: &DataModel,
        expression: &str,
        filter: &FilterContext,
        row_ctx: &RowContext,
    ) -> DaxResult<Value> {
        let parsed = crate::parser::parse(expression)?;
        self.evaluate_expr(model, &parsed, filter, row_ctx)
    }

    pub fn evaluate_expr(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
        row_ctx: &RowContext,
    ) -> DaxResult<Value> {
        let mut env = VarEnv::default();
        self.eval_scalar(model, expr, filter, row_ctx, &mut env)
    }

    fn eval_scalar(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
        row_ctx: &RowContext,
        env: &mut VarEnv,
    ) -> DaxResult<Value> {
        match expr {
            Expr::Number(n) => Ok(Value::from(*n)),
            Expr::Text(s) => Ok(Value::from(s.clone())),
            Expr::Boolean(b) => Ok(Value::from(*b)),
            Expr::TableLiteral { .. } => Err(DaxError::Type(
                "table constructor used in scalar context".into(),
            )),
            Expr::Tuple(_) => Err(DaxError::Type(
                "row constructor used in scalar context".into(),
            )),
            Expr::Measure(name) => {
                let ident = DataModel::normalize_measure_name(name);
                let measure_key = normalize_ident(ident);
                if let Some(measure) = model.measures().get(&measure_key) {
                    // In DAX, evaluating a measure inside a row context implicitly performs a
                    // context transition (equivalent to `CALCULATE([Measure])`).
                    let eval_filter = if !row_ctx.is_empty()
                        && !filter.suppress_implicit_measure_context_transition
                    {
                        self.apply_context_transition(model, filter, row_ctx)?
                    } else {
                        filter.clone()
                    };

                    let mut measure_env = VarEnv::default();
                    return self.eval_scalar(
                        model,
                        &measure.parsed,
                        &eval_filter,
                        &RowContext::default(),
                        &mut measure_env,
                    );
                }

                // DAX allows `[Column]` references in row context. Bracketed identifiers
                // are ambiguous (measure vs. column), so we parse them as `Expr::Measure`
                // and resolve as a column when no measure is defined.
                let normalized = &measure_key;
                //
                // For virtual row contexts (e.g. iterators over `SUMMARIZE`), the "current row"
                // consists of explicit column bindings rather than a physical table row. In that
                // case, resolve `[Column]` by looking for a matching bound column name in the
                // innermost virtual frame.
                if let Some(RowContextFrame::Virtual { bindings }) = row_ctx.stack.last() {
                    let mut matches = bindings
                        .iter()
                        .filter(|((_, c), _)| c == normalized)
                        .map(|(_, v)| v);
                    let first = matches.next();
                    let second = matches.next();
                    match (first, second) {
                        (Some(v), None) => return Ok(v.clone()),
                        (Some(_), Some(_)) => {
                            return Err(DaxError::Eval(format!(
                            "ambiguous column reference [{normalized}] in the current row context"
                        )))
                        }
                        (None, _) => {
                            // Fall through: if there is an outer physical row context, use it as
                            // the bracket identifier target (matching existing behavior).
                        }
                    }
                }
                let Some(current_table) = row_ctx.current_table() else {
                    // Virtual row contexts (e.g. from `SUMMARIZE` or table constructors) do not
                    // have a single "current table". In those cases, attempt to resolve bracketed
                    // identifiers from a unique virtual binding by column name.
                    for frame in row_ctx.stack.iter().rev() {
                        let RowContextFrame::Virtual { bindings } = frame else {
                            continue;
                        };
                        let mut matched: Option<&Value> = None;
                        for ((_, column), value) in bindings {
                            if column == normalized {
                                if matched.is_some() {
                                    return Err(DaxError::Eval(format!(
                                        "ambiguous column reference [{normalized}] in virtual row context"
                                    )));
                                }
                                matched = Some(value);
                            }
                        }
                        if let Some(value) = matched {
                            return Ok(value.clone());
                        }
                    }
                    return Err(DaxError::UnknownMeasure(name.clone()));
                };
                let (row, visible_cols) = row_ctx
                    .physical_row_for(current_table)
                    .ok_or_else(|| DaxError::Eval(format!("no row context for [{ident}]")))?;
                let table_ref = model
                    .table(current_table)
                    .ok_or_else(|| DaxError::UnknownTable(current_table.to_string()))?;
                let Some(col_idx) = table_ref.column_idx(ident) else {
                    return Err(DaxError::Eval(format!(
                        "unknown measure [{ident}] and no column {}[{ident}]",
                        table_ref.name()
                    )));
                };
                if let Some(visible_cols) = visible_cols {
                    if !visible_cols.contains(&col_idx) {
                        return Err(DaxError::Eval(format!(
                            "column {}[{ident}] is not available in the current row context",
                            table_ref.name()
                        )));
                    }
                }
                if row >= table_ref.row_count() {
                    return Ok(Value::Blank);
                }
                Ok(table_ref.value_by_idx(row, col_idx).unwrap_or(Value::Blank))
            }
            Expr::Let { bindings, body } => {
                env.push_scope();
                let result = (|| -> DaxResult<Value> {
                    for (name, binding_expr) in bindings {
                        let value =
                            self.eval_var_value(model, binding_expr, filter, row_ctx, env)?;
                        env.define(name, value);
                    }
                    self.eval_scalar(model, body, filter, row_ctx, env)
                })();
                env.pop_scope();
                result
            }
            Expr::ColumnRef { table, column } => {
                if let Some(value) = row_ctx.virtual_binding(table, column) {
                    return Ok(value.clone());
                }

                let (row, visible_cols) = row_ctx.physical_row_for(table).ok_or_else(|| {
                    DaxError::Eval(format!("no row context for {table}[{column}]"))
                })?;
                let table_ref = model
                    .table(table)
                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                let idx = table_ref
                    .column_idx(column)
                    .ok_or_else(|| DaxError::UnknownColumn {
                        table: table.clone(),
                        column: column.clone(),
                    })?;
                if let Some(visible_cols) = visible_cols {
                    if !visible_cols.contains(&idx) {
                        return Err(DaxError::Eval(format!(
                            "column {table}[{column}] is not available in the current row context"
                        )));
                    }
                }
                if row >= table_ref.row_count() {
                    return Ok(Value::Blank);
                }
                Ok(table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank))
            }
            Expr::UnaryOp { op, expr } => {
                let value = self.eval_scalar(model, expr, filter, row_ctx, env)?;
                match op {
                    UnaryOp::Negate => {
                        let n = coerce_number(&value)?;
                        Ok(Value::from(-n))
                    }
                }
            }
            Expr::BinaryOp {
                op: BinaryOp::In,
                left,
                right,
            } => {
                let lhs_values: Vec<Value> = match left.as_ref() {
                    Expr::Tuple(values) => values
                        .iter()
                        .map(|expr| self.eval_scalar(model, expr, filter, row_ctx, env))
                        .collect::<DaxResult<_>>()?,
                    _ => vec![self.eval_scalar(model, left, filter, row_ctx, env)?],
                };
                let expected_cols = lhs_values.len();

                let col_count_error = |actual: usize| {
                    if expected_cols == 1 {
                        DaxError::Eval("IN currently only supports one-column tables".into())
                    } else {
                        DaxError::Eval(format!(
                            "IN row constructor expected {expected_cols} columns, got {actual}"
                        ))
                    }
                };

                match self.eval_table(model, right, filter, row_ctx, env) {
                    Ok(table_result) => match table_result {
                        TableResult::Physical {
                            table,
                            rows,
                            visible_cols,
                        } => {
                            let table_ref = model
                                .table(&table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;

                            let visible_cols = visible_cols.as_deref();
                            let col_count = visible_cols
                                .map(|cols| cols.len())
                                .unwrap_or_else(|| table_ref.columns().len());
                            if col_count != expected_cols {
                                return Err(col_count_error(col_count));
                            }

                            if let Some(cols) = visible_cols {
                                for row in rows {
                                    let mut matches = true;
                                    for (&col_idx, needle) in cols.iter().zip(&lhs_values) {
                                        let value = table_ref
                                            .value_by_idx(row, col_idx)
                                            .unwrap_or(Value::Blank);
                                        if !compare_values(&BinaryOp::Equals, needle, &value)? {
                                            matches = false;
                                            break;
                                        }
                                    }
                                    if matches {
                                        return Ok(Value::Boolean(true));
                                    }
                                }
                            } else {
                                for row in rows {
                                    let mut matches = true;
                                    for (col_idx, needle) in lhs_values.iter().enumerate() {
                                        let value = table_ref
                                            .value_by_idx(row, col_idx)
                                            .unwrap_or(Value::Blank);
                                        if !compare_values(&BinaryOp::Equals, needle, &value)? {
                                            matches = false;
                                            break;
                                        }
                                    }
                                    if matches {
                                        return Ok(Value::Boolean(true));
                                    }
                                }
                            }

                            Ok(Value::Boolean(false))
                        }
                        TableResult::PhysicalAll {
                            table,
                            row_count,
                            visible_cols,
                        } => {
                            let table_ref = model
                                .table(&table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;

                            let visible_cols = visible_cols.as_deref();
                            let col_count = visible_cols
                                .map(|cols| cols.len())
                                .unwrap_or_else(|| table_ref.columns().len());
                            if col_count != expected_cols {
                                return Err(col_count_error(col_count));
                            }

                            if let Some(cols) = visible_cols {
                                for row in 0..row_count {
                                    let mut matches = true;
                                    for (&col_idx, needle) in cols.iter().zip(&lhs_values) {
                                        let value = table_ref
                                            .value_by_idx(row, col_idx)
                                            .unwrap_or(Value::Blank);
                                        if !compare_values(&BinaryOp::Equals, needle, &value)? {
                                            matches = false;
                                            break;
                                        }
                                    }
                                    if matches {
                                        return Ok(Value::Boolean(true));
                                    }
                                }
                            } else {
                                for row in 0..row_count {
                                    let mut matches = true;
                                    for (col_idx, needle) in lhs_values.iter().enumerate() {
                                        let value = table_ref
                                            .value_by_idx(row, col_idx)
                                            .unwrap_or(Value::Blank);
                                        if !compare_values(&BinaryOp::Equals, needle, &value)? {
                                            matches = false;
                                            break;
                                        }
                                    }
                                    if matches {
                                        return Ok(Value::Boolean(true));
                                    }
                                }
                            }

                            Ok(Value::Boolean(false))
                        }
                        TableResult::PhysicalMask {
                            table,
                            mask,
                            visible_cols,
                        } => {
                            let table_ref = model
                                .table(&table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;

                            let visible_cols = visible_cols.as_deref();
                            let col_count = visible_cols
                                .map(|cols| cols.len())
                                .unwrap_or_else(|| table_ref.columns().len());
                            if col_count != expected_cols {
                                return Err(col_count_error(col_count));
                            }

                            if let Some(cols) = visible_cols {
                                for row in mask.iter_ones() {
                                    let mut matches = true;
                                    for (&col_idx, needle) in cols.iter().zip(&lhs_values) {
                                        let value = table_ref
                                            .value_by_idx(row, col_idx)
                                            .unwrap_or(Value::Blank);
                                        if !compare_values(&BinaryOp::Equals, needle, &value)? {
                                            matches = false;
                                            break;
                                        }
                                    }
                                    if matches {
                                        return Ok(Value::Boolean(true));
                                    }
                                }
                            } else {
                                for row in mask.iter_ones() {
                                    let mut matches = true;
                                    for (col_idx, needle) in lhs_values.iter().enumerate() {
                                        let value = table_ref
                                            .value_by_idx(row, col_idx)
                                            .unwrap_or(Value::Blank);
                                        if !compare_values(&BinaryOp::Equals, needle, &value)? {
                                            matches = false;
                                            break;
                                        }
                                    }
                                    if matches {
                                        return Ok(Value::Boolean(true));
                                    }
                                }
                            }

                            Ok(Value::Boolean(false))
                        }
                        TableResult::Virtual { columns, rows } => {
                            let col_count = columns.len();
                            if col_count != expected_cols {
                                return Err(col_count_error(col_count));
                            }

                            let blank = Value::Blank;
                            for row_values in rows {
                                let mut matches = true;
                                for (i, needle) in lhs_values.iter().enumerate() {
                                    let value = row_values.get(i).unwrap_or(&blank);
                                    if !compare_values(&BinaryOp::Equals, needle, value)? {
                                        matches = false;
                                        break;
                                    }
                                }
                                if matches {
                                    return Ok(Value::Boolean(true));
                                }
                            }
                            Ok(Value::Boolean(false))
                        }
                    },
                    Err(table_err) => {
                        if expected_cols != 1 {
                            return Err(table_err);
                        }

                        // Backward-compatible fallback for MVP: allow table constructors on the RHS.
                        let rhs_values =
                            self.eval_one_column_table_literal(model, right, filter, row_ctx, env);
                        let rhs_values = match rhs_values {
                            Ok(values) => values,
                            Err(_) => return Err(table_err),
                        };
                        let lhs = lhs_values.into_iter().next().expect("expected_cols == 1");
                        for candidate in rhs_values {
                            if compare_values(&BinaryOp::Equals, &lhs, &candidate)? {
                                return Ok(Value::Boolean(true));
                            }
                        }
                        Ok(Value::Boolean(false))
                    }
                }
            }
            Expr::BinaryOp {
                op: BinaryOp::And,
                left,
                right,
            } => {
                // DAX logical operators `&&` / `||` short-circuit.
                //
                // This matters for correctness (avoiding errors in the RHS when the LHS determines
                // the result) and matches common DAX usage patterns.
                let left = self.eval_scalar(model, left, filter, row_ctx, env)?;
                let left = left.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                if !left {
                    return Ok(Value::Boolean(false));
                }
                let right = self.eval_scalar(model, right, filter, row_ctx, env)?;
                let right = right.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                Ok(Value::Boolean(right))
            }
            Expr::BinaryOp {
                op: BinaryOp::Or,
                left,
                right,
            } => {
                // See `BinaryOp::And` above.
                let left = self.eval_scalar(model, left, filter, row_ctx, env)?;
                let left = left.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                if left {
                    return Ok(Value::Boolean(true));
                }
                let right = self.eval_scalar(model, right, filter, row_ctx, env)?;
                let right = right.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                Ok(Value::Boolean(right))
            }
            Expr::BinaryOp { op, left, right } => {
                let left = self.eval_scalar(model, left, filter, row_ctx, env)?;
                let right = self.eval_scalar(model, right, filter, row_ctx, env)?;
                self.eval_binary(op, left, right)
            }
            Expr::Call { name, args } => {
                self.eval_call_scalar(model, name, args, filter, row_ctx, env)
            }
            Expr::TableName(name) => match env.lookup(name) {
                Some(VarValue::Scalar(v)) => Ok(v.clone()),
                Some(VarValue::Table(_) | VarValue::OneColumnTable(_)) => Err(DaxError::Type(
                    format!("table variable {name} used in scalar context"),
                )),
                None => Err(DaxError::Type(format!(
                    "table {name} used in scalar context"
                ))),
            },
        }
    }

    fn eval_var_value(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
        row_ctx: &RowContext,
        env: &mut VarEnv,
    ) -> DaxResult<VarValue> {
        if let Expr::TableLiteral { rows } = expr {
            let col_count = rows.first().map(|row| row.len()).unwrap_or(1);
            if col_count == 1 {
                let values =
                    self.eval_one_column_table_literal(model, expr, filter, row_ctx, env)?;
                return Ok(VarValue::OneColumnTable(values));
            }

            let table = self.eval_table(model, expr, filter, row_ctx, env)?;
            return Ok(VarValue::Table(table));
        }

        match self.eval_scalar(model, expr, filter, row_ctx, env) {
            Ok(v) => Ok(VarValue::Scalar(v)),
            Err(err) => match self.eval_table(model, expr, filter, row_ctx, env) {
                Ok(t) => Ok(VarValue::Table(t)),
                Err(_) => Err(err),
            },
        }
    }

    fn eval_binary(&self, op: &BinaryOp, left: Value, right: Value) -> DaxResult<Value> {
        match op {
            BinaryOp::Add | BinaryOp::Subtract | BinaryOp::Multiply | BinaryOp::Divide => {
                let l = coerce_number(&left)?;
                let r = coerce_number(&right)?;
                let out = match op {
                    BinaryOp::Add => l + r,
                    BinaryOp::Subtract => l - r,
                    BinaryOp::Multiply => l * r,
                    BinaryOp::Divide => l / r,
                    _ => unreachable!(),
                };
                Ok(Value::from(out))
            }
            BinaryOp::Concat => {
                let l = coerce_text(&left);
                let r = coerce_text(&right);
                let mut out = String::with_capacity(l.len() + r.len());
                out.push_str(&l);
                out.push_str(&r);
                Ok(Value::from(out))
            }
            BinaryOp::Equals
            | BinaryOp::NotEquals
            | BinaryOp::Less
            | BinaryOp::LessEquals
            | BinaryOp::Greater
            | BinaryOp::GreaterEquals => Ok(Value::Boolean(compare_values(op, &left, &right)?)),
            BinaryOp::And | BinaryOp::Or => {
                let l = left.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                let r = right.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                Ok(Value::Boolean(match op {
                    BinaryOp::And => l && r,
                    BinaryOp::Or => l || r,
                    _ => unreachable!(),
                }))
            }
            BinaryOp::In => Err(DaxError::Type(
                "IN operator is only supported with a table constructor on the right-hand side"
                    .into(),
            )),
        }
    }

    fn eval_one_column_table_literal(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
        row_ctx: &RowContext,
        env: &mut VarEnv,
    ) -> DaxResult<Vec<Value>> {
        let rows = match expr {
            Expr::TableLiteral { rows } => rows,
            Expr::TableName(name) => match env.lookup(name) {
                Some(VarValue::OneColumnTable(values)) => return Ok(values.clone()),
                _ => {
                    return Err(DaxError::Type(format!(
                        "expected a one-column table constructor, got {expr:?}"
                    )))
                }
            },
            _ => {
                return Err(DaxError::Type(format!(
                    "expected a one-column table constructor, got {expr:?}"
                )))
            }
        };

        let mut out = Vec::with_capacity(rows.len());
        for row in rows {
            let [cell] = row.as_slice() else {
                return Err(DaxError::Type(
                    "only one-column table constructors are supported".into(),
                ));
            };
            out.push(self.eval_scalar(model, cell, filter, row_ctx, env)?);
        }
        Ok(out)
    }

    fn eval_call_scalar(
        &self,
        model: &DataModel,
        name: &str,
        args: &[Expr],
        filter: &FilterContext,
        row_ctx: &RowContext,
        env: &mut VarEnv,
    ) -> DaxResult<Value> {
        match name.to_ascii_uppercase().as_str() {
            "TRUE" => Ok(Value::Boolean(true)),
            "FALSE" => Ok(Value::Boolean(false)),
            "BLANK" => Ok(Value::Blank),
            "ISBLANK" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("ISBLANK expects 1 argument".into()));
                };
                let value = self.eval_scalar(model, arg, filter, row_ctx, env)?;
                Ok(Value::Boolean(value.is_blank()))
            }
            "IF" => {
                if args.len() < 2 || args.len() > 3 {
                    return Err(DaxError::Eval("IF expects 2 or 3 arguments".into()));
                }
                let cond = self.eval_scalar(model, &args[0], filter, row_ctx, env)?;
                let cond = cond.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                if cond {
                    self.eval_scalar(model, &args[1], filter, row_ctx, env)
                } else if args.len() == 3 {
                    self.eval_scalar(model, &args[2], filter, row_ctx, env)
                } else {
                    Ok(Value::Blank)
                }
            }
            "SWITCH" => {
                if args.len() < 3 {
                    return Err(DaxError::Eval("SWITCH expects at least 3 arguments".into()));
                }

                // DAX evaluates the expression once, then compares it against each value in
                // order, returning the result for the first match.
                let expr = self.eval_scalar(model, &args[0], filter, row_ctx, env)?;

                // DAX syntax:
                //   SWITCH(<expr>, <value1>, <result1>, ..., [<else>])
                //
                // After the initial expression, arguments come in (value, result) pairs.
                // If the total arity is even, an <else> expression is provided as the last
                // argument. Otherwise, missing <else> returns BLANK().
                let has_else = args.len() % 2 == 0;
                let pair_end = if has_else { args.len() - 1 } else { args.len() };

                let mut idx = 1usize;
                while idx + 1 < pair_end {
                    let value = self.eval_scalar(model, &args[idx], filter, row_ctx, env)?;
                    if compare_values(&BinaryOp::Equals, &expr, &value)? {
                        return self.eval_scalar(model, &args[idx + 1], filter, row_ctx, env);
                    }
                    idx += 2;
                }

                if has_else {
                    self.eval_scalar(model, &args[args.len() - 1], filter, row_ctx, env)
                } else {
                    Ok(Value::Blank)
                }
            }
            "DIVIDE" => {
                if args.len() < 2 || args.len() > 3 {
                    return Err(DaxError::Eval("DIVIDE expects 2 or 3 arguments".into()));
                }
                let numerator = self.eval_scalar(model, &args[0], filter, row_ctx, env)?;
                let denominator = self.eval_scalar(model, &args[1], filter, row_ctx, env)?;
                let denominator = coerce_number(&denominator)?;
                if denominator == 0.0 {
                    if args.len() == 3 {
                        self.eval_scalar(model, &args[2], filter, row_ctx, env)
                    } else {
                        Ok(Value::Blank)
                    }
                } else {
                    let numerator = coerce_number(&numerator)?;
                    Ok(Value::from(numerator / denominator))
                }
            }
            "COALESCE" => {
                if args.is_empty() {
                    return Err(DaxError::Eval(
                        "COALESCE expects at least 1 argument".into(),
                    ));
                }
                for arg in args {
                    let value = self.eval_scalar(model, arg, filter, row_ctx, env)?;
                    if !value.is_blank() {
                        return Ok(value);
                    }
                }
                Ok(Value::Blank)
            }
            "NOT" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("NOT expects 1 argument".into()));
                };
                let value = self.eval_scalar(model, arg, filter, row_ctx, env)?;
                let b = value.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                Ok(Value::Boolean(!b))
            }
            "AND" => {
                let [left, right] = args else {
                    return Err(DaxError::Eval("AND expects 2 arguments".into()));
                };
                let left = self.eval_scalar(model, left, filter, row_ctx, env)?;
                let right = self.eval_scalar(model, right, filter, row_ctx, env)?;
                let left = left.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                let right = right.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                Ok(Value::Boolean(left && right))
            }
            "OR" => {
                let [left, right] = args else {
                    return Err(DaxError::Eval("OR expects 2 arguments".into()));
                };
                let left = self.eval_scalar(model, left, filter, row_ctx, env)?;
                let right = self.eval_scalar(model, right, filter, row_ctx, env)?;
                let left = left.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                let right = right.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                Ok(Value::Boolean(left || right))
            }
            "DISTINCTCOUNT" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("DISTINCTCOUNT expects 1 argument".into()));
                };
                self.eval_distinctcount(model, arg, filter)
            }
            "DISTINCTCOUNTNOBLANK" => {
                let [arg] = args else {
                    return Err(DaxError::Eval(
                        "DISTINCTCOUNTNOBLANK expects 1 argument".into(),
                    ));
                };
                if !matches!(arg, Expr::ColumnRef { .. }) {
                    return Err(DaxError::Type(
                        "DISTINCTCOUNTNOBLANK expects a column reference".into(),
                    ));
                }
                self.eval_distinctcountnoblank(model, arg, filter)
            }
            "HASONEVALUE" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("HASONEVALUE expects 1 argument".into()));
                };
                let values = self.distinct_column_values(model, arg, filter)?;
                Ok(Value::Boolean(values.len() == 1))
            }
            "ISINSCOPE" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("ISINSCOPE expects 1 argument".into()));
                };
                let Expr::ColumnRef { table, column } = arg else {
                    return Err(DaxError::Type(
                        "ISINSCOPE expects a column reference".into(),
                    ));
                };
                Ok(Value::Boolean(filter.in_scope_columns.contains(&(
                    normalize_ident(table),
                    normalize_ident(column),
                ))))
            }
            "SELECTEDVALUE" => {
                if args.is_empty() || args.len() > 2 {
                    return Err(DaxError::Eval(
                        "SELECTEDVALUE expects 1 or 2 arguments".into(),
                    ));
                }
                let values = self.distinct_column_values(model, &args[0], filter)?;
                if values.len() == 1 {
                    Ok(values.into_iter().next().expect("len==1"))
                } else if args.len() == 2 {
                    self.eval_scalar(model, &args[1], filter, row_ctx, env)
                } else {
                    Ok(Value::Blank)
                }
            }
            "SUM" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("SUM expects 1 argument".into()));
                };
                self.eval_sum(model, arg, filter)
            }
            "AVERAGE" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("AVERAGE expects 1 argument".into()));
                };
                self.eval_average(model, arg, filter)
            }
            "MAX" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("MAX expects 1 argument".into()));
                };
                self.eval_max(model, arg, filter)
            }
            "MIN" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("MIN expects 1 argument".into()));
                };
                self.eval_min(model, arg, filter)
            }
            "COUNT" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("COUNT expects 1 argument".into()));
                };
                self.eval_count(model, arg, filter)
            }
            "COUNTA" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("COUNTA expects 1 argument".into()));
                };
                self.eval_counta(model, arg, filter)
            }
            "COUNTBLANK" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("COUNTBLANK expects 1 argument".into()));
                };
                self.eval_countblank(model, arg, filter)
            }
            "SUMX" => {
                let [table_expr, value_expr] = args else {
                    return Err(DaxError::Eval("SUMX expects 2 arguments".into()));
                };
                self.eval_iterator(
                    model,
                    table_expr,
                    value_expr,
                    filter,
                    row_ctx,
                    env,
                    IteratorKind::Sum,
                )
            }
            "AVERAGEX" => {
                let [table_expr, value_expr] = args else {
                    return Err(DaxError::Eval("AVERAGEX expects 2 arguments".into()));
                };
                self.eval_iterator(
                    model,
                    table_expr,
                    value_expr,
                    filter,
                    row_ctx,
                    env,
                    IteratorKind::Average,
                )
            }
            "MAXX" => {
                let [table_expr, value_expr] = args else {
                    return Err(DaxError::Eval("MAXX expects 2 arguments".into()));
                };
                self.eval_iterator(
                    model,
                    table_expr,
                    value_expr,
                    filter,
                    row_ctx,
                    env,
                    IteratorKind::Max,
                )
            }
            "MINX" => {
                let [table_expr, value_expr] = args else {
                    return Err(DaxError::Eval("MINX expects 2 arguments".into()));
                };
                self.eval_iterator(
                    model,
                    table_expr,
                    value_expr,
                    filter,
                    row_ctx,
                    env,
                    IteratorKind::Min,
                )
            }
            "COUNTROWS" => {
                let [table_expr] = args else {
                    return Err(DaxError::Eval("COUNTROWS expects 1 argument".into()));
                };
                let table_result = self.eval_table(model, table_expr, filter, row_ctx, env)?;
                Ok(Value::from(table_result.row_count() as i64))
            }
            "COUNTX" => {
                let [table_expr, value_expr] = args else {
                    return Err(DaxError::Eval("COUNTX expects 2 arguments".into()));
                };
                self.eval_iterator(
                    model,
                    table_expr,
                    value_expr,
                    filter,
                    row_ctx,
                    env,
                    IteratorKind::Count,
                )
            }
            "LOOKUPVALUE" => {
                if args.len() < 3 {
                    return Err(DaxError::Eval(
                        "LOOKUPVALUE expects at least 3 arguments".into(),
                    ));
                }

                let Expr::ColumnRef {
                    table: result_table,
                    column: result_column,
                } = &args[0]
                else {
                    return Err(DaxError::Type(
                        "LOOKUPVALUE expects a column reference as the first argument".into(),
                    ));
                };

                // DAX allows an optional final alternate result. With only (col, value) pairs,
                // LOOKUPVALUE has an odd argument count:
                //   1 (result column) + 2 * N (search pairs)
                // Adding alternate result makes the total even.
                let (search_args, alternate_result) = if args.len() % 2 == 0 {
                    (&args[1..args.len() - 1], Some(&args[args.len() - 1]))
                } else {
                    (&args[1..], None)
                };

                if search_args.len() < 2 || search_args.len() % 2 != 0 {
                    return Err(DaxError::Eval(
                        "LOOKUPVALUE expects at least one (search_column, search_value) pair"
                            .into(),
                    ));
                }

                // Resolve the search table / columns up front (MVP restriction: all search columns
                // must be in the same table as the result column).
                let table_ref = model
                    .table(result_table)
                    .ok_or_else(|| DaxError::UnknownTable(result_table.clone()))?;
                let result_table_key = normalize_ident(result_table);
                let result_idx =
                    table_ref
                        .column_idx(result_column)
                        .ok_or_else(|| DaxError::UnknownColumn {
                            table: result_table.clone(),
                            column: result_column.clone(),
                        })?;

                let mut search_cols: Vec<usize> = Vec::with_capacity(search_args.len() / 2);
                let mut search_values: Vec<Value> = Vec::with_capacity(search_args.len() / 2);
                for pair in search_args.chunks(2) {
                    let [search_col_expr, search_value_expr] = pair else {
                        unreachable!("validated even number of search args");
                    };

                    let Expr::ColumnRef {
                        table: search_table,
                        column: search_column,
                    } = search_col_expr
                    else {
                        return Err(DaxError::Type(
                            "LOOKUPVALUE expects search columns to be column references".into(),
                        ));
                    };

                    if normalize_ident(search_table) != result_table_key {
                        return Err(DaxError::Eval(
                            "LOOKUPVALUE MVP requires all search columns to be in the same table as the result column".into(),
                        ));
                    }

                    let search_idx = table_ref.column_idx(search_column).ok_or_else(|| {
                        DaxError::UnknownColumn {
                            table: search_table.clone(),
                            column: search_column.clone(),
                        }
                    })?;
                    search_cols.push(search_idx);

                    let search_value =
                        self.eval_scalar(model, search_value_expr, filter, row_ctx, env)?;
                    search_values.push(search_value);
                }

                // Scan visible rows under the current filter context and apply the search
                // conditions.
                let mut matched_rows = Vec::new();
                let row_sets = (!filter.is_empty())
                    .then(|| resolve_row_sets(model, filter))
                    .transpose()?;
                if let Some(sets) = row_sets.as_ref() {
                    let allowed = sets
                        .get(&result_table_key)
                        .ok_or_else(|| DaxError::UnknownTable(result_table.clone()))?;
                    for row in allowed.iter_ones() {
                        let mut matches = true;
                        for (col_idx, search_value) in search_cols.iter().zip(search_values.iter())
                        {
                            let cell_value = table_ref
                                .value_by_idx(row, *col_idx)
                                .unwrap_or(Value::Blank);
                            if !compare_values(&BinaryOp::Equals, &cell_value, search_value)? {
                                matches = false;
                                break;
                            }
                        }
                        if matches {
                            matched_rows.push(row);
                        }
                    }
                } else {
                    for row in 0..table_ref.row_count() {
                        let mut matches = true;
                        for (col_idx, search_value) in search_cols.iter().zip(search_values.iter())
                        {
                            let cell_value = table_ref
                                .value_by_idx(row, *col_idx)
                                .unwrap_or(Value::Blank);
                            if !compare_values(&BinaryOp::Equals, &cell_value, search_value)? {
                                matches = false;
                                break;
                            }
                        }
                        if matches {
                            matched_rows.push(row);
                        }
                    }
                }

                match matched_rows.len() {
                    0 => {
                        if let Some(expr) = alternate_result {
                            self.eval_scalar(model, expr, filter, row_ctx, env)
                        } else {
                            Ok(Value::Blank)
                        }
                    }
                    1 => Ok(table_ref
                        .value_by_idx(matched_rows[0], result_idx)
                        .unwrap_or(Value::Blank)),
                    _ => {
                        // DAX: allow duplicates only when the result values are unambiguous.
                        let mut non_blank: Option<Value> = None;
                        for &row in &matched_rows {
                            let value = table_ref
                                .value_by_idx(row, result_idx)
                                .unwrap_or(Value::Blank);
                            if value.is_blank() {
                                continue;
                            }
                            if let Some(existing) = &non_blank {
                                if existing != &value {
                                    return Err(DaxError::Eval(format!(
                                        "LOOKUPVALUE found multiple values for {result_table}[{result_column}]"
                                    )));
                                }
                            } else {
                                non_blank = Some(value);
                            }
                        }

                        Ok(non_blank.unwrap_or(Value::Blank))
                    }
                }
            }
            "CONCATENATEX" => {
                if args.len() < 2 || args.len() > 5 {
                    return Err(DaxError::Eval(
                        "CONCATENATEX expects 2 to 5 arguments".into(),
                    ));
                }
                let table_expr = &args[0];
                let text_expr = &args[1];

                // The delimiter is evaluated once in the outer context (matching DAX behavior).
                // Default delimiter is an empty string.
                let delimiter = if args.len() >= 3 {
                    let v = self.eval_scalar(model, &args[2], filter, row_ctx, env)?;
                    coerce_text(&v).into_owned()
                } else {
                    String::new()
                };

                let table_result = self.eval_table(model, table_expr, filter, row_ctx, env)?;

                let descending = if args.len() >= 5 {
                    let order_arg = &args[4];
                    let order = match order_arg {
                        // DAX passes ASC/DESC as bare identifiers, which we parse as `TableName`.
                        Expr::TableName(name) => name.clone(),
                        _ => {
                            let v = self.eval_scalar(model, order_arg, filter, row_ctx, env)?;
                            coerce_text(&v).into_owned()
                        }
                    };
                    match order.to_ascii_uppercase().as_str() {
                        "ASC" => false,
                        "DESC" => true,
                        other => {
                            return Err(DaxError::Eval(format!(
                                "CONCATENATEX order must be ASC or DESC, got {other}"
                            )));
                        }
                    }
                } else {
                    false
                };

                let mut out = String::new();
                let mut first = true;
                // If an order-by expression is provided, precompute both the sort key and the
                // formatted text for each row, then stable-sort before joining.
                if args.len() >= 4 {
                    let order_by_expr = &args[3];
                    let mut keyed: Vec<(Value, String)> =
                        Vec::with_capacity(table_result.row_count());
                    let mut saw_text = false;

                    for row in table_result.iter_rows() {
                        let inner_ctx = table_result.push_row_ctx(row_ctx, row);
                        let value = self.eval_scalar(model, text_expr, filter, &inner_ctx, env)?;
                        let text = coerce_text(&value).into_owned();

                        let key =
                            self.eval_scalar(model, order_by_expr, filter, &inner_ctx, env)?;
                        if matches!(&key, Value::Text(_)) {
                            saw_text = true;
                        }
                        keyed.push((key, text));
                    }

                    if saw_text {
                        let mut items: Vec<(String, String)> = keyed
                            .into_iter()
                            .map(|(key, text)| (coerce_text(&key).into_owned(), text))
                            .collect();
                        items.sort_by(|a, b| {
                            // Match Excel-like case-insensitive text ordering, with a deterministic
                            // case-sensitive tiebreak (so ordering remains total).
                            let ord = cmp_text_case_insensitive(&a.0, &b.0);
                            let ord = if ord != Ordering::Equal {
                                ord
                            } else {
                                a.0.cmp(&b.0)
                            };
                            if descending {
                                ord.reverse()
                            } else {
                                ord
                            }
                        });

                        for (_key, text) in items {
                            if !first {
                                out.push_str(&delimiter);
                            }
                            out.push_str(&text);
                            first = false;
                        }
                    } else {
                        let mut items: Vec<(OrderedFloat<f64>, String)> =
                            Vec::with_capacity(keyed.len());
                        for (key, text) in keyed {
                            let n = coerce_number(&key)?;
                            if !n.is_finite() {
                                return Err(DaxError::Eval(
                                    "CONCATENATEX order_by_expr must return a finite number".into(),
                                ));
                            }
                            items.push((OrderedFloat(n), text));
                        }
                        if descending {
                            items.sort_by(|a, b| b.0.cmp(&a.0));
                        } else {
                            items.sort_by(|a, b| a.0.cmp(&b.0));
                        }

                        for (_key, text) in items {
                            if !first {
                                out.push_str(&delimiter);
                            }
                            out.push_str(&text);
                            first = false;
                        }
                    }
                } else {
                    for row in table_result.iter_rows() {
                        let inner_ctx = table_result.push_row_ctx(row_ctx, row);
                        let value = self.eval_scalar(model, text_expr, filter, &inner_ctx, env)?;
                        let text = coerce_text(&value);
                        if !first {
                            out.push_str(&delimiter);
                        }
                        out.push_str(&text);
                        first = false;
                    }
                }
                Ok(Value::from(out))
            }
            "CALCULATE" => {
                if args.is_empty() {
                    return Err(DaxError::Eval(
                        "CALCULATE expects at least 1 argument".into(),
                    ));
                }
                self.eval_calculate(model, args, filter, row_ctx, env)
            }
            "RELATED" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("RELATED expects 1 argument".into()));
                };
                self.eval_related(model, arg, filter, row_ctx)
            }
            "CONTAINSROW" => {
                if args.len() < 2 {
                    return Err(DaxError::Eval(
                        "CONTAINSROW expects at least 2 arguments".into(),
                    ));
                }
                let (table_expr, value_exprs) = args.split_first().expect("checked above");

                let table_result = self.eval_table(model, table_expr, filter, row_ctx, env)?;
                match table_result {
                    TableResult::Physical {
                        table,
                        rows,
                        visible_cols,
                    } => {
                        let table_ref = model
                            .table(&table)
                            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;

                        let visible_cols = visible_cols.as_deref();
                        let col_count = visible_cols
                            .map(|cols| cols.len())
                            .unwrap_or_else(|| table_ref.columns().len());

                        if col_count != value_exprs.len() {
                            return Err(DaxError::Eval(format!(
                                "CONTAINSROW expected {} value arguments, got {}",
                                col_count,
                                value_exprs.len()
                            )));
                        }

                        let needles: Vec<Value> = value_exprs
                            .iter()
                            .map(|expr| self.eval_scalar(model, expr, filter, row_ctx, env))
                            .collect::<DaxResult<_>>()?;

                        if let Some(cols) = visible_cols {
                            for row in rows {
                                let mut matches = true;
                                for (&col_idx, needle) in cols.iter().zip(&needles) {
                                    let value = table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank);
                                    if !compare_values(&BinaryOp::Equals, &value, needle)? {
                                        matches = false;
                                        break;
                                    }
                                }
                                if matches {
                                    return Ok(Value::Boolean(true));
                                }
                            }
                        } else {
                            for row in rows {
                                let mut matches = true;
                                for (col_idx, needle) in needles.iter().enumerate() {
                                    let value = table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank);
                                    if !compare_values(&BinaryOp::Equals, &value, needle)? {
                                        matches = false;
                                        break;
                                    }
                                }
                                if matches {
                                    return Ok(Value::Boolean(true));
                                }
                            }
                        }

                        Ok(Value::Boolean(false))
                    }
                    TableResult::PhysicalAll {
                        table,
                        row_count,
                        visible_cols,
                    } => {
                        let table_ref = model
                            .table(&table)
                            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;

                        let visible_cols = visible_cols.as_deref();
                        let col_count = visible_cols
                            .map(|cols| cols.len())
                            .unwrap_or_else(|| table_ref.columns().len());

                        if col_count != value_exprs.len() {
                            return Err(DaxError::Eval(format!(
                                "CONTAINSROW expected {} value arguments, got {}",
                                col_count,
                                value_exprs.len()
                            )));
                        }

                        let needles: Vec<Value> = value_exprs
                            .iter()
                            .map(|expr| self.eval_scalar(model, expr, filter, row_ctx, env))
                            .collect::<DaxResult<_>>()?;

                        if let Some(cols) = visible_cols {
                            for row in 0..row_count {
                                let mut matches = true;
                                for (&col_idx, needle) in cols.iter().zip(&needles) {
                                    let value = table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank);
                                    if !compare_values(&BinaryOp::Equals, &value, needle)? {
                                        matches = false;
                                        break;
                                    }
                                }
                                if matches {
                                    return Ok(Value::Boolean(true));
                                }
                            }
                        } else {
                            for row in 0..row_count {
                                let mut matches = true;
                                for (col_idx, needle) in needles.iter().enumerate() {
                                    let value = table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank);
                                    if !compare_values(&BinaryOp::Equals, &value, needle)? {
                                        matches = false;
                                        break;
                                    }
                                }
                                if matches {
                                    return Ok(Value::Boolean(true));
                                }
                            }
                        }

                        Ok(Value::Boolean(false))
                    }
                    TableResult::PhysicalMask {
                        table,
                        mask,
                        visible_cols,
                    } => {
                        let table_ref = model
                            .table(&table)
                            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;

                        let visible_cols = visible_cols.as_deref();
                        let col_count = visible_cols
                            .map(|cols| cols.len())
                            .unwrap_or_else(|| table_ref.columns().len());

                        if col_count != value_exprs.len() {
                            return Err(DaxError::Eval(format!(
                                "CONTAINSROW expected {} value arguments, got {}",
                                col_count,
                                value_exprs.len()
                            )));
                        }

                        let needles: Vec<Value> = value_exprs
                            .iter()
                            .map(|expr| self.eval_scalar(model, expr, filter, row_ctx, env))
                            .collect::<DaxResult<_>>()?;

                        if let Some(cols) = visible_cols {
                            for row in mask.iter_ones() {
                                let mut matches = true;
                                for (&col_idx, needle) in cols.iter().zip(&needles) {
                                    let value = table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank);
                                    if !compare_values(&BinaryOp::Equals, &value, needle)? {
                                        matches = false;
                                        break;
                                    }
                                }
                                if matches {
                                    return Ok(Value::Boolean(true));
                                }
                            }
                        } else {
                            for row in mask.iter_ones() {
                                let mut matches = true;
                                for (col_idx, needle) in needles.iter().enumerate() {
                                    let value = table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank);
                                    if !compare_values(&BinaryOp::Equals, &value, needle)? {
                                        matches = false;
                                        break;
                                    }
                                }
                                if matches {
                                    return Ok(Value::Boolean(true));
                                }
                            }
                        }

                        Ok(Value::Boolean(false))
                    }
                    TableResult::Virtual { columns, rows } => {
                        if columns.len() != value_exprs.len() {
                            return Err(DaxError::Eval(format!(
                                "CONTAINSROW expected {} value arguments, got {}",
                                columns.len(),
                                value_exprs.len()
                            )));
                        }
                        let needles: Vec<Value> = value_exprs
                            .iter()
                            .map(|expr| self.eval_scalar(model, expr, filter, row_ctx, env))
                            .collect::<DaxResult<_>>()?;

                        for row_values in rows {
                            let mut matches = true;
                            for (i, needle) in needles.iter().enumerate() {
                                let value = row_values.get(i).cloned().unwrap_or(Value::Blank);
                                if !compare_values(&BinaryOp::Equals, &value, needle)? {
                                    matches = false;
                                    break;
                                }
                            }
                            if matches {
                                return Ok(Value::Boolean(true));
                            }
                        }
                        Ok(Value::Boolean(false))
                    }
                }
            }
            "EARLIER" => {
                if args.is_empty() || args.len() > 2 {
                    return Err(DaxError::Eval("EARLIER expects 1 or 2 arguments".into()));
                }

                let Expr::ColumnRef { table, column } = &args[0] else {
                    return Err(DaxError::Type(
                        "EARLIER expects a column reference as the first argument".into(),
                    ));
                };

                let level_from_inner: usize = if args.len() == 2 {
                    let value = self.eval_scalar(model, &args[1], filter, row_ctx, env)?;
                    let n = coerce_number(&value)?;
                    if !n.is_finite() {
                        return Err(DaxError::Eval(
                            "EARLIER expects a finite number for the optional second argument"
                                .into(),
                        ));
                    }
                    let n = n.trunc() as i64;
                    if n < 1 {
                        return Err(DaxError::Eval(
                            "EARLIER expects the optional second argument to be >= 1".into(),
                        ));
                    }
                    n as usize
                } else {
                    1
                };

                let Some((row, visible_cols)) =
                    row_ctx.physical_row_for_level(table, level_from_inner)
                else {
                    let table_key = normalize_ident(table);
                    let available = row_ctx
                        .stack
                        .iter()
                        .filter(|frame| {
                            matches!(frame, RowContextFrame::Physical { table: t, .. } if t == &table_key)
                        })
                        .count();
                    return Err(DaxError::Eval(format!(
                        "EARLIER refers to an outer row context that does not exist for {table}[{column}] (requested level {level_from_inner}, available {available})"
                        )));
                };

                let table_ref = model
                    .table(table)
                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                let idx = table_ref
                    .column_idx(column)
                    .ok_or_else(|| DaxError::UnknownColumn {
                        table: table.clone(),
                        column: column.clone(),
                    })?;
                if let Some(visible_cols) = visible_cols {
                    if !visible_cols.contains(&idx) {
                        return Err(DaxError::Eval(format!(
                            "column {table}[{column}] is not available in the current row context"
                        )));
                    }
                }
                if row >= table_ref.row_count() {
                    return Ok(Value::Blank);
                }
                Ok(table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank))
            }
            "EARLIEST" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("EARLIEST expects 1 argument".into()));
                };
                let Expr::ColumnRef { table, column } = arg else {
                    return Err(DaxError::Type(
                        "EARLIEST expects a column reference as the first argument".into(),
                    ));
                };

                let Some((row, visible_cols)) = row_ctx.physical_row_for_outermost(table) else {
                    return Err(DaxError::Eval(format!(
                        "EARLIEST requires row context for {table}[{column}]"
                    )));
                };

                let table_ref = model
                    .table(table)
                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                let idx = table_ref
                    .column_idx(column)
                    .ok_or_else(|| DaxError::UnknownColumn {
                        table: table.clone(),
                        column: column.clone(),
                    })?;
                if let Some(visible_cols) = visible_cols {
                    if !visible_cols.contains(&idx) {
                        return Err(DaxError::Eval(format!(
                            "column {table}[{column}] is not available in the current row context"
                        )));
                    }
                }
                if row >= table_ref.row_count() {
                    return Ok(Value::Blank);
                }
                Ok(table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank))
            }
            other => Err(DaxError::Eval(format!("unsupported function {other}"))),
        }
    }

    fn eval_sum(&self, model: &DataModel, expr: &Expr, filter: &FilterContext) -> DaxResult<Value> {
        let (table, column) = match expr {
            Expr::ColumnRef { table, column } => (table.as_str(), column.as_str()),
            _ => {
                return Err(DaxError::Type(
                    "SUM currently only supports a column reference".into(),
                ))
            }
        };

        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.into()))?;
        let idx = table_ref
            .column_idx(column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.to_string(),
                column: column.to_string(),
            })?;

        if filter.is_empty() {
            if let (Some(sum), Some(count)) = (
                table_ref.stats_sum(idx),
                table_ref.stats_non_blank_count(idx),
            ) {
                return Ok(if count == 0 {
                    Value::Blank
                } else {
                    Value::from(sum)
                });
            }
        }

        let mut sum = 0.0;
        let mut count = 0usize;
        let row_sets = (!filter.is_empty())
            .then(|| resolve_row_sets(model, filter))
            .transpose()?;
        let table_key = normalize_ident(table);
        if let Some(sets) = row_sets.as_ref() {
            let allowed = sets
                .get(&table_key)
                .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
            for row in allowed.iter_ones() {
                if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                    sum += n.0;
                    count += 1;
                }
            }
        } else {
            for row in 0..table_ref.row_count() {
                if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                    sum += n.0;
                    count += 1;
                }
            }
        }
        if count == 0 {
            Ok(Value::Blank)
        } else {
            Ok(Value::from(sum))
        }
    }

    fn eval_average(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
    ) -> DaxResult<Value> {
        let (table, column) = match expr {
            Expr::ColumnRef { table, column } => (table.as_str(), column.as_str()),
            _ => {
                return Err(DaxError::Type(
                    "AVERAGE currently only supports a column reference".into(),
                ))
            }
        };

        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.into()))?;
        let idx = table_ref
            .column_idx(column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.to_string(),
                column: column.to_string(),
            })?;

        if filter.is_empty() {
            if let (Some(sum), Some(count)) = (
                table_ref.stats_sum(idx),
                table_ref.stats_non_blank_count(idx),
            ) {
                return Ok(if count == 0 {
                    Value::Blank
                } else {
                    Value::from(sum / count as f64)
                });
            }
        }

        let mut sum = 0.0;
        let mut count = 0usize;
        let row_sets = (!filter.is_empty())
            .then(|| resolve_row_sets(model, filter))
            .transpose()?;
        let table_key = normalize_ident(table);
        if let Some(sets) = row_sets.as_ref() {
            let allowed = sets
                .get(&table_key)
                .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
            for row in allowed.iter_ones() {
                if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                    sum += n.0;
                    count += 1;
                }
            }
        } else {
            for row in 0..table_ref.row_count() {
                if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                    sum += n.0;
                    count += 1;
                }
            }
        }
        if count == 0 {
            Ok(Value::Blank)
        } else {
            Ok(Value::from(sum / count as f64))
        }
    }

    fn eval_max(&self, model: &DataModel, expr: &Expr, filter: &FilterContext) -> DaxResult<Value> {
        let (table, column) = match expr {
            Expr::ColumnRef { table, column } => (table.as_str(), column.as_str()),
            _ => {
                return Err(DaxError::Type(
                    "MAX currently only supports a column reference".into(),
                ))
            }
        };

        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.into()))?;
        let idx = table_ref
            .column_idx(column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.to_string(),
                column: column.to_string(),
            })?;

        if filter.is_empty() {
            if let Some(v @ Value::Number(_)) = table_ref.stats_max(idx) {
                return Ok(v);
            }
        }

        let mut best: Option<f64> = None;
        let row_sets = (!filter.is_empty())
            .then(|| resolve_row_sets(model, filter))
            .transpose()?;
        let table_key = normalize_ident(table);
        if let Some(sets) = row_sets.as_ref() {
            let allowed = sets
                .get(&table_key)
                .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
            for row in allowed.iter_ones() {
                if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                    best = Some(best.map_or(n.0, |current| current.max(n.0)));
                }
            }
        } else {
            for row in 0..table_ref.row_count() {
                if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                    best = Some(best.map_or(n.0, |current| current.max(n.0)));
                }
            }
        }
        Ok(best.map(Value::from).unwrap_or(Value::Blank))
    }

    fn eval_min(&self, model: &DataModel, expr: &Expr, filter: &FilterContext) -> DaxResult<Value> {
        let (table, column) = match expr {
            Expr::ColumnRef { table, column } => (table.as_str(), column.as_str()),
            _ => {
                return Err(DaxError::Type(
                    "MIN currently only supports a column reference".into(),
                ))
            }
        };

        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.into()))?;
        let idx = table_ref
            .column_idx(column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.to_string(),
                column: column.to_string(),
            })?;

        if filter.is_empty() {
            if let Some(v @ Value::Number(_)) = table_ref.stats_min(idx) {
                return Ok(v);
            }
        }

        let mut best: Option<f64> = None;
        let row_sets = (!filter.is_empty())
            .then(|| resolve_row_sets(model, filter))
            .transpose()?;
        let table_key = normalize_ident(table);
        if let Some(sets) = row_sets.as_ref() {
            let allowed = sets
                .get(&table_key)
                .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
            for row in allowed.iter_ones() {
                if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                    best = Some(best.map_or(n.0, |current| current.min(n.0)));
                }
            }
        } else {
            for row in 0..table_ref.row_count() {
                if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                    best = Some(best.map_or(n.0, |current| current.min(n.0)));
                }
            }
        }
        Ok(best.map(Value::from).unwrap_or(Value::Blank))
    }

    fn eval_count(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
    ) -> DaxResult<Value> {
        let (table, column) = match expr {
            Expr::ColumnRef { table, column } => (table.as_str(), column.as_str()),
            _ => {
                return Err(DaxError::Type(
                    "COUNT currently only supports a column reference".into(),
                ))
            }
        };

        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.into()))?;
        let idx = table_ref
            .column_idx(column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.to_string(),
                column: column.to_string(),
            })?;

        // Fast path: for columnar numeric columns, COUNT == non-blank count.
        if filter.is_empty() {
            if let (Some(non_blank), Some(is_numeric)) = (
                table_ref.stats_non_blank_count(idx),
                column_is_dax_numeric(table_ref, idx),
            ) {
                return Ok(Value::from(if is_numeric { non_blank as i64 } else { 0 }));
            }
        }

        let mut count = 0usize;
        let row_sets = (!filter.is_empty())
            .then(|| resolve_row_sets(model, filter))
            .transpose()?;
        let table_key = normalize_ident(table);
        if let Some(sets) = row_sets.as_ref() {
            let allowed = sets
                .get(&table_key)
                .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
            for row in allowed.iter_ones() {
                if matches!(table_ref.value_by_idx(row, idx), Some(Value::Number(_))) {
                    count += 1;
                }
            }
        } else {
            for row in 0..table_ref.row_count() {
                if matches!(table_ref.value_by_idx(row, idx), Some(Value::Number(_))) {
                    count += 1;
                }
            }
        }

        Ok(Value::from(count as i64))
    }

    fn eval_counta(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
    ) -> DaxResult<Value> {
        let (table, column) = match expr {
            Expr::ColumnRef { table, column } => (table.as_str(), column.as_str()),
            _ => {
                return Err(DaxError::Type(
                    "COUNTA currently only supports a column reference".into(),
                ))
            }
        };
        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.into()))?;
        let idx = table_ref
            .column_idx(column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.to_string(),
                column: column.to_string(),
            })?;

        if filter.is_empty() {
            if let Some(non_blank) = table_ref.stats_non_blank_count(idx) {
                return Ok(Value::from(non_blank as i64));
            }
        }

        let mut count = 0usize;
        let row_sets = (!filter.is_empty())
            .then(|| resolve_row_sets(model, filter))
            .transpose()?;
        let table_key = normalize_ident(table);
        if let Some(sets) = row_sets.as_ref() {
            let allowed = sets
                .get(&table_key)
                .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
            for row in allowed.iter_ones() {
                if !table_ref
                    .value_by_idx(row, idx)
                    .unwrap_or(Value::Blank)
                    .is_blank()
                {
                    count += 1;
                }
            }
        } else {
            for row in 0..table_ref.row_count() {
                if !table_ref
                    .value_by_idx(row, idx)
                    .unwrap_or(Value::Blank)
                    .is_blank()
                {
                    count += 1;
                }
            }
        }

        Ok(Value::from(count as i64))
    }

    fn eval_countblank(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
    ) -> DaxResult<Value> {
        let (table, column) = match expr {
            Expr::ColumnRef { table, column } => (table.as_str(), column.as_str()),
            _ => {
                return Err(DaxError::Type(
                    "COUNTBLANK currently only supports a column reference".into(),
                ))
            }
        };
        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.into()))?;
        let idx = table_ref
            .column_idx(column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.to_string(),
                column: column.to_string(),
            })?;

        if filter.is_empty() {
            let include_virtual_blank = blank_row_allowed(filter, table)
                && virtual_blank_row_exists(model, filter, table, None)?;
            if let Some(non_blank) = table_ref.stats_non_blank_count(idx) {
                let mut blanks = table_ref.row_count().saturating_sub(non_blank);
                if include_virtual_blank {
                    blanks += 1;
                }
                return Ok(Value::from(blanks as i64));
            }

            let mut blanks = 0usize;
            for row in 0..table_ref.row_count() {
                if table_ref
                    .value_by_idx(row, idx)
                    .unwrap_or(Value::Blank)
                    .is_blank()
                {
                    blanks += 1;
                }
            }
            if include_virtual_blank {
                blanks += 1;
            }
            return Ok(Value::from(blanks as i64));
        }

        let sets = resolve_row_sets(model, filter)?;
        let table_key = normalize_ident(table);
        let Some(rows_set) = sets.get(&table_key) else {
            return Err(DaxError::UnknownTable(table.to_string()));
        };

        let include_virtual_blank = blank_row_allowed(filter, table)
            && virtual_blank_row_exists(model, filter, table, Some(&sets))?;

        let mut blanks = 0usize;
        for row in rows_set.iter_ones() {
            if table_ref
                .value_by_idx(row, idx)
                .unwrap_or(Value::Blank)
                .is_blank()
            {
                blanks += 1;
            }
        }
        if include_virtual_blank {
            blanks += 1;
        }
        Ok(Value::from(blanks as i64))
    }

    fn eval_iterator(
        &self,
        model: &DataModel,
        table_expr: &Expr,
        value_expr: &Expr,
        filter: &FilterContext,
        row_ctx: &RowContext,
        env: &mut VarEnv,
        kind: IteratorKind,
    ) -> DaxResult<Value> {
        let table_result = self.eval_table(model, table_expr, filter, row_ctx, env)?;
        let mut sum = 0.0;
        let mut count = 0usize;
        let mut best: Option<f64> = None;

        for row in table_result.iter_rows() {
            let inner_ctx = table_result.push_row_ctx(row_ctx, row);
            let value = self.eval_scalar(model, value_expr, filter, &inner_ctx, env)?;
            match kind {
                IteratorKind::Sum | IteratorKind::Average => match value {
                    Value::Number(n) => {
                        sum += n.0;
                        count += 1;
                    }
                    Value::Blank => {}
                    other => {
                        return Err(DaxError::Type(format!(
                            "iterator expected numeric expression, got {other}"
                        )))
                    }
                },
                IteratorKind::Count => {
                    if !value.is_blank() {
                        count += 1;
                    }
                }
                IteratorKind::Max | IteratorKind::Min => match value {
                    Value::Number(n) => {
                        best = Some(match (kind, best) {
                            (IteratorKind::Max, Some(current)) => current.max(n.0),
                            (IteratorKind::Min, Some(current)) => current.min(n.0),
                            (_, None) => n.0,
                            _ => unreachable!(),
                        });
                        count += 1;
                    }
                    Value::Blank => {}
                    other => {
                        return Err(DaxError::Type(format!(
                            "iterator expected numeric expression, got {other}"
                        )))
                    }
                },
            };
        }

        match kind {
            IteratorKind::Sum => {
                if count == 0 {
                    Ok(Value::Blank)
                } else {
                    Ok(Value::from(sum))
                }
            }
            IteratorKind::Average => {
                if count == 0 {
                    Ok(Value::Blank)
                } else {
                    Ok(Value::from(sum / count as f64))
                }
            }
            IteratorKind::Count => Ok(Value::from(count as i64)),
            IteratorKind::Max | IteratorKind::Min => {
                Ok(best.map(Value::from).unwrap_or(Value::Blank))
            }
        }
    }

    fn eval_distinctcount(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
    ) -> DaxResult<Value> {
        if filter.is_empty() {
            if let Expr::ColumnRef { table, column } = expr {
                let table_ref = model
                    .table(table)
                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                let idx = table_ref
                    .column_idx(column)
                    .ok_or_else(|| DaxError::UnknownColumn {
                        table: table.clone(),
                        column: column.clone(),
                    })?;
                if let Some(distinct) = table_ref.stats_distinct_count(idx) {
                    let mut out = distinct as i64;
                    let has_blank = table_ref.stats_has_blank(idx).unwrap_or(false);
                    if has_blank {
                        out += 1;
                    } else if blank_row_allowed(filter, table)
                        && virtual_blank_row_exists(model, filter, table, None)?
                    {
                        out += 1;
                    }
                    return Ok(Value::from(out));
                }
            }
        }

        let values = self.distinct_column_values(model, expr, filter)?;
        Ok(Value::from(values.len() as i64))
    }

    fn eval_distinctcountnoblank(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
    ) -> DaxResult<Value> {
        let Expr::ColumnRef { table, column } = expr else {
            return Err(DaxError::Type("expected a column reference".to_string()));
        };

        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
        let idx = table_ref
            .column_idx(column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.clone(),
                column: column.clone(),
            })?;

        // Fast path: precomputed distinct count (excluding blanks) when unfiltered.
        if filter.is_empty() {
            if let Some(distinct) = table_ref.stats_distinct_count(idx) {
                return Ok(Value::from(distinct as i64));
            }

            // Fallback: dictionary/group-by distinct values (may include BLANK; filter it out).
            if let Some(values) = table_ref.distinct_values_filtered(idx, None) {
                let count = values.iter().filter(|v| !v.is_blank()).count();
                return Ok(Value::from(count as i64));
            }
        }

        let row_sets = (!filter.is_empty())
            .then(|| resolve_row_sets(model, filter))
            .transpose()?;
        let table_key = normalize_ident(table);
        let mask = if let Some(sets) = row_sets.as_ref() {
            Some(
                sets.get(&table_key)
                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?,
            )
        } else {
            None
        };

        if let Some(values) = table_ref.distinct_values_filtered_mask(idx, mask) {
            let count = values.iter().filter(|v| !v.is_blank()).count();
            return Ok(Value::from(count as i64));
        }

        let mut out = HashSet::new();
        if let Some(mask) = mask {
            for row in mask.iter_ones() {
                let value = table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                if !value.is_blank() {
                    out.insert(value);
                }
            }
        } else {
            for row in 0..table_ref.row_count() {
                let value = table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                if !value.is_blank() {
                    out.insert(value);
                }
            }
        }

        Ok(Value::from(out.len() as i64))
    }

    /// Returns the set of distinct values visible for a column under the provided filter context.
    ///
    /// This includes the relationship-generated virtual BLANK member when applicable (matching
    /// `VALUES(Table[Column])` semantics).
    pub fn distinct_column_values(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
    ) -> DaxResult<HashSet<Value>> {
        let Expr::ColumnRef { table, column } = expr else {
            return Err(DaxError::Type("expected a column reference".to_string()));
        };

        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
        let idx = table_ref
            .column_idx(column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.clone(),
                column: column.clone(),
            })?;

        if filter.is_empty() {
            let include_virtual_blank = blank_row_allowed(filter, table)
                && virtual_blank_row_exists(model, filter, table, None)?;
            if let Some(values) = table_ref.distinct_values_filtered(idx, None) {
                let mut out: HashSet<Value> = values.into_iter().collect();
                if include_virtual_blank {
                    out.insert(Value::Blank);
                }
                return Ok(out);
            }
            let mut out = HashSet::new();
            for row in 0..table_ref.row_count() {
                let value = table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                out.insert(value);
            }
            if include_virtual_blank {
                out.insert(Value::Blank);
            }
            return Ok(out);
        }

        let sets = resolve_row_sets(model, filter)?;
        let table_key = normalize_ident(table);
        let Some(rows_set) = sets.get(&table_key) else {
            return Err(DaxError::UnknownTable(table.to_string()));
        };

        let include_virtual_blank = blank_row_allowed(filter, table)
            && virtual_blank_row_exists(model, filter, table, Some(&sets))?;

        if let Some(values) = table_ref.distinct_values_filtered_mask(idx, Some(rows_set)) {
            let mut out: HashSet<Value> = values.into_iter().collect();
            if include_virtual_blank {
                out.insert(Value::Blank);
            }
            return Ok(out);
        }

        let mut out = HashSet::new();
        for row in rows_set.iter_ones() {
            let value = table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
            out.insert(value);
        }
        if include_virtual_blank {
            out.insert(Value::Blank);
        }
        Ok(out)
    }

    fn eval_calculate(
        &self,
        model: &DataModel,
        args: &[Expr],
        filter: &FilterContext,
        row_ctx: &RowContext,
        env: &mut VarEnv,
    ) -> DaxResult<Value> {
        let (expr, filter_args) = args.split_first().expect("checked above");
        let new_filter = self.build_calculate_filter(model, filter, row_ctx, filter_args, env)?;
        let mut expr_filter = new_filter;
        // `CALCULATE` already performs context transition before evaluating the expression, so
        // measure references inside should not re-apply an implicit transition that would
        // undo filter modifiers like `ALL(...)`.
        expr_filter.suppress_implicit_measure_context_transition = true;
        self.eval_scalar(model, expr, &expr_filter, row_ctx, env)
    }

    fn build_calculate_filter(
        &self,
        model: &DataModel,
        filter: &FilterContext,
        row_ctx: &RowContext,
        filter_args: &[Expr],
        env: &mut VarEnv,
    ) -> DaxResult<FilterContext> {
        let mut base_filter = filter.clone();
        base_filter.suppress_implicit_measure_context_transition = false;
        let mut new_filter = self.apply_context_transition(model, &base_filter, row_ctx)?;
        self.apply_calculate_filter_args(model, &mut new_filter, row_ctx, filter_args, env)?;
        Ok(new_filter)
    }

    fn apply_context_transition(
        &self,
        model: &DataModel,
        filter: &FilterContext,
        row_ctx: &RowContext,
    ) -> DaxResult<FilterContext> {
        let mut new_filter = filter.clone();
        let mut seen_physical_tables: HashSet<&str> = HashSet::new();

        for frame in row_ctx.stack.iter().rev() {
            match frame {
                RowContextFrame::Virtual { bindings } => {
                    for ((table, column), value) in bindings {
                        // Virtual rows can contain columns that don't correspond to a physical model
                        // column (e.g. table constructors or computed columns from future
                        // `SELECTCOLUMNS`/`ADDCOLUMNS`). DAX context transition only adds filters
                        // for *real* model columns, so ignore bindings that don't resolve.
                        let Some(table_ref) = model.table(table) else {
                            continue;
                        };
                        if table_ref.column_idx(column).is_none() {
                            continue;
                        }
                        let key = (table.clone(), column.clone());
                        match new_filter.column_filters.get_mut(&key) {
                            Some(existing) => existing.retain(|v| v == value),
                            None => {
                                new_filter
                                    .column_filters
                                    .insert(key, HashSet::from([value.clone()]));
                            }
                        }
                    }
                }
                RowContextFrame::Physical {
                    table,
                    row,
                    visible_cols,
                } => {
                    // Nested row contexts for the same physical table should only use the most
                    // recent (innermost) row, matching DAX's "current row" semantics.
                    if !seen_physical_tables.insert(table.as_str()) {
                        continue;
                    }

                    let table_ref = model
                        .table(table)
                        .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;

                    if let Some(visible_cols) = visible_cols {
                        for &col_idx in visible_cols {
                            let Some(column) = table_ref.columns().get(col_idx) else {
                                return Err(DaxError::Eval(format!(
                                    "row context refers to out-of-bounds column index {col_idx} for table {table}"
                                )));
                            };
                            let value = table_ref
                                .value_by_idx(*row, col_idx)
                                .unwrap_or(Value::Blank);
                            let key = (table.clone(), normalize_ident(column));
                            match new_filter.column_filters.get_mut(&key) {
                                Some(existing) => existing.retain(|v| v == &value),
                                None => {
                                    new_filter
                                        .column_filters
                                        .insert(key, HashSet::from([value]));
                                }
                            }
                        }
                    } else {
                        for (col_idx, column) in table_ref.columns().iter().enumerate() {
                            let value = table_ref
                                .value_by_idx(*row, col_idx)
                                .unwrap_or(Value::Blank);
                            let key = (table.clone(), normalize_ident(column));
                            match new_filter.column_filters.get_mut(&key) {
                                Some(existing) => existing.retain(|v| v == &value),
                                None => {
                                    new_filter
                                        .column_filters
                                        .insert(key, HashSet::from([value]));
                                }
                            }
                        }
                    }
                }
            }
        }
        Ok(new_filter)
    }

    fn apply_calculate_filter_args(
        &self,
        model: &DataModel,
        filter: &mut FilterContext,
        row_ctx: &RowContext,
        filter_args: &[Expr],
        env: &mut VarEnv,
    ) -> DaxResult<()> {
        // `CALCULATE` filter arguments are order-independent. Evaluate all arguments in the
        // original filter context (after context transition), then apply their effects together.
        let mut eval_filter = filter.clone();

        // Relationship modifiers affect how other filters propagate.
        //
        // In DAX, `USERELATIONSHIP` / `CROSSFILTER` are *filter arguments* to CALCULATE, not
        // general-purpose scalar functions. We therefore only apply them when they appear as a
        // (possibly wrapped) top-level filter argument, not when nested inside other expressions
        // (e.g. inside a scalar subexpression like `Sales[Amount] > CALCULATE(..., USERELATIONSHIP(...))`).
        fn unwrap_filter_arg_for_modifiers<'a>(arg: &'a Expr) -> DaxResult<&'a Expr> {
            match arg {
                Expr::Call { name, args } if name.eq_ignore_ascii_case("KEEPFILTERS") => {
                    let [inner] = args.as_slice() else {
                        return Err(DaxError::Eval(
                            "KEEPFILTERS expects exactly 1 argument".into(),
                        ));
                    };
                    unwrap_filter_arg_for_modifiers(inner)
                }
                Expr::Let { body, .. } => unwrap_filter_arg_for_modifiers(body),
                _ => Ok(arg),
            }
        }

        let mut crossfilter_overrides: HashMap<usize, RelationshipOverride> = HashMap::new();

        for arg in filter_args {
            let arg = unwrap_filter_arg_for_modifiers(arg)?;
            if let Expr::Call { name, args } = arg {
                if name.eq_ignore_ascii_case("USERELATIONSHIP") {
                    self.apply_userelationship(model, filter, args)?;
                    self.apply_userelationship(model, &mut eval_filter, args)?;
                } else if name.eq_ignore_ascii_case("CROSSFILTER") {
                    let (rel_idx, override_dir) = self.apply_crossfilter(model, filter, args)?;
                    if let Some(existing) = crossfilter_overrides.get(&rel_idx) {
                        if *existing != override_dir {
                            let relationship = model
                                .relationships()
                                .get(rel_idx)
                                .expect("relationship index from find_relationship_index");
                            return Err(DaxError::Eval(format!(
                                "CALCULATE contains multiple conflicting CROSSFILTER modifiers for relationship {}",
                                relationship.rel.name
                            )));
                        }
                    } else {
                        crossfilter_overrides.insert(rel_idx, override_dir);
                    }
                    eval_filter
                        .cross_filter_overrides
                        .insert(rel_idx, override_dir);
                }
            }
        }

        let mut clear_tables: HashSet<String> = HashSet::new();
        let mut clear_columns: HashSet<(String, String)> = HashSet::new();
        let mut row_filters: Vec<(String, RowFilter)> = Vec::new();
        let mut column_filters: Vec<((String, String), HashSet<Value>)> = Vec::new();

        fn collect_column_refs(
            expr: &Expr,
            tables: &mut HashSet<String>,
            columns: &mut HashSet<(String, String)>,
        ) {
            match expr {
                Expr::Let { bindings, body } => {
                    for (_, binding_expr) in bindings {
                        collect_column_refs(binding_expr, tables, columns);
                    }
                    collect_column_refs(body, tables, columns);
                }
                Expr::TableLiteral { rows } => {
                    for row in rows {
                        for cell in row {
                            collect_column_refs(cell, tables, columns);
                        }
                    }
                }
                Expr::Tuple(values) => {
                    for value in values {
                        collect_column_refs(value, tables, columns);
                    }
                }
                Expr::ColumnRef { table, column } => {
                    tables.insert(normalize_ident(table));
                    columns.insert((normalize_ident(table), normalize_ident(column)));
                }
                Expr::UnaryOp { expr, .. } => collect_column_refs(expr, tables, columns),
                Expr::BinaryOp { left, right, .. } => {
                    collect_column_refs(left, tables, columns);
                    collect_column_refs(right, tables, columns);
                }
                Expr::Call { args, .. } => {
                    for arg in args {
                        collect_column_refs(arg, tables, columns);
                    }
                }
                _ => {}
            }
        }

        fn apply_boolean_filter_expr(
            engine: &DaxEngine,
            model: &DataModel,
            expr: &Expr,
            eval_filter: &FilterContext,
            row_ctx: &RowContext,
            env: &mut VarEnv,
            keep_filters: bool,
            clear_columns: &mut HashSet<(String, String)>,
            row_filters: &mut Vec<(String, RowFilter)>,
        ) -> DaxResult<()> {
            let mut referenced_tables: HashSet<String> = HashSet::new();
            let mut referenced_columns: HashSet<(String, String)> = HashSet::new();
            collect_column_refs(expr, &mut referenced_tables, &mut referenced_columns);

            let table = if referenced_tables.len() == 1 {
                referenced_tables.into_iter().next().expect("len==1")
            } else {
                let mut tables: Vec<String> = referenced_tables
                    .into_iter()
                    .map(|t| {
                        model
                            .table(&t)
                            .map(|table| table.name().to_string())
                            .unwrap_or(t)
                    })
                    .collect();
                tables.sort();
                return Err(DaxError::Eval(format!(
                    "CALCULATE boolean filter expression must reference columns from exactly one table, got {}",
                    if tables.is_empty() {
                        "no tables".to_string()
                    } else {
                        format!("tables: {}", tables.join(", "))
                    }
                )));
            };

            // Boolean filter arguments have replacement semantics for the columns they reference.
            // Evaluate the predicate over candidate rows with existing filters on those columns
            // removed.
            let mut base_filter = eval_filter.clone();
            for key in &referenced_columns {
                // Evaluate the boolean filter expression in the same context as an ordinary
                // CALCULATE filter argument (replacement semantics): ignore any existing filters
                // on the referenced columns when determining candidate rows.
                //
                // `KEEPFILTERS` changes *application* semantics (intersection vs replacement),
                // but should not change how the inner filter is evaluated.
                base_filter.column_filters.remove(key);
                if !keep_filters {
                    clear_columns.insert(key.clone());
                }
            }

            let table_ref = model
                .table(&table)
                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
            let row_count = table_ref.row_count();
            let sparse_to_dense_threshold = row_count / 64;

            // Fast path: boolean predicates that are true for *all* candidate rows should not
            // materialize a huge `HashSet<usize>` just so we can throw it away again when we
            // detect `allowed == candidates`.
            //
            // We implement this with a small-prefix buffer:
            // - Before we see the first failing row, we only buffer up to `threshold` passing
            //   row indices.
            // - If we see a failure after we've exceeded `threshold`, we switch directly to the
            //   dense bitmask representation, seeded from the candidate row mask (all-true when
            //   unfiltered, or the `resolve_row_sets` bitmask when filtered) and clear failing
            //   rows as we go.
            //
            // If the predicate never fails, we return `RowFilter::All`.
            let mut prefix_passes: Vec<usize> = Vec::new();
            let mut prefix_dense = false;
            let mut sparse_passes: Vec<usize> = Vec::new();
            enum DenseMode {
                CandidateMask(BitVec),
                PassMask(BitVec),
            }
            let mut dense: Option<DenseMode> = None;

            let mut any_fail = false;
            let row_sets = (!base_filter.is_empty())
                .then(|| resolve_row_sets(model, &base_filter))
                .transpose()?;
            if let Some(sets) = row_sets.as_ref() {
                let allowed = sets
                    .get(&table)
                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                for row in allowed.iter_ones() {
                    let mut inner_ctx = row_ctx.clone();
                    inner_ctx.push(&table, row);
                    let pred = engine.eval_scalar(model, expr, &base_filter, &inner_ctx, env)?;
                    let keep = pred.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                    match &mut dense {
                        Some(DenseMode::CandidateMask(mask)) => {
                            if !keep {
                                mask.set(row, false);
                            }
                        }
                        Some(DenseMode::PassMask(mask)) => {
                            if keep {
                                mask.set(row, true);
                            }
                        }
                        None => {
                            if keep {
                                if any_fail {
                                    sparse_passes.push(row);
                                    if sparse_passes.len() > sparse_to_dense_threshold {
                                        let mut mask = BitVec::with_len_all_false(row_count);
                                        for &row in &sparse_passes {
                                            mask.set(row, true);
                                        }
                                        sparse_passes.clear();
                                        dense = Some(DenseMode::PassMask(mask));
                                    }
                                } else if !prefix_dense {
                                    prefix_passes.push(row);
                                    if prefix_passes.len() > sparse_to_dense_threshold {
                                        prefix_dense = true;
                                        prefix_passes.clear();
                                    }
                                }
                            } else {
                                any_fail = true;
                                if prefix_dense {
                                    let mut mask = allowed.clone();
                                    mask.set(row, false);
                                    dense = Some(DenseMode::CandidateMask(mask));
                                } else {
                                    sparse_passes = std::mem::take(&mut prefix_passes);
                                }
                            }
                        }
                    }
                }
            } else {
                for row in 0..row_count {
                    let mut inner_ctx = row_ctx.clone();
                    inner_ctx.push(&table, row);
                    let pred = engine.eval_scalar(model, expr, &base_filter, &inner_ctx, env)?;
                    let keep = pred.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                    match &mut dense {
                        Some(DenseMode::CandidateMask(mask)) => {
                            if !keep {
                                mask.set(row, false);
                            }
                        }
                        Some(DenseMode::PassMask(mask)) => {
                            if keep {
                                mask.set(row, true);
                            }
                        }
                        None => {
                            if keep {
                                if any_fail {
                                    sparse_passes.push(row);
                                    if sparse_passes.len() > sparse_to_dense_threshold {
                                        let mut mask = BitVec::with_len_all_false(row_count);
                                        for &row in &sparse_passes {
                                            mask.set(row, true);
                                        }
                                        sparse_passes.clear();
                                        dense = Some(DenseMode::PassMask(mask));
                                    }
                                } else if !prefix_dense {
                                    prefix_passes.push(row);
                                    if prefix_passes.len() > sparse_to_dense_threshold {
                                        prefix_dense = true;
                                        prefix_passes.clear();
                                    }
                                }
                            } else {
                                any_fail = true;
                                if prefix_dense {
                                    let mut mask = BitVec::with_len_all_true(row_count);
                                    mask.set(row, false);
                                    dense = Some(DenseMode::CandidateMask(mask));
                                } else {
                                    sparse_passes = std::mem::take(&mut prefix_passes);
                                }
                            }
                        }
                    }
                }
            }

            if !any_fail {
                row_filters.push((table, RowFilter::All));
            } else {
                match dense {
                    Some(DenseMode::CandidateMask(mask)) | Some(DenseMode::PassMask(mask)) => {
                        let visible = mask.count_ones();
                        if visible <= sparse_to_dense_threshold {
                            row_filters.push((table, RowFilter::Rows(mask.iter_ones().collect())));
                        } else {
                            row_filters.push((table, RowFilter::Mask(Arc::new(mask))));
                        }
                    }
                    None => {
                        row_filters
                            .push((table, RowFilter::Rows(sparse_passes.into_iter().collect())));
                    }
                }
            }
            Ok(())
        }

        fn apply_filter_arg(
            engine: &DaxEngine,
            model: &DataModel,
            arg: &Expr,
            eval_filter: &FilterContext,
            row_ctx: &RowContext,
            env: &mut VarEnv,
            keep_filters: bool,
            clear_tables: &mut HashSet<String>,
            clear_columns: &mut HashSet<(String, String)>,
            row_filters: &mut Vec<(String, RowFilter)>,
            column_filters: &mut Vec<((String, String), HashSet<Value>)>,
        ) -> DaxResult<()> {
            // `KEEPFILTERS` wraps a normal filter argument, but changes its semantics from
            // replacement (clear existing filters on the target table/column) to intersection.
            // We implement that by evaluating the inner argument as usual, but skipping any
            // additions to `clear_tables` / `clear_columns`.
            //
            // Note: `KEEPFILTERS` can appear redundantly nested; treat multiple wrappers as a
            // single `keep_filters = true` flag.
            let mut arg = arg;
            let mut keep_filters = keep_filters;
            loop {
                match arg {
                    Expr::Call { name, args } if name.eq_ignore_ascii_case("KEEPFILTERS") => {
                        let [inner] = args.as_slice() else {
                            return Err(DaxError::Eval(
                                "KEEPFILTERS expects exactly 1 argument".into(),
                            ));
                        };
                        arg = inner;
                        keep_filters = true;
                    }
                    _ => break,
                }
            }

            match arg {
                Expr::Let { bindings, body } => {
                    env.push_scope();
                    let result = (|| -> DaxResult<()> {
                        for (name, binding_expr) in bindings {
                            let value = engine.eval_var_value(
                                model,
                                binding_expr,
                                eval_filter,
                                row_ctx,
                                env,
                            )?;
                            env.define(name, value);
                        }
                        apply_filter_arg(
                            engine,
                            model,
                            body,
                            eval_filter,
                            row_ctx,
                            env,
                            keep_filters,
                            clear_tables,
                            clear_columns,
                            row_filters,
                            column_filters,
                        )
                    })();
                    env.pop_scope();
                    result
                }
                Expr::Call { name, .. } if name.eq_ignore_ascii_case("USERELATIONSHIP") => Ok(()),
                Expr::Call { name, .. } if name.eq_ignore_ascii_case("CROSSFILTER") => Ok(()),
                Expr::Call { name, args }
                    if name.eq_ignore_ascii_case("ALL")
                        || name.eq_ignore_ascii_case("REMOVEFILTERS") =>
                {
                    // `REMOVEFILTERS` is an alias for the `ALL` filter modifier semantics:
                    // clear filters for the referenced table/column.
                    let function_name = if name.eq_ignore_ascii_case("ALL") {
                        "ALL"
                    } else {
                        "REMOVEFILTERS"
                    };

                    let [inner] = args.as_slice() else {
                        return Err(DaxError::Eval(format!(
                            "{function_name} expects 1 argument"
                        )));
                    };
                    match inner {
                        Expr::TableName(table) => {
                            if !keep_filters {
                                clear_tables.insert(table.clone());
                            }
                            Ok(())
                        }
                        Expr::ColumnRef { table, column } => {
                            if !keep_filters {
                                clear_columns.insert((table.clone(), column.clone()));
                            }
                            Ok(())
                        }
                        other => Err(DaxError::Type(format!(
                            "{function_name} expects a table name or column reference, got {other:?}"
                        ))),
                    }
                }
                Expr::Call { name, args } if name.eq_ignore_ascii_case("ALLNOBLANKROW") => {
                    let [inner] = args.as_slice() else {
                        return Err(DaxError::Eval("ALLNOBLANKROW expects 1 argument".into()));
                    };
                    match inner {
                        Expr::TableName(table) => {
                            let table_key = normalize_ident(table);
                            if !keep_filters {
                                clear_tables.insert(table_key.clone());
                            }
                            // Apply an explicit row filter containing all physical rows. This
                            // matches `ALL(Table)` while ensuring the relationship-generated blank
                            // member is excluded (`blank_row_allowed` is false when a row filter is
                            // present).
                            row_filters.push((table_key, RowFilter::All));
                            Ok(())
                        }
                        Expr::ColumnRef { table, column } => {
                            let key = (normalize_ident(table), normalize_ident(column));
                            if !keep_filters {
                                clear_columns.insert(key.clone());
                            }

                            let mut base_filter = eval_filter.clone();
                            base_filter.clear_column_filter(table, column);
                            let mut values =
                                engine.distinct_column_values(model, inner, &base_filter)?;
                            values.retain(|v| !v.is_blank());
                            column_filters.push((key, values));
                            Ok(())
                        }
                        other => Err(DaxError::Type(format!(
                            "ALLNOBLANKROW expects a table name or column reference, got {other:?}"
                        ))),
                    }
                }
                // Boolean filter expressions like:
                //   Orders[Amount] > 10 && Orders[Amount] < 20
                //   NOT(Orders[Amount] > 10)
                // These are treated like table filters against the one referenced table.
                Expr::BinaryOp {
                    op: BinaryOp::And | BinaryOp::Or,
                    ..
                } => apply_boolean_filter_expr(
                    engine,
                    model,
                    arg,
                    eval_filter,
                    row_ctx,
                    env,
                    keep_filters,
                    clear_columns,
                    row_filters,
                ),
                Expr::Call { name, .. }
                    if name.eq_ignore_ascii_case("NOT")
                        || name.eq_ignore_ascii_case("AND")
                        || name.eq_ignore_ascii_case("OR") =>
                {
                    apply_boolean_filter_expr(
                        engine,
                        model,
                        arg,
                        eval_filter,
                        row_ctx,
                        env,
                        keep_filters,
                        clear_columns,
                        row_filters,
                    )
                }
                Expr::BinaryOp {
                    op: BinaryOp::In,
                    left,
                    right,
                } if matches!(left.as_ref(), Expr::Tuple(_)) => {
                    // Multi-column `IN` filters in CALCULATE use a row constructor on the LHS:
                    //   (T[Col1], T[Col2]) IN {(1,2), (3,4)}
                    //
                    // These filters cannot be expressed as independent per-column value filters
                    // because they preserve correlation across columns, so we evaluate them as a
                    // row filter against the referenced table.
                    let Expr::Tuple(tuple_exprs) = left.as_ref() else {
                        unreachable!("guarded by matches! above")
                    };

                    if tuple_exprs.is_empty() {
                        return Err(DaxError::Eval(
                            "IN row constructor cannot be empty".to_string(),
                        ));
                    }

                    let mut target_table_key: Option<String> = None;
                    let mut target_table_display: Option<String> = None;
                    let mut target_columns: Vec<(String, String)> =
                        Vec::with_capacity(tuple_exprs.len());
                    for expr in tuple_exprs {
                        let Expr::ColumnRef { table, column } = expr else {
                            return Err(DaxError::Eval(
                                 "CALCULATE row constructor IN filter must contain only column references"
                                     .into(),
                             ));
                        };
                        let table_key = normalize_ident(table);
                        match target_table_key.as_deref() {
                            Some(existing) if existing != table_key => {
                                return Err(DaxError::Eval(
                                     "CALCULATE row constructor IN filter must reference exactly one table"
                                         .into(),
                                 ));
                            }
                            Some(_) => {}
                            None => {
                                target_table_key = Some(table_key.to_string());
                                target_table_display = Some(table.clone());
                            }
                        }
                        target_columns.push((column.clone(), normalize_ident(column)));
                    }

                    let target_table_key = target_table_key.expect("set above");
                    let target_table_display =
                        target_table_display.unwrap_or_else(|| target_table_key.clone());
                    let target_table_ref = model
                        .table(&target_table_key)
                        .ok_or_else(|| DaxError::UnknownTable(target_table_display.clone()))?;

                    let mut target_col_indices: Vec<usize> =
                        Vec::with_capacity(target_columns.len());
                    let mut referenced_column_keys: Vec<(String, String)> =
                        Vec::with_capacity(target_columns.len());
                    for (column_display, column_key) in &target_columns {
                        let idx = target_table_ref.column_idx(column_display).ok_or_else(|| {
                            DaxError::UnknownColumn {
                                table: target_table_display.clone(),
                                column: column_display.clone(),
                            }
                        })?;
                        target_col_indices.push(idx);
                        referenced_column_keys.push((target_table_key.clone(), column_key.clone()));
                    }

                    // Evaluate the RHS in a context where existing filters on the referenced
                    // columns are removed (replacement semantics).
                    let mut base_filter = eval_filter.clone();
                    for key in &referenced_column_keys {
                        base_filter.column_filters.remove(key);
                        if !keep_filters {
                            clear_columns.insert(key.clone());
                        }
                    }

                    let expected_cols = target_col_indices.len();

                    // Materialize RHS row tuples so we can test membership efficiently across
                    // target rows.
                    let rhs_table = engine.eval_table(model, right, &base_filter, row_ctx, env)?;
                    let rhs_rows: Vec<Vec<Value>> = match rhs_table {
                        TableResult::Physical {
                            table,
                            rows,
                            visible_cols,
                        } => {
                            let table_ref = model
                                .table(&table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                            let visible_cols = visible_cols.as_deref();
                            let col_count = visible_cols
                                .map(|cols| cols.len())
                                .unwrap_or_else(|| table_ref.columns().len());
                            if col_count != expected_cols {
                                return Err(DaxError::Eval(format!(
                                     "IN row constructor expected {expected_cols} columns, got {col_count}"
                                 )));
                            }
                            let mut out = Vec::with_capacity(rows.len());
                            if let Some(cols) = visible_cols {
                                for row in rows {
                                    out.push(
                                        cols.iter()
                                            .map(|&col_idx| {
                                                table_ref
                                                    .value_by_idx(row, col_idx)
                                                    .unwrap_or(Value::Blank)
                                            })
                                            .collect(),
                                    );
                                }
                            } else {
                                for row in rows {
                                    out.push(
                                        (0..expected_cols)
                                            .map(|col_idx| {
                                                table_ref
                                                    .value_by_idx(row, col_idx)
                                                    .unwrap_or(Value::Blank)
                                            })
                                            .collect(),
                                    );
                                }
                            }
                            out
                        }
                        TableResult::PhysicalAll {
                            table,
                            row_count,
                            visible_cols,
                        } => {
                            let table_ref = model
                                .table(&table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                            let visible_cols = visible_cols.as_deref();
                            let col_count = visible_cols
                                .map(|cols| cols.len())
                                .unwrap_or_else(|| table_ref.columns().len());
                            if col_count != expected_cols {
                                return Err(DaxError::Eval(format!(
                                     "IN row constructor expected {expected_cols} columns, got {col_count}"
                                 )));
                            }
                            let mut out = Vec::with_capacity(row_count);
                            if let Some(cols) = visible_cols {
                                for row in 0..row_count {
                                    out.push(
                                        cols.iter()
                                            .map(|&col_idx| {
                                                table_ref
                                                    .value_by_idx(row, col_idx)
                                                    .unwrap_or(Value::Blank)
                                            })
                                            .collect(),
                                    );
                                }
                            } else {
                                for row in 0..row_count {
                                    out.push(
                                        (0..expected_cols)
                                            .map(|col_idx| {
                                                table_ref
                                                    .value_by_idx(row, col_idx)
                                                    .unwrap_or(Value::Blank)
                                            })
                                            .collect(),
                                    );
                                }
                            }
                            out
                        }
                        TableResult::PhysicalMask {
                            table,
                            mask,
                            visible_cols,
                        } => {
                            let table_ref = model
                                .table(&table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                            let visible_cols = visible_cols.as_deref();
                            let col_count = visible_cols
                                .map(|cols| cols.len())
                                .unwrap_or_else(|| table_ref.columns().len());
                            if col_count != expected_cols {
                                return Err(DaxError::Eval(format!(
                                     "IN row constructor expected {expected_cols} columns, got {col_count}"
                                 )));
                            }
                            let mut out = Vec::with_capacity(mask.count_ones());
                            if let Some(cols) = visible_cols {
                                for row in mask.iter_ones() {
                                    out.push(
                                        cols.iter()
                                            .map(|&col_idx| {
                                                table_ref
                                                    .value_by_idx(row, col_idx)
                                                    .unwrap_or(Value::Blank)
                                            })
                                            .collect(),
                                    );
                                }
                            } else {
                                for row in mask.iter_ones() {
                                    out.push(
                                        (0..expected_cols)
                                            .map(|col_idx| {
                                                table_ref
                                                    .value_by_idx(row, col_idx)
                                                    .unwrap_or(Value::Blank)
                                            })
                                            .collect(),
                                    );
                                }
                            }
                            out
                        }
                        TableResult::Virtual { columns, rows } => {
                            if columns.len() != expected_cols {
                                return Err(DaxError::Eval(format!(
                                    "IN row constructor expected {expected_cols} columns, got {}",
                                    columns.len()
                                )));
                            }
                            rows
                        }
                    };

                    // Test membership across candidate rows in the target table, respecting any
                    // other active filters (but ignoring existing filters on the referenced
                    // columns, unless KEEPFILTERS is used).
                    let row_sets = (!base_filter.is_empty())
                        .then(|| resolve_row_sets(model, &base_filter))
                        .transpose()?;

                    let mut allowed_rows = HashSet::new();
                    let candidate_count: usize;
                    if let Some(sets) = row_sets.as_ref() {
                        let allowed = sets
                            .get(&target_table_key)
                            .ok_or_else(|| DaxError::UnknownTable(target_table_display.clone()))?;
                        candidate_count = allowed.count_ones();
                        for row in allowed.iter_ones() {
                            let mut row_values = Vec::with_capacity(expected_cols);
                            for &col_idx in &target_col_indices {
                                row_values.push(
                                    target_table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank),
                                );
                            }

                            let mut matched = false;
                            for rhs in &rhs_rows {
                                let mut all_equal = true;
                                for (lhs_value, rhs_value) in row_values.iter().zip(rhs.iter()) {
                                    if !compare_values(&BinaryOp::Equals, lhs_value, rhs_value)? {
                                        all_equal = false;
                                        break;
                                    }
                                }
                                if all_equal {
                                    matched = true;
                                    break;
                                }
                            }

                            if matched {
                                allowed_rows.insert(row);
                            }
                        }
                    } else {
                        candidate_count = target_table_ref.row_count();
                        for row in 0..target_table_ref.row_count() {
                            let mut row_values = Vec::with_capacity(expected_cols);
                            for &col_idx in &target_col_indices {
                                row_values.push(
                                    target_table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank),
                                );
                            }

                            let mut matched = false;
                            for rhs in &rhs_rows {
                                let mut all_equal = true;
                                for (lhs_value, rhs_value) in row_values.iter().zip(rhs.iter()) {
                                    if !compare_values(&BinaryOp::Equals, lhs_value, rhs_value)? {
                                        all_equal = false;
                                        break;
                                    }
                                }
                                if all_equal {
                                    matched = true;
                                    break;
                                }
                            }

                            if matched {
                                allowed_rows.insert(row);
                            }
                        }
                    }

                    if allowed_rows.len() == candidate_count {
                        row_filters.push((target_table_key, RowFilter::All));
                    } else {
                        row_filters.push((target_table_key, RowFilter::Rows(allowed_rows)));
                    }

                    Ok(())
                }
                Expr::BinaryOp { op, left, right } => {
                    let Expr::ColumnRef { table, column } = left.as_ref() else {
                        return Err(DaxError::Eval(
                            "CALCULATE filter must be a column comparison".into(),
                        ));
                    };

                    let key = (normalize_ident(table), normalize_ident(column));
                    if !keep_filters {
                        clear_columns.insert(key.clone());
                    }

                    match op {
                        BinaryOp::In => {
                            let table_result =
                                engine.eval_table(model, right, eval_filter, row_ctx, env)?;
                            let values: HashSet<Value> = match table_result {
                                TableResult::Physical {
                                    table,
                                    rows,
                                    visible_cols,
                                } => {
                                    let table_ref = model
                                        .table(&table)
                                        .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                                    let (col_count, col_idx) = match visible_cols.as_deref() {
                                        Some(cols) => {
                                            (cols.len(), cols.get(0).copied().unwrap_or(0))
                                        }
                                        None => (table_ref.columns().len(), 0),
                                    };
                                    if col_count != 1 {
                                        return Err(DaxError::Eval(
                                            "IN currently only supports one-column tables".into(),
                                        ));
                                    }
                                    let mut out = HashSet::new();
                                    for row in rows {
                                        let value = table_ref
                                            .value_by_idx(row, col_idx)
                                            .unwrap_or(Value::Blank);
                                        out.insert(value);
                                    }
                                    out
                                }
                                TableResult::PhysicalMask {
                                    table,
                                    mask,
                                    visible_cols,
                                } => {
                                    let table_ref = model
                                        .table(&table)
                                        .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                                    let (col_count, col_idx) = match visible_cols.as_deref() {
                                        Some(cols) => {
                                            (cols.len(), cols.get(0).copied().unwrap_or(0))
                                        }
                                        None => (table_ref.columns().len(), 0),
                                    };
                                    if col_count != 1 {
                                        return Err(DaxError::Eval(
                                            "IN currently only supports one-column tables".into(),
                                        ));
                                    }
                                    let mut out = HashSet::new();
                                    for row in mask.iter_ones() {
                                        let value = table_ref
                                            .value_by_idx(row, col_idx)
                                            .unwrap_or(Value::Blank);
                                        out.insert(value);
                                    }
                                    out
                                }
                                TableResult::PhysicalAll {
                                    table,
                                    row_count,
                                    visible_cols,
                                } => {
                                    let table_ref = model
                                        .table(&table)
                                        .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                                    let (col_count, col_idx) = match visible_cols.as_deref() {
                                        Some(cols) => {
                                            (cols.len(), cols.get(0).copied().unwrap_or(0))
                                        }
                                        None => (table_ref.columns().len(), 0),
                                    };
                                    if col_count != 1 {
                                        return Err(DaxError::Eval(
                                            "IN currently only supports one-column tables".into(),
                                        ));
                                    }
                                    let mut out = HashSet::new();
                                    for row in 0..row_count {
                                        let value = table_ref
                                            .value_by_idx(row, col_idx)
                                            .unwrap_or(Value::Blank);
                                        out.insert(value);
                                    }
                                    out
                                }
                                TableResult::Virtual { columns, rows } => {
                                    if columns.len() != 1 {
                                        return Err(DaxError::Eval(
                                            "IN currently only supports one-column tables".into(),
                                        ));
                                    }
                                    rows.into_iter()
                                        .map(|row_values| {
                                            row_values.get(0).cloned().unwrap_or(Value::Blank)
                                        })
                                        .collect()
                                }
                             };

                            column_filters.push((key, values));
                            Ok(())
                        }
                        BinaryOp::Equals => {
                            let rhs =
                                engine.eval_scalar(model, right, eval_filter, row_ctx, env)?;
                            column_filters.push((key, HashSet::from([rhs])));
                            Ok(())
                        }
                        BinaryOp::NotEquals
                        | BinaryOp::Less
                        | BinaryOp::LessEquals
                        | BinaryOp::Greater
                        | BinaryOp::GreaterEquals => {
                            let rhs =
                                engine.eval_scalar(model, right, eval_filter, row_ctx, env)?;
                            let mut base_filter = eval_filter.clone();
                            base_filter.column_filters.remove(&key);

                            let table_ref = model
                                .table(table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                            let idx = table_ref.column_idx(column).ok_or_else(|| {
                                DaxError::UnknownColumn {
                                    table: table.clone(),
                                    column: column.clone(),
                                }
                            })?;

                            let mut allowed = HashSet::new();
                            let row_sets = (!base_filter.is_empty())
                                .then(|| resolve_row_sets(model, &base_filter))
                                .transpose()?;
                            let table_key = normalize_ident(table);
                            if let Some(sets) = row_sets.as_ref() {
                                let allowed_rows = sets
                                    .get(&table_key)
                                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                                for row in allowed_rows.iter_ones() {
                                    let lhs =
                                        table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                    let keep =
                                        match engine.eval_binary(op, lhs.clone(), rhs.clone())? {
                                            Value::Boolean(b) => b,
                                            other => {
                                                return Err(DaxError::Type(format!(
                                                "expected comparison to return boolean, got {other}"
                                            )))
                                            }
                                        };

                                    if keep {
                                        allowed.insert(lhs);
                                    }
                                }
                            } else {
                                for row in 0..table_ref.row_count() {
                                    let lhs =
                                        table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                    let keep =
                                        match engine.eval_binary(op, lhs.clone(), rhs.clone())? {
                                            Value::Boolean(b) => b,
                                            other => {
                                                return Err(DaxError::Type(format!(
                                                "expected comparison to return boolean, got {other}"
                                            )))
                                            }
                                        };

                                    if keep {
                                        allowed.insert(lhs);
                                    }
                                }
                            }

                            column_filters.push((key, allowed));
                            Ok(())
                        }
                        _ => Err(DaxError::Eval(format!(
                            "unsupported CALCULATE filter operator {op:?}"
                        ))),
                    }
                }
                Expr::Call { name, args } if name.eq_ignore_ascii_case("TREATAS") => {
                    let [source, target] = args.as_slice() else {
                        return Err(DaxError::Eval("TREATAS expects 2 arguments".into()));
                    };

                    let source_col_expr = match source {
                        Expr::Call {
                            name: source_fn,
                            args: source_args,
                        } if (source_fn.eq_ignore_ascii_case("VALUES")
                            || source_fn.eq_ignore_ascii_case("DISTINCT"))
                            && matches!(source_args.as_slice(), [Expr::ColumnRef { .. }]) =>
                        {
                            &source_args[0]
                        }
                        _ => {
                            return Err(DaxError::Type(
                                "TREATAS currently only supports VALUES(column) or DISTINCT(column) as its first argument"
                                    .into(),
                            ))
                        }
                    };

                    let Expr::ColumnRef {
                        table: target_table,
                        column: target_column,
                    } = target
                    else {
                        return Err(DaxError::Type(
                            "TREATAS expects a target column reference as its second argument"
                                .into(),
                        ));
                    };

                    let key = (
                        normalize_ident(target_table),
                        normalize_ident(target_column),
                    );
                    if !keep_filters {
                        clear_columns.insert(key.clone());
                    }
                    let values =
                        engine.distinct_column_values(model, source_col_expr, eval_filter)?;
                    column_filters.push((key, values));
                    Ok(())
                }
                Expr::Call { name, args }
                    if (name.eq_ignore_ascii_case("VALUES")
                        || name.eq_ignore_ascii_case("DISTINCT"))
                        && matches!(args.as_slice(), [Expr::ColumnRef { .. }]) =>
                {
                    let Expr::ColumnRef { table, column } = &args[0] else {
                        unreachable!("checked above");
                    };
                    let key = (normalize_ident(table), normalize_ident(column));
                    if !keep_filters {
                        clear_columns.insert(key.clone());
                    }
                    let values = engine.distinct_column_values(model, &args[0], eval_filter)?;
                    column_filters.push((key, values));
                    Ok(())
                }
                Expr::Call { .. } | Expr::TableName(_) => {
                    let table_filter = engine.eval_table(model, arg, eval_filter, row_ctx, env)?;
                    match table_filter {
                        TableResult::Physical {
                            table,
                            rows,
                            visible_cols,
                        } => {
                            // Table results are expected to use normalized identifiers. Normalize
                            // again here to ensure CALCULATE filter intersection semantics remain
                            // case-insensitive even if a table expression returns a non-normalized
                            // name.
                            let table_key = normalize_ident(&table);
                            // `VALUES(Table[Column])`-shaped physical tables are represented as
                            // `(table, representative_row_indices, visible_cols=[col])`.
                            //
                            // When used as a CALCULATE table filter argument, those tables should
                            // behave like a *column filter* (apply the set of visible values for
                            // the visible column), not like a row filter on representative physical
                            // rows.
                            //
                            // This ensures patterns like:
                            //   CALCULATE([Measure], FILTER(VALUES(T[Col]), ...))
                            // filter by the chosen column values rather than arbitrarily selecting
                            // one physical row per distinct value.
                            if let Some(visible_cols) = visible_cols {
                                if visible_cols.len() != 1 {
                                    return Err(DaxError::Eval(
                                        "CALCULATE table filters for projected physical tables currently only support one column"
                                            .into(),
                                    ));
                                }
                                let col_idx = visible_cols[0];
                                let table_ref = model
                                    .table(&table)
                                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                                let Some(column) = table_ref.columns().get(col_idx) else {
                                    return Err(DaxError::Eval(format!(
                                        "row context refers to out-of-bounds column index {col_idx} for table {table}"
                                    )));
                                };

                                let mut values = HashSet::new();
                                for row in rows {
                                    let v = table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank);
                                    values.insert(v);
                                }

                                let key = (table_key.clone(), normalize_ident(column));
                                if !keep_filters {
                                    clear_columns.insert(key.clone());
                                }
                                column_filters.push((key, values));
                                Ok(())
                            } else {
                                if !keep_filters {
                                    clear_tables.insert(table_key.clone());
                                }
                                row_filters
                                    .push((table_key, RowFilter::Rows(rows.into_iter().collect())));
                                Ok(())
                            }
                        }
                        TableResult::PhysicalAll {
                            table,
                            row_count,
                            visible_cols,
                        } => {
                            // See `TableResult::Physical` above for semantics.
                            let table_key = normalize_ident(&table);
                            if let Some(visible_cols) = visible_cols {
                                if visible_cols.len() != 1 {
                                    return Err(DaxError::Eval(
                                        "CALCULATE table filters for projected physical tables currently only support one column"
                                            .into(),
                                    ));
                                }
                                let col_idx = visible_cols[0];
                                let table_ref = model
                                    .table(&table)
                                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                                let Some(column) = table_ref.columns().get(col_idx) else {
                                    return Err(DaxError::Eval(format!(
                                        "row context refers to out-of-bounds column index {col_idx} for table {table}"
                                    )));
                                };

                                let mut values = HashSet::new();
                                for row in 0..row_count {
                                    let v = table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank);
                                    values.insert(v);
                                }

                                let key = (table_key.clone(), normalize_ident(column));
                                if !keep_filters {
                                    clear_columns.insert(key.clone());
                                }
                                column_filters.push((key, values));
                                Ok(())
                            } else {
                                if !keep_filters {
                                    clear_tables.insert(table_key.clone());
                                }
                                row_filters.push((table_key, RowFilter::All));
                                Ok(())
                            }
                        }
                        TableResult::PhysicalMask {
                            table,
                            mask,
                            visible_cols,
                        } => {
                            // See `TableResult::Physical` above for semantics.
                            let table_key = normalize_ident(&table);
                            if let Some(visible_cols) = visible_cols {
                                if visible_cols.len() != 1 {
                                    return Err(DaxError::Eval(
                                        "CALCULATE table filters for projected physical tables currently only support one column"
                                            .into(),
                                    ));
                                }
                                let col_idx = visible_cols[0];
                                let table_ref = model
                                    .table(&table)
                                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                                let Some(column) = table_ref.columns().get(col_idx) else {
                                    return Err(DaxError::Eval(format!(
                                        "row context refers to out-of-bounds column index {col_idx} for table {table}"
                                    )));
                                };

                                let mut values = HashSet::new();
                                for row in mask.iter_ones() {
                                    let v = table_ref
                                        .value_by_idx(row, col_idx)
                                        .unwrap_or(Value::Blank);
                                    values.insert(v);
                                }

                                let key = (table_key.clone(), normalize_ident(column));
                                if !keep_filters {
                                    clear_columns.insert(key.clone());
                                }
                                column_filters.push((key, values));
                                Ok(())
                            } else {
                                if !keep_filters {
                                    clear_tables.insert(table_key.clone());
                                }
                                if mask.all_true() {
                                    row_filters.push((table_key, RowFilter::All));
                                } else {
                                    row_filters.push((table_key, RowFilter::Mask(mask)));
                                }
                                Ok(())
                            }
                        }
                        TableResult::Virtual { .. } => Err(DaxError::Eval(
                            "CALCULATE table filter must be a physical table expression".into(),
                        )),
                    }
                }
                other => Err(DaxError::Eval(format!(
                    "unsupported CALCULATE filter argument {other:?}"
                ))),
            }
        }

        for arg in filter_args {
            apply_filter_arg(
                self,
                model,
                arg,
                &eval_filter,
                row_ctx,
                env,
                false,
                &mut clear_tables,
                &mut clear_columns,
                &mut row_filters,
                &mut column_filters,
            )?;
        }

        for table in clear_tables {
            filter.clear_table_filters(&table);
        }
        for (table, column) in clear_columns {
            filter.clear_column_filter(&table, &column);
        }

        for (table, rows) in row_filters {
            match filter.row_filters.get_mut(&table) {
                Some(existing) => match rows {
                    RowFilter::All => {
                        // Intersecting with "all rows" does not change the existing restriction.
                    }
                    RowFilter::Rows(rows) => match existing {
                        RowFilter::All => {
                            *existing = RowFilter::Rows(rows);
                        }
                        RowFilter::Rows(existing_rows) => {
                            // Intersect row sets. Prefer retaining the smaller set to minimize
                            // hash lookups.
                            if rows.len() < existing_rows.len() {
                                let mut next = rows;
                                next.retain(|row| existing_rows.contains(row));
                                *existing_rows = next;
                            } else {
                                existing_rows.retain(|row| rows.contains(row));
                            }
                        }
                        RowFilter::Mask(existing_mask) => {
                            let mask = existing_mask.clone();
                            let mut next = rows;
                            next.retain(|row| *row < mask.len() && mask.get(*row));
                            *existing = RowFilter::Rows(next);
                        }
                    },
                    RowFilter::Mask(mask) => match existing {
                        RowFilter::All => {
                            *existing = RowFilter::Mask(mask);
                        }
                        RowFilter::Rows(existing_rows) => {
                            existing_rows.retain(|row| *row < mask.len() && mask.get(*row));
                        }
                        RowFilter::Mask(existing_mask) => {
                            // Prefer keeping the dense representation unless the result becomes
                            // sparse enough that a `HashSet` is more efficient.
                            if existing_mask.len() != mask.len() {
                                *existing = RowFilter::Rows(
                                    existing_mask
                                        .iter_ones()
                                        .filter(|row| *row < mask.len() && mask.get(*row))
                                        .collect(),
                                );
                            } else {
                                let mut next = existing_mask.as_ref().clone();
                                next.and_inplace(&mask);
                                let visible = next.count_ones();
                                let row_count = next.len();
                                let sparse_to_dense_threshold = row_count / 64;
                                if visible == row_count {
                                    *existing = RowFilter::All;
                                } else if visible <= sparse_to_dense_threshold {
                                    *existing = RowFilter::Rows(next.iter_ones().collect());
                                } else {
                                    *existing = RowFilter::Mask(Arc::new(next));
                                }
                            }
                        }
                    },
                },
                None => filter.set_row_filter(&table, rows),
            }
        }

        for (key, values) in column_filters {
            match filter.column_filters.get_mut(&key) {
                Some(existing) => existing.retain(|v| values.contains(v)),
                None => {
                    filter.column_filters.insert(key, values);
                }
            }
        }

        Ok(())
    }

    fn apply_userelationship(
        &self,
        model: &DataModel,
        filter: &mut FilterContext,
        args: &[Expr],
    ) -> DaxResult<()> {
        let [left, right] = args else {
            return Err(DaxError::Eval("USERELATIONSHIP expects 2 arguments".into()));
        };
        let Expr::ColumnRef {
            table: left_table,
            column: left_column,
        } = left
        else {
            return Err(DaxError::Type(
                "USERELATIONSHIP expects column references".into(),
            ));
        };
        let Expr::ColumnRef {
            table: right_table,
            column: right_column,
        } = right
        else {
            return Err(DaxError::Type(
                "USERELATIONSHIP expects column references".into(),
            ));
        };

        let Some(rel_idx) =
            model.find_relationship_index(left_table, left_column, right_table, right_column)?
        else {
            return Err(DaxError::Eval(format!(
                "no relationship found between {left_table}[{left_column}] and {right_table}[{right_column}]"
            )));
        };

        filter.activate_relationship(rel_idx);
        Ok(())
    }

    fn apply_crossfilter(
        &self,
        model: &DataModel,
        filter: &mut FilterContext,
        args: &[Expr],
    ) -> DaxResult<(usize, RelationshipOverride)> {
        let [left, right, direction] = args else {
            return Err(DaxError::Eval("CROSSFILTER expects 3 arguments".into()));
        };
        let Expr::ColumnRef {
            table: left_table,
            column: left_column,
        } = left
        else {
            return Err(DaxError::Type(
                "CROSSFILTER expects column references".into(),
            ));
        };
        let Expr::ColumnRef {
            table: right_table,
            column: right_column,
        } = right
        else {
            return Err(DaxError::Type(
                "CROSSFILTER expects column references".into(),
            ));
        };

        // The third argument is a bare identifier in DAX, and our parser represents bare
        // identifiers as `Expr::TableName`. Some users also write it as a string literal.
        let direction = match direction {
            Expr::TableName(name) => name.as_str(),
            Expr::Text(s) => s.as_str(),
            other => {
                return Err(DaxError::Type(format!(
                    "CROSSFILTER expects a direction identifier or string, got {other:?}"
                )))
            }
        };
        let direction = direction.trim().to_ascii_uppercase();

        let Some(rel_idx) =
            model.find_relationship_index(left_table, left_column, right_table, right_column)?
        else {
            return Err(DaxError::Eval(format!(
                "no relationship found between {left_table}[{left_column}] and {right_table}[{right_column}]"
            )));
        };

        let rel = model
            .relationships()
            .get(rel_idx)
            .expect("relationship index from find_relationship_index");

        let resolve_one_way = |source_table: &str,
                               source_column: &str,
                               target_table: &str,
                               target_column: &str|
         -> DaxResult<RelationshipOverride> {
            let source_table_key = normalize_ident(source_table);
            let source_column_key = normalize_ident(source_column);
            let target_table_key = normalize_ident(target_table);
            let target_column_key = normalize_ident(target_column);
            if rel.to_table_key == source_table_key
                && rel.to_column_key == source_column_key
                && rel.from_table_key == target_table_key
                && rel.from_column_key == target_column_key
            {
                // `to_table` filters `from_table` (relationship's default orientation).
                return Ok(RelationshipOverride::Active(CrossFilterDirection::Single));
            }
            if rel.from_table_key == source_table_key
                && rel.from_column_key == source_column_key
                && rel.to_table_key == target_table_key
                && rel.to_column_key == target_column_key
            {
                // Reverse of the relationship's default orientation.
                return Ok(RelationshipOverride::OneWayReverse);
            }
            Err(DaxError::Eval(format!(
                "CROSSFILTER direction {direction} does not match relationship {}",
                rel.rel.name
            )))
        };

        let override_dir = match direction.as_str() {
            "BOTH" => RelationshipOverride::Active(CrossFilterDirection::Both),
            // DAX uses `ONEWAY` but we'll accept the more explicit `SINGLE` as well.
            "ONEWAY" | "SINGLE" => RelationshipOverride::Active(CrossFilterDirection::Single),
            "NONE" => RelationshipOverride::Disabled,
            "ONEWAY_LEFTFILTERSRIGHT" => resolve_one_way(
                left_table.as_str(),
                left_column.as_str(),
                right_table.as_str(),
                right_column.as_str(),
            )?,
            "ONEWAY_RIGHTFILTERSLEFT" => resolve_one_way(
                right_table.as_str(),
                right_column.as_str(),
                left_table.as_str(),
                left_column.as_str(),
            )?,
            other => {
                return Err(DaxError::Eval(format!(
                    "unsupported CROSSFILTER direction {other}"
                )))
            }
        };

        filter.cross_filter_overrides.insert(rel_idx, override_dir);
        Ok((rel_idx, override_dir))
    }

    fn eval_related(
        &self,
        model: &DataModel,
        arg: &Expr,
        filter: &FilterContext,
        row_ctx: &RowContext,
    ) -> DaxResult<Value> {
        let Expr::ColumnRef { table, column } = arg else {
            return Err(DaxError::Type("RELATED expects a column reference".into()));
        };
        let Some(current_table) = row_ctx.current_table() else {
            return Err(DaxError::Eval("RELATED requires row context".into()));
        };

        let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
        for &idx in filter.relationship_overrides() {
            if let Some(rel) = model.relationships().get(idx) {
                override_pairs.insert((rel.rel.from_table.as_str(), rel.rel.to_table.as_str()));
            }
        }
        let (current_row, current_visible_cols) = row_ctx
            .physical_row_for(current_table)
            .ok_or_else(|| DaxError::Eval("missing row for current table".into()))?;

        let Some(path) = model.find_unique_active_relationship_path(
            current_table,
            table,
            RelationshipPathDirection::ManyToOne,
            |idx, rel| {
                let pair = (rel.rel.from_table.as_str(), rel.rel.to_table.as_str());
                let is_active = if override_pairs.contains(&pair) {
                    filter.relationship_overrides().contains(&idx)
                } else {
                    rel.rel.is_active
                };
                is_active && !filter.is_relationship_disabled(idx)
            },
        )?
        else {
            return Err(DaxError::Eval(format!(
                "no active relationship from {current_table} to {table} for RELATED"
            )));
        };

        let mut row = current_row;
        for (hop_idx, rel_idx) in path.into_iter().enumerate() {
            let rel_info = model
                .relationships()
                .get(rel_idx)
                .expect("relationship index from path");

            // If the current row context is restricted (e.g. iterating `VALUES(Table[Column])`),
            // prevent `RELATED` from reading join key columns that are not visible in the row
            // context.
            if hop_idx == 0 {
                if let Some(visible_cols) = current_visible_cols {
                    if !visible_cols.contains(&rel_info.from_idx) {
                        return Err(DaxError::Eval(format!(
                            "column {current_table}[{}] is not available in the current row context",
                            rel_info.rel.from_column
                        )));
                    }
                }
            }

            let from_table = model
                .table(&rel_info.rel.from_table)
                .ok_or_else(|| DaxError::UnknownTable(rel_info.rel.from_table.clone()))?;
            let key = from_table
                .value_by_idx(row, rel_info.from_idx)
                .unwrap_or(Value::Blank);
            if key.is_blank() {
                return Ok(Value::Blank);
            }

            let to_table = model
                .table(&rel_info.rel.to_table)
                .ok_or_else(|| DaxError::UnknownTable(rel_info.rel.to_table.clone()))?;
            let to_row = match &rel_info.to_index {
                ToIndex::RowSets { map, .. } => {
                    let Some(to_row_set) = map.get(&key) else {
                        return Ok(Value::Blank);
                    };
                    match to_row_set {
                        RowSet::One(row) => *row,
                        RowSet::Many(rows) => {
                            if rows.len() == 1 {
                                rows[0]
                            } else {
                                return Err(DaxError::Eval(format!(
                                    "RELATED is ambiguous: key {key} matches multiple rows in {} (relationship {})",
                                    rel_info.rel.to_table, rel_info.rel.name
                                )));
                            }
                        }
                    }
                }
                ToIndex::KeySet { keys, .. } => {
                    if !keys.contains(&key) {
                        return Ok(Value::Blank);
                    }
                    let rows = to_table
                        .filter_eq(rel_info.to_idx, &key)
                        .unwrap_or_else(|| {
                            let mut out = Vec::new();
                            for row in 0..to_table.row_count() {
                                let v = to_table
                                    .value_by_idx(row, rel_info.to_idx)
                                    .unwrap_or(Value::Blank);
                                if v == key {
                                    out.push(row);
                                }
                            }
                            out
                        });
                    match rows.as_slice() {
                        [] => return Ok(Value::Blank),
                        [row] => *row,
                        _ => {
                            return Err(DaxError::Eval(format!(
                                "RELATED is ambiguous: key {key} matches multiple rows in {} (relationship {})",
                                rel_info.rel.to_table, rel_info.rel.name
                            )))
                        }
                    }
                }
            };

            row = to_row;
            // `rel_info.rel.to_table` becomes the "current" table for the next hop.
        }

        let to_table = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
        let value = to_table
            .value(row, column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: table.clone(),
                column: column.clone(),
            })?;
        Ok(value)
    }

    fn eval_table(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
        row_ctx: &RowContext,
        env: &mut VarEnv,
    ) -> DaxResult<TableResult> {
        match expr {
            Expr::Let { bindings, body } => {
                env.push_scope();
                let result = (|| -> DaxResult<TableResult> {
                    for (name, binding_expr) in bindings {
                        let value =
                            self.eval_var_value(model, binding_expr, filter, row_ctx, env)?;
                        env.define(name, value);
                    }
                    self.eval_table(model, body, filter, row_ctx, env)
                })();
                env.pop_scope();
                result
            }
            Expr::TableName(name) => match env.lookup(name) {
                Some(VarValue::Table(t)) => Ok(t.clone()),
                Some(VarValue::Scalar(_)) => Err(DaxError::Type(format!(
                    "scalar variable {name} used in table context"
                ))),
                Some(VarValue::OneColumnTable(values)) => Ok(TableResult::Virtual {
                    // DAX table constructors expose a single implicit column named `Value`.
                    columns: vec![("__TABLE_LITERAL__".to_string(), "Value".to_string())],
                    rows: values.iter().cloned().map(|v| vec![v]).collect(),
                }),
                None => {
                    let table_key = normalize_ident(name);
                    let table_ref = model
                        .table(name)
                        .ok_or_else(|| DaxError::UnknownTable(name.clone()))?;

                    if filter.is_empty() {
                        Ok(TableResult::PhysicalAll {
                            table: table_key,
                            row_count: table_ref.row_count(),
                            visible_cols: None,
                        })
                    } else {
                        let sets = resolve_row_sets(model, filter)?;
                        let allowed = sets
                            .get(&table_key)
                            .ok_or_else(|| DaxError::UnknownTable(name.clone()))?;
                        if allowed.all_true() {
                            Ok(TableResult::PhysicalAll {
                                table: table_key,
                                row_count: table_ref.row_count(),
                                visible_cols: None,
                            })
                        } else {
                            let visible = allowed.count_ones();
                            let row_count = table_ref.row_count();
                            let sparse_to_dense_threshold = row_count / 64;
                            if visible <= sparse_to_dense_threshold {
                                Ok(TableResult::Physical {
                                    table: table_key,
                                    rows: allowed.iter_ones().collect(),
                                    visible_cols: None,
                                })
                            } else {
                                Ok(TableResult::PhysicalMask {
                                    table: table_key,
                                    mask: Arc::new(allowed.clone()),
                                    visible_cols: None,
                                })
                            }
                        }
                    }
                }
            },
            Expr::TableLiteral { .. } => {
                let Expr::TableLiteral { rows } = expr else {
                    unreachable!();
                };

                let col_count = rows.first().map(|row| row.len()).unwrap_or(1);
                for row in rows {
                    if row.len() != col_count {
                        return Err(DaxError::Type(
                            "table constructor rows must all have the same number of values".into(),
                        ));
                    }
                }

                let columns: Vec<(String, String)> = if col_count == 1 {
                    // DAX one-column table constructors expose a single implicit column named
                    // `Value`.
                    vec![("__TABLE_LITERAL__".to_string(), "Value".to_string())]
                } else {
                    // Multi-column table constructors expose synthetic columns `Value1`, `Value2`,
                    // ... in order.
                    (1..=col_count)
                        .map(|idx| ("__TABLE_LITERAL__".to_string(), format!("Value{idx}")))
                        .collect()
                };

                let mut out_rows: Vec<Vec<Value>> = Vec::with_capacity(rows.len());
                for row in rows {
                    let mut out_row: Vec<Value> = Vec::with_capacity(col_count);
                    for cell in row {
                        out_row.push(self.eval_scalar(model, cell, filter, row_ctx, env)?);
                    }
                    out_rows.push(out_row);
                }

                Ok(TableResult::Virtual {
                    columns,
                    rows: out_rows,
                })
            }
            Expr::Call { name, args } => match name.to_ascii_uppercase().as_str() {
                "FILTER" => {
                    let [table_expr, predicate] = args.as_slice() else {
                        return Err(DaxError::Eval("FILTER expects 2 arguments".into()));
                    };
                    let base = self.eval_table(model, table_expr, filter, row_ctx, env)?;

                    struct FilterRowsBuilder {
                        row_count: usize,
                        threshold: usize,
                        sparse: Vec<usize>,
                        dense: Option<BitVec>,
                    }

                    impl FilterRowsBuilder {
                        fn new(row_count: usize) -> Self {
                            Self {
                                row_count,
                                threshold: row_count / 64,
                                sparse: Vec::new(),
                                dense: None,
                            }
                        }

                        fn push(&mut self, row: usize) {
                            match &mut self.dense {
                                Some(mask) => mask.set(row, true),
                                None => {
                                    self.sparse.push(row);
                                    if self.sparse.len() > self.threshold {
                                        let mut mask = BitVec::with_len_all_false(self.row_count);
                                        for &row in &self.sparse {
                                            mask.set(row, true);
                                        }
                                        self.sparse.clear();
                                        self.dense = Some(mask);
                                    }
                                }
                            }
                        }

                        fn finish(
                            self,
                            table: String,
                            visible_cols: Option<Vec<usize>>,
                        ) -> TableResult {
                            match self.dense {
                                Some(mask) => TableResult::PhysicalMask {
                                    table,
                                    mask: Arc::new(mask),
                                    visible_cols,
                                },
                                None => TableResult::Physical {
                                    table,
                                    rows: self.sparse,
                                    visible_cols,
                                },
                            }
                        }
                    }

                    match base {
                        TableResult::Physical {
                            table,
                            rows,
                            visible_cols,
                        } => {
                            if visible_cols.is_some() {
                                let mut out_rows = Vec::new();
                                for row in rows.iter().copied() {
                                    let mut inner_ctx = row_ctx.clone();
                                    inner_ctx.push_physical(&table, row, visible_cols.clone());
                                    let pred = self
                                        .eval_scalar(model, predicate, filter, &inner_ctx, env)?;
                                    if pred.truthy().map_err(|e| DaxError::Type(e.to_string()))? {
                                        out_rows.push(row);
                                    }
                                }
                                Ok(TableResult::Physical {
                                    table,
                                    rows: out_rows,
                                    visible_cols,
                                })
                            } else {
                                let row_count = model
                                    .table(&table)
                                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?
                                    .row_count();

                                // Some physical table expressions (e.g. RELATEDTABLE) can include a
                                // relationship-generated virtual blank row at `row_count`. That
                                // row cannot be represented in `PhysicalMask` (which is sized to
                                // physical rows only), so fall back to the sparse representation.
                                if rows.iter().any(|&row| row >= row_count) {
                                    let mut out_rows = Vec::new();
                                    for row in rows.iter().copied() {
                                        let mut inner_ctx = row_ctx.clone();
                                        inner_ctx.push_physical(&table, row, None);
                                        let pred = self.eval_scalar(
                                            model, predicate, filter, &inner_ctx, env,
                                        )?;
                                        if pred
                                            .truthy()
                                            .map_err(|e| DaxError::Type(e.to_string()))?
                                        {
                                            out_rows.push(row);
                                        }
                                    }
                                    return Ok(TableResult::Physical {
                                        table,
                                        rows: out_rows,
                                        visible_cols: None,
                                    });
                                }

                                let mut builder = FilterRowsBuilder::new(row_count);
                                for row in rows.iter().copied() {
                                    let mut inner_ctx = row_ctx.clone();
                                    inner_ctx.push_physical(&table, row, None);
                                    let pred = self
                                        .eval_scalar(model, predicate, filter, &inner_ctx, env)?;
                                    if pred.truthy().map_err(|e| DaxError::Type(e.to_string()))? {
                                        builder.push(row);
                                    }
                                }
                                Ok(builder.finish(table, None))
                            }
                        }
                        TableResult::PhysicalAll {
                            table,
                            row_count,
                            visible_cols,
                        } => {
                            if visible_cols.is_some() {
                                let mut out_rows = Vec::new();
                                for row in 0..row_count {
                                    let mut inner_ctx = row_ctx.clone();
                                    inner_ctx.push_physical(&table, row, visible_cols.clone());
                                    let pred = self
                                        .eval_scalar(model, predicate, filter, &inner_ctx, env)?;
                                    if pred.truthy().map_err(|e| DaxError::Type(e.to_string()))? {
                                        out_rows.push(row);
                                    }
                                }
                                Ok(TableResult::Physical {
                                    table,
                                    rows: out_rows,
                                    visible_cols,
                                })
                            } else {
                                let mut builder = FilterRowsBuilder::new(row_count);
                                for row in 0..row_count {
                                    let mut inner_ctx = row_ctx.clone();
                                    inner_ctx.push_physical(&table, row, None);
                                    let pred = self
                                        .eval_scalar(model, predicate, filter, &inner_ctx, env)?;
                                    if pred.truthy().map_err(|e| DaxError::Type(e.to_string()))? {
                                        builder.push(row);
                                    }
                                }
                                Ok(builder.finish(table, None))
                            }
                        }
                        TableResult::PhysicalMask {
                            table,
                            mask,
                            visible_cols,
                        } => {
                            if visible_cols.is_some() {
                                let mut out_rows = Vec::new();
                                for row in mask.iter_ones() {
                                    let mut inner_ctx = row_ctx.clone();
                                    inner_ctx.push_physical(&table, row, visible_cols.clone());
                                    let pred = self
                                        .eval_scalar(model, predicate, filter, &inner_ctx, env)?;
                                    if pred.truthy().map_err(|e| DaxError::Type(e.to_string()))? {
                                        out_rows.push(row);
                                    }
                                }
                                Ok(TableResult::Physical {
                                    table,
                                    rows: out_rows,
                                    visible_cols,
                                })
                            } else {
                                let row_count = mask.len();
                                let mut builder = FilterRowsBuilder::new(row_count);
                                for row in mask.iter_ones() {
                                    let mut inner_ctx = row_ctx.clone();
                                    inner_ctx.push_physical(&table, row, None);
                                    let pred = self
                                        .eval_scalar(model, predicate, filter, &inner_ctx, env)?;
                                    if pred.truthy().map_err(|e| DaxError::Type(e.to_string()))? {
                                        builder.push(row);
                                    }
                                }
                                Ok(builder.finish(table, None))
                            }
                        }
                        TableResult::Virtual { columns, rows } => {
                            let mut out_rows = Vec::new();
                            for row_values in rows.into_iter() {
                                let mut inner_ctx = row_ctx.clone();
                                let bindings: Vec<((String, String), Value)> = columns
                                    .iter()
                                    .cloned()
                                    .zip(row_values.iter().cloned())
                                    .map(|(col, v)| (col, v))
                                    .collect();
                                inner_ctx.push_virtual(bindings);
                                let pred =
                                    self.eval_scalar(model, predicate, filter, &inner_ctx, env)?;
                                if pred.truthy().map_err(|e| DaxError::Type(e.to_string()))? {
                                    out_rows.push(row_values);
                                }
                            }
                            Ok(TableResult::Virtual {
                                columns,
                                rows: out_rows,
                            })
                        }
                    }
                }
                "ALL" => {
                    let [arg] = args.as_slice() else {
                        return Err(DaxError::Eval("ALL expects 1 argument".into()));
                    };
                    match arg {
                        Expr::TableName(name) => {
                            let table_ref = model
                                .table(name)
                                .ok_or_else(|| DaxError::UnknownTable(name.clone()))?;
                            Ok(TableResult::PhysicalAll {
                                table: normalize_ident(name),
                                row_count: table_ref.row_count(),
                                visible_cols: None,
                            })
                        }
                        Expr::ColumnRef { table, column } => {
                            // `ALL(Table[Column])` removes filters from the target column but
                            // preserves other filters on the same table.
                            let mut modified_filter = filter.clone();
                            modified_filter.clear_column_filter(table, column);

                            let table_ref = model
                                .table(table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                            let idx = table_ref.column_idx(column).ok_or_else(|| {
                                DaxError::UnknownColumn {
                                    table: table.clone(),
                                    column: column.clone(),
                                }
                            })?;

                            let mut seen = HashSet::new();
                            let mut rows = Vec::new();
                            let sets = (!modified_filter.is_empty())
                                .then(|| resolve_row_sets(model, &modified_filter))
                                .transpose()?;
                            let table_key = normalize_ident(table);
                            if let Some(sets) = sets.as_ref() {
                                let allowed = sets
                                    .get(&table_key)
                                    .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
                                for row in allowed.iter_ones() {
                                    let value =
                                        table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                    if seen.insert(value) {
                                        rows.push(row);
                                    }
                                }
                            } else {
                                for row in 0..table_ref.row_count() {
                                    let value =
                                        table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                    if seen.insert(value) {
                                        rows.push(row);
                                    }
                                }
                            }
                            // If the table participates as the one-side of a relationship and has
                            // unmatched fact-side keys, tabular models materialize an "unknown"
                            // (blank) member. Include that member when it exists and is not
                            // excluded by the remaining filters.
                            if !seen.contains(&Value::Blank)
                                && blank_row_allowed(&modified_filter, table)
                                && virtual_blank_row_exists(
                                    model,
                                    &modified_filter,
                                    table,
                                    sets.as_ref(),
                                )?
                            {
                                rows.push(table_ref.row_count());
                            }
                            Ok(TableResult::Physical {
                                table: normalize_ident(table),
                                rows,
                                visible_cols: Some(vec![idx]),
                            })
                        }
                        other => Err(DaxError::Type(format!(
                            "ALL expects a table name or column reference, got {other:?}"
                        ))),
                    }
                }
                "ALLNOBLANKROW" => {
                    let [arg] = args.as_slice() else {
                        return Err(DaxError::Eval("ALLNOBLANKROW expects 1 argument".into()));
                    };
                    match arg {
                        Expr::TableName(name) => {
                            // Like `ALL(Table)`, return all physical rows (excluding any
                            // relationship-generated blank member).
                            let table_key = normalize_ident(name);
                            let table_ref = model
                                .table(name)
                                .ok_or_else(|| DaxError::UnknownTable(name.clone()))?;
                            Ok(TableResult::PhysicalAll {
                                table: table_key,
                                row_count: table_ref.row_count(),
                                visible_cols: None,
                            })
                        }
                        Expr::ColumnRef { table, column } => {
                            // Like `ALL(Table[Column])`, but exclude both:
                            //   - physical blank values in the column
                            //   - the relationship-generated "unknown" (blank) member
                            let table_ref = model
                                .table(table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                            let idx = table_ref.column_idx(column).ok_or_else(|| {
                                DaxError::UnknownColumn {
                                    table: table.clone(),
                                    column: column.clone(),
                                }
                            })?;

                            let mut modified_filter = filter.clone();
                            modified_filter.clear_column_filter(table, column);

                            let mut seen = HashSet::new();
                            let mut rows = Vec::new();
                            let row_sets = (!modified_filter.is_empty())
                                .then(|| resolve_row_sets(model, &modified_filter))
                                .transpose()?;
                            let table_key = normalize_ident(table);
                            if let Some(sets) = row_sets.as_ref() {
                                let allowed = sets
                                    .get(&table_key)
                                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                                for row in allowed.iter_ones() {
                                    let value =
                                        table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                    if value.is_blank() {
                                        continue;
                                    }
                                    if seen.insert(value) {
                                        rows.push(row);
                                    }
                                }
                            } else {
                                for row in 0..table_ref.row_count() {
                                    let value =
                                        table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                    if value.is_blank() {
                                        continue;
                                    }
                                    if seen.insert(value) {
                                        rows.push(row);
                                    }
                                }
                            }

                            Ok(TableResult::Physical {
                                table: normalize_ident(table),
                                rows,
                                visible_cols: Some(vec![idx]),
                            })
                        }
                        other => Err(DaxError::Type(format!(
                            "ALLNOBLANKROW expects a table name or column reference, got {other:?}"
                        ))),
                    }
                }
                "VALUES" => {
                    let [arg] = args.as_slice() else {
                        return Err(DaxError::Eval("VALUES expects 1 argument".into()));
                    };
                    match arg {
                        Expr::ColumnRef { table, column } => {
                            let table_ref = model
                                .table(table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                            let idx = table_ref.column_idx(column).ok_or_else(|| {
                                DaxError::UnknownColumn {
                                    table: table.clone(),
                                    column: column.clone(),
                                }
                            })?;

                            let mut seen = HashSet::new();
                            let mut rows = Vec::new();
                            let sets = (!filter.is_empty())
                                .then(|| resolve_row_sets(model, filter))
                                .transpose()?;
                            let table_key = normalize_ident(table);
                            if let Some(sets) = sets.as_ref() {
                                let allowed = sets
                                    .get(&table_key)
                                    .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
                                for row in allowed.iter_ones() {
                                    let value =
                                        table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                    if seen.insert(value) {
                                        rows.push(row);
                                    }
                                }
                            } else {
                                for row in 0..table_ref.row_count() {
                                    let value =
                                        table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                    if seen.insert(value) {
                                        rows.push(row);
                                    }
                                }
                            }
                            if !seen.contains(&Value::Blank)
                                && blank_row_allowed(filter, table)
                                && virtual_blank_row_exists(model, filter, table, sets.as_ref())?
                            {
                                rows.push(table_ref.row_count());
                            }
                            Ok(TableResult::Physical {
                                table: normalize_ident(table),
                                rows,
                                visible_cols: Some(vec![idx]),
                            })
                        }
                        _ => {
                            let base = self.eval_table(model, arg, filter, row_ctx, env)?;
                            distinct_rows_by_all_columns(model, &base)
                        }
                    }
                }
                "DISTINCT" => {
                    let [arg] = args.as_slice() else {
                        return Err(DaxError::Eval("DISTINCT expects 1 argument".into()));
                    };
                    match arg {
                        Expr::ColumnRef { table, column } => self.eval_table(
                            model,
                            &Expr::Call {
                                name: "VALUES".to_string(),
                                args: vec![Expr::ColumnRef {
                                    table: table.clone(),
                                    column: column.clone(),
                                }],
                            },
                            filter,
                            row_ctx,
                            env,
                        ),
                        _ => {
                            let base = self.eval_table(model, arg, filter, row_ctx, env)?;
                            distinct_rows_by_all_columns(model, &base)
                        }
                    }
                }
                "ALLEXCEPT" => {
                    let (table_expr, keep_cols) = args.split_first().ok_or_else(|| {
                        DaxError::Eval("ALLEXCEPT expects at least 2 arguments".into())
                    })?;
                    if keep_cols.is_empty() {
                        return Err(DaxError::Eval(
                            "ALLEXCEPT expects at least 2 arguments".into(),
                        ));
                    }

                    let Expr::TableName(table) = table_expr else {
                        return Err(DaxError::Type(
                            "ALLEXCEPT expects a table name as the first argument".into(),
                        ));
                    };
                    let table_key = normalize_ident(table);

                    let mut keep: HashSet<String> = HashSet::new();
                    for expr in keep_cols {
                        let Expr::ColumnRef {
                            table: col_table,
                            column,
                        } = expr
                        else {
                            return Err(DaxError::Type(
                                "ALLEXCEPT expects column references after the table name".into(),
                            ));
                        };
                        if normalize_ident(col_table) != table_key {
                            return Err(DaxError::Eval(format!(
                                "ALLEXCEPT column must belong to {table}, got {col_table}[{column}]",
                            )));
                        }
                        keep.insert(normalize_ident(column));
                    }

                    let mut modified_filter = filter.clone();
                    modified_filter.clear_table_filters(table);
                    for ((t, c), values) in &filter.column_filters {
                        if t == &table_key && keep.contains(c) {
                            modified_filter
                                .column_filters
                                .insert((t.clone(), c.clone()), values.clone());
                        }
                    }

                    let table_ref = model
                        .table(table)
                        .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                    if modified_filter.is_empty() {
                        return Ok(TableResult::PhysicalAll {
                            table: table_key,
                            row_count: table_ref.row_count(),
                            visible_cols: None,
                        });
                    }

                    let sets = resolve_row_sets(model, &modified_filter)?;
                    let allowed = sets
                        .get(&table_key)
                        .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                    if allowed.all_true() {
                        Ok(TableResult::PhysicalAll {
                            table: table_key,
                            row_count: table_ref.row_count(),
                            visible_cols: None,
                        })
                    } else {
                        let visible = allowed.count_ones();
                        let row_count = table_ref.row_count();
                        let sparse_to_dense_threshold = row_count / 64;
                        if visible <= sparse_to_dense_threshold {
                            Ok(TableResult::Physical {
                                table: table_key,
                                rows: allowed.iter_ones().collect(),
                                visible_cols: None,
                            })
                        } else {
                            Ok(TableResult::PhysicalMask {
                                table: table_key,
                                mask: Arc::new(allowed.clone()),
                                visible_cols: None,
                            })
                        }
                    }
                }
                "CALCULATETABLE" => {
                    if args.is_empty() {
                        return Err(DaxError::Eval(
                            "CALCULATETABLE expects at least 1 argument".into(),
                        ));
                    }
                    let (table_expr, filter_args) = args.split_first().expect("checked above");
                    let new_filter =
                        self.build_calculate_filter(model, filter, row_ctx, filter_args, env)?;
                    let mut table_filter = new_filter;
                    table_filter.suppress_implicit_measure_context_transition = true;
                    self.eval_table(model, table_expr, &table_filter, row_ctx, env)
                }
                "SUMMARIZE" => {
                    let (table_expr, group_exprs) = args.split_first().ok_or_else(|| {
                        DaxError::Eval("SUMMARIZE expects at least 2 arguments".into())
                    })?;
                    if group_exprs.is_empty() {
                        return Err(DaxError::Eval(
                            "SUMMARIZE expects at least 2 arguments".into(),
                        ));
                    }

                    let base = self.eval_table(model, table_expr, filter, row_ctx, env)?;
                    let base_table = match &base {
                        TableResult::Physical { table, .. }
                        | TableResult::PhysicalAll { table, .. }
                        | TableResult::PhysicalMask { table, .. } => table.clone(),
                        TableResult::Virtual { .. } => {
                            return Err(DaxError::Type(
                                "SUMMARIZE currently only supports a physical base table".into(),
                            ))
                        }
                    };

                    let table_ref = model
                        .table(&base_table)
                        .ok_or_else(|| DaxError::UnknownTable(base_table.clone()))?;

                    let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
                    for &idx in filter.relationship_overrides() {
                        if let Some(rel) = model.relationships().get(idx) {
                            override_pairs
                                .insert((rel.from_table_key.as_str(), rel.to_table_key.as_str()));
                        }
                    }

                    let is_relationship_active =
                        |idx: usize, rel: &RelationshipInfo, overrides: &HashSet<(&str, &str)>| {
                            let pair = (rel.from_table_key.as_str(), rel.to_table_key.as_str());
                            let is_active = if overrides.contains(&pair) {
                                filter.relationship_overrides().contains(&idx)
                            } else {
                                rel.rel.is_active
                            };
                            is_active && !filter.is_relationship_disabled(idx)
                        };

                    #[derive(Clone, Copy)]
                    struct Hop {
                        relationship_idx: usize,
                        from_idx: usize,
                    }

                    enum GroupAccessor {
                        BaseColumn(usize),
                        RelatedPath { hops: Vec<Hop>, to_col_idx: usize },
                    }

                    let mut out_columns: Vec<(String, String)> =
                        Vec::with_capacity(group_exprs.len());
                    let mut accessors = Vec::with_capacity(group_exprs.len());
                    for expr in group_exprs {
                        let Expr::ColumnRef { table, column } = expr else {
                            return Err(DaxError::Type(
                                "SUMMARIZE currently only supports grouping by columns".into(),
                            ));
                        };
                        let group_table_key = normalize_ident(table);
                        out_columns.push((group_table_key.clone(), normalize_ident(column)));
                        if group_table_key != base_table {
                            let Some(path) = model.find_unique_active_relationship_path(
                                &base_table,
                                &group_table_key,
                                RelationshipPathDirection::ManyToOne,
                                |idx, rel| is_relationship_active(idx, rel, &override_pairs),
                            )?
                            else {
                                return Err(DaxError::Eval(format!(
                                    "SUMMARIZE grouping column {table}[{column}] is not reachable from {base_table}"
                                )));
                            };

                            let mut hops: Vec<Hop> = Vec::with_capacity(path.len());
                            for rel_idx in path {
                                let rel_info = model
                                    .relationships()
                                    .get(rel_idx)
                                    .expect("relationship index from path");

                                let from_table_ref =
                                    model.table(&rel_info.rel.from_table).ok_or_else(|| {
                                        DaxError::UnknownTable(rel_info.rel.from_table.clone())
                                    })?;
                                let from_idx = from_table_ref
                                    .column_idx(&rel_info.rel.from_column)
                                    .ok_or_else(|| DaxError::UnknownColumn {
                                        table: rel_info.rel.from_table.clone(),
                                        column: rel_info.rel.from_column.clone(),
                                    })?;

                                hops.push(Hop {
                                    relationship_idx: rel_idx,
                                    from_idx,
                                });
                            }

                            let to_table_ref = model
                                .table(table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                            let to_col_idx = to_table_ref.column_idx(column).ok_or_else(|| {
                                DaxError::UnknownColumn {
                                    table: table.clone(),
                                    column: column.clone(),
                                }
                            })?;

                            accessors.push(GroupAccessor::RelatedPath { hops, to_col_idx });
                            continue;
                        }
                        let idx = table_ref.column_idx(column).ok_or_else(|| {
                            DaxError::UnknownColumn {
                                table: table.clone(),
                                column: column.clone(),
                            }
                        })?;
                        accessors.push(GroupAccessor::BaseColumn(idx));
                    }

                    let row_sets = resolve_row_sets(model, filter)?;

                    #[derive(Default)]
                    struct PathNode {
                        columns: Vec<(usize, usize)>,
                        children: Vec<(Hop, PathNode)>,
                    }

                    fn child_node_mut<'a>(node: &'a mut PathNode, hop: Hop) -> &'a mut PathNode {
                        if let Some(idx) = node
                            .children
                            .iter()
                            .position(|(h, _)| h.relationship_idx == hop.relationship_idx)
                        {
                            return &mut node.children[idx].1;
                        }

                        node.children.push((hop, PathNode::default()));
                        let idx = node.children.len().saturating_sub(1);
                        &mut node.children[idx].1
                    }

                    fn next_rows_for_hop(
                        model: &DataModel,
                        row_sets: &HashMap<String, BitVec>,
                        table: &crate::model::Table,
                        row: usize,
                        hop: Hop,
                    ) -> DaxResult<Vec<usize>> {
                        let key = table
                            .value_by_idx(row, hop.from_idx)
                            .unwrap_or(Value::Blank);
                        if key.is_blank() {
                            return Ok(Vec::new());
                        }

                        let rel_info = model
                            .relationships()
                            .get(hop.relationship_idx)
                            .expect("valid relationship index");

                        let allowed_to = row_sets
                            .get(rel_info.to_table_key.as_str())
                            .ok_or_else(|| DaxError::UnknownTable(rel_info.rel.to_table.clone()))?;

                        let mut rows: Vec<usize> = Vec::new();
                        match &rel_info.to_index {
                            ToIndex::RowSets { map, .. } => {
                                let Some(to_row_set) = map.get(&key) else {
                                    return Ok(Vec::new());
                                };
                                to_row_set.for_each_row(|to_row: usize| {
                                    if to_row < allowed_to.len() && allowed_to.get(to_row) {
                                        rows.push(to_row);
                                    }
                                });
                            }
                            ToIndex::KeySet { .. } => {
                                // Fast path: if the key doesn't exist in the related table, there
                                // can be no matching rows.
                                if !rel_info.to_index.contains_key(&key) {
                                    return Ok(Vec::new());
                                }

                                let to_table_ref =
                                    model.table(rel_info.rel.to_table.as_str()).ok_or_else(
                                        || DaxError::UnknownTable(rel_info.rel.to_table.clone()),
                                    )?;

                                if let Some(matches) = to_table_ref.filter_eq(rel_info.to_idx, &key)
                                {
                                    for to_row in matches {
                                        if to_row < allowed_to.len() && allowed_to.get(to_row) {
                                            rows.push(to_row);
                                        }
                                    }
                                } else {
                                    // Fallback: scan the allowed row set and compare values.
                                    for to_row in allowed_to.iter_ones() {
                                        let v = to_table_ref
                                            .value_by_idx(to_row, rel_info.to_idx)
                                            .unwrap_or(Value::Blank);
                                        if v == key {
                                            rows.push(to_row);
                                        }
                                    }
                                }
                            }
                        }
                        rows.sort_unstable();
                        rows.dedup();
                        Ok(rows)
                    }

                    fn collect_keys_for_node(
                        node: &PathNode,
                        model: &DataModel,
                        row_sets: &HashMap<String, BitVec>,
                        table: &crate::model::Table,
                        rows: &[usize],
                        key_template: Vec<Value>,
                    ) -> DaxResult<Vec<Vec<Value>>> {
                        let mut results: Vec<Vec<Value>> = Vec::new();

                        let row_opts: Vec<Option<usize>> = if rows.is_empty() {
                            vec![None]
                        } else {
                            rows.iter().copied().map(Some).collect()
                        };

                        for row_opt in row_opts {
                            let mut key = key_template.clone();
                            for (pos, col_idx) in &node.columns {
                                key[*pos] = match row_opt {
                                    Some(row) => {
                                        table.value_by_idx(row, *col_idx).unwrap_or(Value::Blank)
                                    }
                                    None => Value::Blank,
                                };
                            }

                            let mut partials = vec![key];
                            for (hop, child) in &node.children {
                                let rel_info = model
                                    .relationships()
                                    .get(hop.relationship_idx)
                                    .expect("valid relationship index");
                                let to_table_ref =
                                    model.table(rel_info.rel.to_table.as_str()).ok_or_else(
                                        || DaxError::UnknownTable(rel_info.rel.to_table.clone()),
                                    )?;

                                let child_rows = match row_opt {
                                    Some(row) => {
                                        next_rows_for_hop(model, row_sets, table, row, *hop)?
                                    }
                                    None => Vec::new(),
                                };

                                let mut next: Vec<Vec<Value>> = Vec::new();
                                for partial in partials {
                                    next.extend(collect_keys_for_node(
                                        child,
                                        model,
                                        row_sets,
                                        to_table_ref,
                                        &child_rows,
                                        partial,
                                    )?);
                                }
                                partials = next;
                                if partials.is_empty() {
                                    break;
                                }
                            }

                            results.extend(partials);
                        }

                        Ok(results)
                    }

                    // Build a trie of relationship paths so group keys stay correlated across
                    // snowflake hops.
                    let mut root: PathNode = PathNode::default();
                    for (pos, accessor) in accessors.iter().enumerate() {
                        match accessor {
                            GroupAccessor::BaseColumn(idx) => root.columns.push((pos, *idx)),
                            GroupAccessor::RelatedPath { hops, to_col_idx } => {
                                let mut node = &mut root;
                                for hop in hops {
                                    node = child_node_mut(node, *hop);
                                }
                                node.columns.push((pos, *to_col_idx));
                            }
                        }
                    }

                    let blank_key = vec![Value::Blank; accessors.len()];
                    let mut seen: HashSet<Vec<Value>> = HashSet::new();
                    let mut out_rows: Vec<Vec<Value>> = Vec::new();
                    for row_handle in base.iter_rows() {
                        let RowHandle::Physical(row) = row_handle else {
                            unreachable!("SUMMARIZE base table is always physical");
                        };
                        for key in collect_keys_for_node(
                            &root,
                            model,
                            &row_sets,
                            table_ref,
                            std::slice::from_ref(&row),
                            blank_key.clone(),
                        )? {
                            if seen.insert(key.clone()) {
                                out_rows.push(key);
                            }
                        }
                    }

                    Ok(TableResult::Virtual {
                        columns: out_columns,
                        rows: out_rows,
                    })
                }
                "SUMMARIZECOLUMNS" => {
                    // MVP: only support grouping columns and (optionally) CALCULATE-style filter
                    // arguments. Name/expression pairs ("Name", expr) are parsed/validated but are
                    // not currently materialized in the resulting table (the engine returns a row
                    // set of the chosen base table).
                    let mut group_cols: Vec<(String, String)> = Vec::new();
                    let mut group_tables: HashSet<String> = HashSet::new();
                    let mut arg_idx = 0usize;
                    while arg_idx < args.len() {
                        match &args[arg_idx] {
                            Expr::ColumnRef { table, column } => {
                                let table_key = normalize_ident(table);
                                let column_key = normalize_ident(column);
                                group_tables.insert(table_key.clone());
                                group_cols.push((table_key, column_key));
                                arg_idx += 1;
                            }
                            _ => break,
                        }
                    }

                    if group_cols.is_empty() {
                        return Err(DaxError::Eval(
                            "SUMMARIZECOLUMNS expects at least 1 grouping column".into(),
                        ));
                    }

                    // After grouping columns, SUMMARIZECOLUMNS accepts:
                    //   - zero or more filter table arguments (table expressions)
                    //   - zero or more name/expression pairs ("Name", expr)
                    // We detect the transition to name/expression pairs by finding the first
                    // string literal argument.
                    let mut name_start = args.len();
                    for (idx, arg) in args.iter().enumerate().skip(arg_idx) {
                        if matches!(arg, Expr::Text(_)) {
                            name_start = idx;
                            break;
                        }
                    }

                    let filter_args = &args[arg_idx..name_start];
                    let name_expr_args = &args[name_start..];
                    if !name_expr_args.is_empty() {
                        if name_expr_args.len() % 2 != 0 {
                            return Err(DaxError::Eval(
                                "SUMMARIZECOLUMNS name/expression pairs must come in (\"Name\", expr) pairs".into(),
                            ));
                        }
                        for pair in name_expr_args.chunks(2) {
                            match &pair[0] {
                                Expr::Text(_) => {}
                                other => {
                                    return Err(DaxError::Type(format!(
                                        "SUMMARIZECOLUMNS expected a string literal for the name/expression pair name, got {other:?}"
                                    )))
                                }
                            }
                            // `pair[1]` is a scalar expression; we intentionally don't evaluate it
                            // yet because the current table representation can't materialize the
                            // resulting column.
                        }
                    }

                    // Apply filter arguments with CALCULATE-style semantics, but do **not** perform
                    // an implicit context transition from the current row context. This matches
                    // DAX behavior: row context does not become filter context unless an explicit
                    // `CALCULATE`/`CALCULATETABLE` is invoked (or a measure is evaluated).
                    let mut summarize_filter = filter.clone();
                    if !filter_args.is_empty() {
                        self.apply_calculate_filter_args(
                            model,
                            &mut summarize_filter,
                            row_ctx,
                            filter_args,
                            env,
                        )?;
                    }

                    let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
                    for &idx in &summarize_filter.active_relationship_overrides {
                        if let Some(rel) = model.relationships().get(idx) {
                            override_pairs
                                .insert((rel.rel.from_table.as_str(), rel.rel.to_table.as_str()));
                        }
                    }

                    let is_relationship_active =
                        |idx: usize, rel: &RelationshipInfo, overrides: &HashSet<(&str, &str)>| {
                            let pair = (rel.rel.from_table.as_str(), rel.rel.to_table.as_str());
                            let is_active = if overrides.contains(&pair) {
                                summarize_filter
                                    .active_relationship_overrides
                                    .contains(&idx)
                            } else {
                                rel.rel.is_active
                            };
                            if !is_active {
                                return false;
                            }
                            !matches!(
                                summarize_filter.cross_filter_overrides.get(&idx).copied(),
                                Some(RelationshipOverride::Disabled)
                            )
                        };

                    // Determine the base table to scan for groups.
                    let base_table = if group_tables.len() == 1 {
                        group_tables.iter().next().expect("len==1").clone()
                    } else {
                        let mut tables_vec: Vec<&String> = group_tables.iter().collect();
                        tables_vec.sort();
                        let groups_list = tables_vec
                            .iter()
                            .map(|t| t.as_str())
                            .collect::<Vec<_>>()
                            .join(", ");

                        let mut candidates: Vec<String> = Vec::new();
                        let mut ambiguous_path_error: Option<DaxError> = None;
                        for table in model.tables.keys() {
                            let mut ok = true;
                            for target in &group_tables {
                                if target == table {
                                    continue;
                                }
                                match model.find_unique_active_relationship_path(
                                    table,
                                    target,
                                    RelationshipPathDirection::ManyToOne,
                                    |idx, rel| is_relationship_active(idx, rel, &override_pairs),
                                ) {
                                    Ok(Some(_)) => {}
                                    Ok(None) => {
                                        ok = false;
                                        break;
                                    }
                                    Err(err) => {
                                        ambiguous_path_error = Some(err);
                                        ok = false;
                                        break;
                                    }
                                }
                            }
                            if ok {
                                candidates.push(table.clone());
                            }
                        }
                        candidates.sort();

                        match candidates.len() {
                            1 => candidates[0].clone(),
                            0 => {
                                if let Some(err) = ambiguous_path_error {
                                    return Err(err);
                                }
                                return Err(DaxError::Eval(format!(
                                    "SUMMARIZECOLUMNS columns ({groups_list}) are not reachable from a single base table via active relationships"
                                )));
                            }
                            _ => {
                                return Err(DaxError::Eval(format!(
                                    "SUMMARIZECOLUMNS columns ({groups_list}) are reachable from multiple base tables: {}",
                                    candidates.join(", ")
                                )));
                            }
                        }
                    };

                    let base_table_ref = model
                        .table(&base_table)
                        .ok_or_else(|| DaxError::UnknownTable(base_table.clone()))?;

                    #[derive(Clone, Copy)]
                    struct Hop {
                        relationship_idx: usize,
                        from_idx: usize,
                    }

                    enum GroupAccessor {
                        BaseColumn(usize),
                        RelatedColumn { hops: Vec<Hop>, to_col_idx: usize },
                    }

                    let mut accessors = Vec::with_capacity(group_cols.len());
                    for (table, column) in &group_cols {
                        if table == &base_table {
                            let idx = base_table_ref.column_idx(column).ok_or_else(|| {
                                DaxError::UnknownColumn {
                                    table: table.clone(),
                                    column: column.clone(),
                                }
                            })?;
                            accessors.push(GroupAccessor::BaseColumn(idx));
                            continue;
                        }

                        let Some(path) = model.find_unique_active_relationship_path(
                            &base_table,
                            table,
                            RelationshipPathDirection::ManyToOne,
                            |idx, rel| is_relationship_active(idx, rel, &override_pairs),
                        )?
                        else {
                            return Err(DaxError::Eval(format!(
                                "SUMMARIZECOLUMNS grouping column {table}[{column}] is not reachable from {base_table}"
                            )));
                        };

                        let mut hops: Vec<Hop> = Vec::with_capacity(path.len());
                        for rel_idx in path {
                            let rel_info = model
                                .relationships()
                                .get(rel_idx)
                                .expect("relationship index from path");

                            let from_table_ref =
                                model.table(&rel_info.rel.from_table).ok_or_else(|| {
                                    DaxError::UnknownTable(rel_info.rel.from_table.clone())
                                })?;
                            let from_idx = from_table_ref
                                .column_idx(&rel_info.rel.from_column)
                                .ok_or_else(|| DaxError::UnknownColumn {
                                    table: rel_info.rel.from_table.clone(),
                                    column: rel_info.rel.from_column.clone(),
                                })?;

                            hops.push(Hop {
                                relationship_idx: rel_idx,
                                from_idx,
                            });
                        }

                        let target_ref = model
                            .table(table)
                            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                        let to_col_idx = target_ref.column_idx(column).ok_or_else(|| {
                            DaxError::UnknownColumn {
                                table: table.clone(),
                                column: column.clone(),
                            }
                        })?;

                        accessors.push(GroupAccessor::RelatedColumn { hops, to_col_idx });
                    }

                    // Resolve the current filter context once so we can reuse it both for the
                    // base table scan and for determining whether the relationship-generated blank
                    // member is visible.
                    let row_sets = resolve_row_sets(model, &summarize_filter)?;

                    let base_rows_set = row_sets
                        .get(&base_table)
                        .ok_or_else(|| DaxError::UnknownTable(base_table.clone()))?;
                    let include_virtual_blank_row = group_tables.len() == 1
                        && blank_row_allowed(&summarize_filter, &base_table)
                        && virtual_blank_row_exists(
                            model,
                            &summarize_filter,
                            &base_table,
                            Some(&row_sets),
                        )?;

                    #[derive(Default)]
                    struct PathNode {
                        columns: Vec<(usize, usize)>,
                        children: Vec<(Hop, PathNode)>,
                    }

                    fn child_node_mut<'a>(node: &'a mut PathNode, hop: Hop) -> &'a mut PathNode {
                        if let Some(idx) = node
                            .children
                            .iter()
                            .position(|(h, _)| h.relationship_idx == hop.relationship_idx)
                        {
                            return &mut node.children[idx].1;
                        }

                        node.children.push((hop, PathNode::default()));
                        let idx = node.children.len().saturating_sub(1);
                        &mut node.children[idx].1
                    }

                    fn next_rows_for_hop(
                        model: &DataModel,
                        row_sets: &HashMap<String, BitVec>,
                        table: &crate::model::Table,
                        row: usize,
                        hop: Hop,
                    ) -> DaxResult<Vec<usize>> {
                        let key = table
                            .value_by_idx(row, hop.from_idx)
                            .unwrap_or(Value::Blank);
                        if key.is_blank() {
                            return Ok(Vec::new());
                        }

                        let rel_info = model
                            .relationships()
                            .get(hop.relationship_idx)
                            .expect("valid relationship index");

                        let allowed_to = row_sets
                            .get(rel_info.to_table_key.as_str())
                            .ok_or_else(|| DaxError::UnknownTable(rel_info.rel.to_table.clone()))?;

                        let mut rows: Vec<usize> = Vec::new();
                        match &rel_info.to_index {
                            ToIndex::RowSets { map, .. } => {
                                let Some(to_row_set) = map.get(&key) else {
                                    return Ok(Vec::new());
                                };
                                to_row_set.for_each_row(|to_row: usize| {
                                    if to_row < allowed_to.len() && allowed_to.get(to_row) {
                                        rows.push(to_row);
                                    }
                                });
                            }
                            ToIndex::KeySet { .. } => {
                                if !rel_info.to_index.contains_key(&key) {
                                    return Ok(Vec::new());
                                }

                                let to_table_ref =
                                    model.table(rel_info.rel.to_table.as_str()).ok_or_else(
                                        || DaxError::UnknownTable(rel_info.rel.to_table.clone()),
                                    )?;

                                if let Some(matches) = to_table_ref.filter_eq(rel_info.to_idx, &key)
                                {
                                    for to_row in matches {
                                        if to_row < allowed_to.len() && allowed_to.get(to_row) {
                                            rows.push(to_row);
                                        }
                                    }
                                } else {
                                    for to_row in allowed_to.iter_ones() {
                                        let v = to_table_ref
                                            .value_by_idx(to_row, rel_info.to_idx)
                                            .unwrap_or(Value::Blank);
                                        if v == key {
                                            rows.push(to_row);
                                        }
                                    }
                                }
                            }
                        }
                        rows.sort_unstable();
                        rows.dedup();
                        Ok(rows)
                    }

                    fn collect_keys_for_node(
                        node: &PathNode,
                        model: &DataModel,
                        row_sets: &HashMap<String, BitVec>,
                        table: &crate::model::Table,
                        rows: &[usize],
                        key_template: Vec<Value>,
                    ) -> DaxResult<Vec<Vec<Value>>> {
                        let mut results: Vec<Vec<Value>> = Vec::new();

                        let row_opts: Vec<Option<usize>> = if rows.is_empty() {
                            vec![None]
                        } else {
                            rows.iter().copied().map(Some).collect()
                        };

                        for row_opt in row_opts {
                            let mut key = key_template.clone();
                            for (pos, col_idx) in &node.columns {
                                key[*pos] = match row_opt {
                                    Some(row) => {
                                        table.value_by_idx(row, *col_idx).unwrap_or(Value::Blank)
                                    }
                                    None => Value::Blank,
                                };
                            }

                            let mut partials = vec![key];
                            for (hop, child) in &node.children {
                                let rel_info = model
                                    .relationships()
                                    .get(hop.relationship_idx)
                                    .expect("valid relationship index");
                                let to_table_ref =
                                    model.table(rel_info.rel.to_table.as_str()).ok_or_else(
                                        || DaxError::UnknownTable(rel_info.rel.to_table.clone()),
                                    )?;

                                let child_rows = match row_opt {
                                    Some(row) => {
                                        next_rows_for_hop(model, row_sets, table, row, *hop)?
                                    }
                                    None => Vec::new(),
                                };

                                let mut next: Vec<Vec<Value>> = Vec::new();
                                for partial in partials {
                                    next.extend(collect_keys_for_node(
                                        child,
                                        model,
                                        row_sets,
                                        to_table_ref,
                                        &child_rows,
                                        partial,
                                    )?);
                                }
                                partials = next;
                                if partials.is_empty() {
                                    break;
                                }
                            }

                            results.extend(partials);
                        }

                        Ok(results)
                    }

                    // Build a trie of relationship paths so group keys stay correlated across
                    // snowflake hops.
                    let mut root: PathNode = PathNode::default();
                    for (pos, accessor) in accessors.iter().enumerate() {
                        match accessor {
                            GroupAccessor::BaseColumn(idx) => root.columns.push((pos, *idx)),
                            GroupAccessor::RelatedColumn { hops, to_col_idx } => {
                                let mut node = &mut root;
                                for hop in hops {
                                    node = child_node_mut(node, *hop);
                                }
                                node.columns.push((pos, *to_col_idx));
                            }
                        }
                    }

                    let blank_key = vec![Value::Blank; accessors.len()];
                    let mut seen: HashSet<Vec<Value>> = HashSet::new();
                    let mut out_rows: Vec<Vec<Value>> = Vec::new();
                    let base_rows = base_rows_set
                        .iter_ones()
                        .chain(include_virtual_blank_row.then_some(base_table_ref.row_count()));
                    for row in base_rows {
                        for key in collect_keys_for_node(
                            &root,
                            model,
                            &row_sets,
                            base_table_ref,
                            std::slice::from_ref(&row),
                            blank_key.clone(),
                        )? {
                            if seen.insert(key.clone()) {
                                out_rows.push(key);
                            }
                        }
                    }

                    Ok(TableResult::Virtual {
                        columns: group_cols,
                        rows: out_rows,
                    })
                }
                "RELATEDTABLE" => {
                    let [table_arg] = args.as_slice() else {
                        return Err(DaxError::Eval("RELATEDTABLE expects 1 argument".into()));
                    };
                    let Expr::TableName(target_table) = table_arg else {
                        return Err(DaxError::Type(
                            "RELATEDTABLE currently expects a table name".into(),
                        ));
                    };
                    let Some(current_table) = row_ctx.current_table() else {
                        return Err(DaxError::Eval("RELATEDTABLE requires row context".into()));
                    };
                    let (current_row, current_visible_cols) = row_ctx
                        .physical_row_for(current_table)
                        .ok_or_else(|| DaxError::Eval("missing current row".into()))?;

                    // Resolve a unique active relationship chain in the reverse direction
                    // (one-to-many at each hop). This also catches ambiguous cases where multiple
                    // relationship paths exist.
                    let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
                    for &idx in filter.relationship_overrides() {
                        if let Some(rel) = model.relationships().get(idx) {
                            override_pairs
                                .insert((rel.rel.from_table.as_str(), rel.rel.to_table.as_str()));
                        }
                    }

                    let is_relationship_active = |idx: usize, rel: &RelationshipInfo| {
                        let pair = (rel.rel.from_table.as_str(), rel.rel.to_table.as_str());
                        let is_active = if override_pairs.contains(&pair) {
                            filter.relationship_overrides().contains(&idx)
                        } else {
                            rel.rel.is_active
                        };

                        is_active && !filter.is_relationship_disabled(idx)
                    };

                    let Some(path) = model.find_unique_active_relationship_path(
                        current_table,
                        target_table,
                        RelationshipPathDirection::OneToMany,
                        |idx, rel| is_relationship_active(idx, rel),
                    )?
                    else {
                        return Err(DaxError::Eval(format!(
                            "no active relationship between {current_table} and {target_table}"
                        )));
                    };

                    // If the current row context is restricted (e.g. iterating `VALUES(Table[Column])`),
                    // ensure `RELATEDTABLE` does not read hidden join key columns from a
                    // representative physical row.
                    if let Some(visible_cols) = current_visible_cols {
                        if let Some(&first_rel_idx) = path.first() {
                            let rel_info = model
                                .relationships()
                                .get(first_rel_idx)
                                .expect("relationship index from path");
                            if rel_info.to_table_key == current_table
                                && !visible_cols.contains(&rel_info.to_idx)
                            {
                                return Err(DaxError::Eval(format!(
                                    "column {current_table}[{}] is not available in the current row context",
                                    rel_info.rel.to_column
                                )));
                            }
                        }
                    }

                    // Fast path: direct relationship `target_table (many) -> current_table (one)`.
                    if path.len() == 1 {
                        let rel = model
                            .relationships()
                            .get(path[0])
                            .expect("relationship index from path");
                        let to_table_ref = model
                            .table(current_table)
                            .ok_or_else(|| DaxError::UnknownTable(current_table.to_string()))?;
                        let to_row_count = to_table_ref.row_count();
                        let is_virtual_blank_row = current_row >= to_row_count;
                        let key = if is_virtual_blank_row {
                            Value::Blank
                        } else {
                            to_table_ref
                                .value_by_idx(current_row, rel.to_idx)
                                .unwrap_or(Value::Blank)
                        };

                        let sets = resolve_row_sets(model, filter)?;
                        let allowed = sets
                            .get(&normalize_ident(target_table))
                            .ok_or_else(|| DaxError::UnknownTable(target_table.to_string()))?;

                        let mut rows = Vec::new();
                        if is_virtual_blank_row {
                            // The relationship-generated blank/unknown member is distinct from a
                            // *physical* BLANK key on the dimension side. Only the virtual blank
                            // member should expose unmatched fact rows.
                            if let Some(unmatched) = rel.unmatched_fact_rows.as_ref() {
                                unmatched.for_each_row(|row| {
                                    if row < allowed.len() && allowed.get(row) {
                                        rows.push(row);
                                    }
                                });
                            }
                        } else if key.is_blank() {
                            // A physical BLANK key on the dimension side is *not* the
                            // relationship-generated blank member. Rows with BLANK foreign keys
                            // belong to the virtual blank member, and should not match this row.
                        } else if let Some(from_index) = rel.from_index.as_ref() {
                            if let Some(candidates) = from_index.get(&key) {
                                for &row in candidates {
                                    if row < allowed.len() && allowed.get(row) {
                                        rows.push(row);
                                    }
                                }
                            }
                        } else {
                            let from_table_ref = model
                                .table(target_table)
                                .ok_or_else(|| DaxError::UnknownTable(target_table.to_string()))?;
                            if let Some(candidates) = from_table_ref.filter_eq(rel.from_idx, &key) {
                                for row in candidates {
                                    if row < allowed.len() && allowed.get(row) {
                                        rows.push(row);
                                    }
                                }
                            } else {
                                // Fallback: scan allowed rows and compare.
                                for row in allowed.iter_ones() {
                                    let v = from_table_ref
                                        .value_by_idx(row, rel.from_idx)
                                        .unwrap_or(Value::Blank);
                                    if v == key {
                                        rows.push(row);
                                    }
                                }
                            }
                        }

                        return Ok(TableResult::Physical {
                            table: normalize_ident(target_table),
                            rows,
                            visible_cols: None,
                        });
                    }

                    // Pre-compute the current filter row sets once so we can reuse them both for
                    // intermediate blank-row checks and the final intersection with filter
                    // context.
                    let sets = (!filter.is_empty())
                        .then(|| resolve_row_sets(model, filter))
                        .transpose()?;
                    let mut current_rows: Vec<usize> = vec![current_row];
                    for rel_idx in path {
                        let rel_info = model
                            .relationships()
                            .get(rel_idx)
                            .expect("relationship index from path");

                        let to_table_ref = model
                            .table(&rel_info.rel.to_table)
                            .ok_or_else(|| DaxError::UnknownTable(rel_info.rel.to_table.clone()))?;

                        let mut key_set: HashSet<Value> = HashSet::new();
                        let mut keys: Vec<Value> = Vec::new();
                        let mut include_blank = false;
                        let to_row_count = to_table_ref.row_count();
                        for &to_row in &current_rows {
                            if to_row >= to_row_count {
                                include_blank = true;
                                continue;
                            }
                            let key = to_table_ref
                                .value_by_idx(to_row, rel_info.to_idx)
                                .unwrap_or(Value::Blank);
                            if key.is_blank() {
                                // Physical BLANK keys do not participate in relationship joins.
                                // They are distinct from the relationship-generated virtual blank
                                // member (`row == row_count()`), which is handled via
                                // `include_blank`.
                                continue;
                            }
                            if key_set.insert(key.clone()) {
                                keys.push(key);
                            }
                        }

                        let mut next_rows: Vec<usize> = Vec::new();
                        if let Some(from_index) = rel_info.from_index.as_ref() {
                            for key in &keys {
                                if let Some(candidates) = from_index.get(key) {
                                    next_rows.extend(candidates.iter().copied());
                                }
                            }

                            if include_blank {
                                if let Some(unmatched) = rel_info.unmatched_fact_rows.as_ref() {
                                    unmatched.extend_into(&mut next_rows);
                                }

                                // In snowflake schemas, the "blank row" can cascade: an unmatched
                                // key in a lower-level fact table creates a virtual blank row in
                                // an intermediate dimension table. That intermediate blank row
                                // should in turn be considered a member of this relationship's
                                // blank row. Include it as a candidate so subsequent hops can
                                // discover unmatched rows further down the chain.
                                if blank_row_allowed(filter, &rel_info.rel.from_table)
                                    && virtual_blank_row_exists(
                                        model,
                                        filter,
                                        &rel_info.rel.from_table,
                                        sets.as_ref(),
                                    )?
                                {
                                    let from_table_ref =
                                        model.table(&rel_info.rel.from_table).ok_or_else(|| {
                                            DaxError::UnknownTable(rel_info.rel.from_table.clone())
                                        })?;
                                    next_rows.push(from_table_ref.row_count());
                                }
                            }
                        } else {
                            let from_table_ref =
                                model.table(&rel_info.rel.from_table).ok_or_else(|| {
                                    DaxError::UnknownTable(rel_info.rel.from_table.clone())
                                })?;

                            if !keys.is_empty() {
                                if let Some(rows) =
                                    from_table_ref.filter_in(rel_info.from_idx, &keys)
                                {
                                    next_rows.extend(rows);
                                } else {
                                    // Fallback: scan and check membership.
                                    let key_set: HashSet<Value> = keys.iter().cloned().collect();
                                    for row in 0..from_table_ref.row_count() {
                                        let v = from_table_ref
                                            .value_by_idx(row, rel_info.from_idx)
                                            .unwrap_or(Value::Blank);
                                        if key_set.contains(&v) {
                                            next_rows.push(row);
                                        }
                                    }
                                }
                            }

                            if include_blank {
                                if let Some(unmatched) = rel_info.unmatched_fact_rows.as_ref() {
                                    unmatched.extend_into(&mut next_rows);
                                }

                                if blank_row_allowed(filter, &rel_info.rel.from_table)
                                    && virtual_blank_row_exists(
                                        model,
                                        filter,
                                        &rel_info.rel.from_table,
                                        sets.as_ref(),
                                    )?
                                {
                                    next_rows.push(from_table_ref.row_count());
                                }
                            }
                        }

                        if next_rows.is_empty() {
                            current_rows.clear();
                            break;
                        }
                        next_rows.sort_unstable();
                        next_rows.dedup();
                        current_rows = next_rows;
                    }

                    let target_key = normalize_ident(target_table);
                    let rows = if let Some(sets) = sets {
                        let allowed = sets
                            .get(&target_key)
                            .ok_or_else(|| DaxError::UnknownTable(target_table.to_string()))?;
                        current_rows
                            .into_iter()
                            .filter(|row| *row < allowed.len() && allowed.get(*row))
                            .collect()
                    } else {
                        current_rows
                    };

                    Ok(TableResult::Physical {
                        table: target_key,
                        rows,
                        visible_cols: None,
                    })
                }
                other => Err(DaxError::Eval(format!(
                    "unsupported table function {other}"
                ))),
            },
            other => Err(DaxError::Type(format!(
                "expression {other:?} cannot be evaluated as a table"
            ))),
        }
    }
}

#[derive(Clone, Debug)]
enum TableResult {
    Physical {
        table: String,
        rows: Vec<usize>,
        /// Restrict row context visibility and context transition to only these column indices.
        visible_cols: Option<Vec<usize>>,
    },
    /// A physical table expression that includes **all** physical rows of the table.
    ///
    /// This avoids materializing a potentially huge `Vec<usize>` row index list for common table
    /// expressions like `Sales` (unfiltered table references) and `ALL(Sales)` when used as an
    /// iterator input.
    PhysicalAll {
        table: String,
        /// Physical row count (virtual relationship blank rows are *not* included).
        row_count: usize,
        /// Restrict row context visibility and context transition to only these column indices.
        visible_cols: Option<Vec<usize>>,
    },
    /// A physical table expression represented as a bitmap of allowed physical rows.
    ///
    /// This avoids materializing a potentially huge `Vec<usize>` row index list for large filtered
    /// row sets (e.g. iterators like `SUMX(FILTER(...), ...)` over dense filters).
    PhysicalMask {
        table: String,
        mask: Arc<BitVec>,
        /// Restrict row context visibility and context transition to only these column indices.
        visible_cols: Option<Vec<usize>>,
    },
    Virtual {
        /// Columns (with lineage) present in the virtual table, in order.
        columns: Vec<(String, String)>,
        /// Row values aligned with `columns`.
        rows: Vec<Vec<Value>>,
    },
}

#[derive(Clone, Copy, Debug)]
enum RowHandle {
    Physical(usize),
    Virtual(usize),
}

struct TableRowIter<'a> {
    inner: TableRowIterInner<'a>,
}

struct BitVecOnesIter<'a> {
    words: &'a [u64],
    len: usize,
    word_idx: usize,
    current_word: u64,
    base: usize,
}

impl<'a> BitVecOnesIter<'a> {
    fn new(bits: &'a BitVec) -> Self {
        Self {
            words: bits.as_words(),
            len: bits.len(),
            word_idx: 0,
            current_word: 0,
            base: 0,
        }
    }
}

impl<'a> Iterator for BitVecOnesIter<'a> {
    type Item = usize;

    fn next(&mut self) -> Option<Self::Item> {
        loop {
            if self.current_word != 0 {
                let bit = self.current_word.trailing_zeros() as usize;
                // Clear lowest set bit.
                self.current_word &= self.current_word - 1;
                let idx = self.base + bit;
                if idx < self.len {
                    return Some(idx);
                }
                continue;
            }

            if self.word_idx >= self.words.len() {
                return None;
            }

            self.current_word = self.words[self.word_idx];
            self.base = self.word_idx * 64;
            self.word_idx += 1;
        }
    }
}

enum TableRowIterInner<'a> {
    Physical(std::slice::Iter<'a, usize>),
    PhysicalRange(std::ops::Range<usize>),
    PhysicalMask(BitVecOnesIter<'a>),
    Virtual(std::ops::Range<usize>),
}

impl<'a> Iterator for TableRowIter<'a> {
    type Item = RowHandle;

    fn next(&mut self) -> Option<Self::Item> {
        match &mut self.inner {
            TableRowIterInner::Physical(iter) => iter.next().copied().map(RowHandle::Physical),
            TableRowIterInner::PhysicalRange(iter) => iter.next().map(RowHandle::Physical),
            TableRowIterInner::PhysicalMask(iter) => iter.next().map(RowHandle::Physical),
            TableRowIterInner::Virtual(iter) => iter.next().map(RowHandle::Virtual),
        }
    }
}

impl TableResult {
    fn row_count(&self) -> usize {
        match self {
            TableResult::Physical { rows, .. } => rows.len(),
            TableResult::PhysicalAll { row_count, .. } => *row_count,
            TableResult::PhysicalMask { mask, .. } => mask.count_ones(),
            TableResult::Virtual { rows, .. } => rows.len(),
        }
    }

    fn iter_rows(&self) -> TableRowIter<'_> {
        match self {
            TableResult::Physical { rows, .. } => TableRowIter {
                inner: TableRowIterInner::Physical(rows.iter()),
            },
            TableResult::PhysicalAll { row_count, .. } => TableRowIter {
                inner: TableRowIterInner::PhysicalRange(0..*row_count),
            },
            TableResult::PhysicalMask { mask, .. } => TableRowIter {
                inner: TableRowIterInner::PhysicalMask(BitVecOnesIter::new(mask)),
            },
            TableResult::Virtual { rows, .. } => TableRowIter {
                inner: TableRowIterInner::Virtual(0..rows.len()),
            },
        }
    }

    fn push_row_ctx(&self, base: &RowContext, row: RowHandle) -> RowContext {
        let mut out = base.clone();
        match (self, row) {
            (
                TableResult::Physical {
                    table,
                    visible_cols,
                    ..
                },
                RowHandle::Physical(row),
            ) => {
                out.push_physical(table, row, visible_cols.clone());
            }
            (
                TableResult::PhysicalAll {
                    table,
                    visible_cols,
                    ..
                },
                RowHandle::Physical(row),
            ) => {
                out.push_physical(table, row, visible_cols.clone());
            }
            (
                TableResult::PhysicalMask {
                    table,
                    visible_cols,
                    ..
                },
                RowHandle::Physical(row),
            ) => {
                out.push_physical(table, row, visible_cols.clone());
            }
            (TableResult::Virtual { columns, rows }, RowHandle::Virtual(row_idx)) => {
                let values = rows.get(row_idx).cloned().unwrap_or_default();
                let bindings: Vec<((String, String), Value)> = columns
                    .iter()
                    .cloned()
                    .zip(values)
                    .map(|(col, v)| (col, v))
                    .collect();
                out.push_virtual(bindings);
            }
            _ => unreachable!("row handle type does not match table result kind"),
        }
        out
    }
}

#[derive(Clone, Copy, Debug)]
enum IteratorKind {
    Sum,
    Average,
    Count,
    Max,
    Min,
}

#[cfg(test)]
pub(crate) fn resolve_table_rows(
    model: &DataModel,
    filter: &FilterContext,
    table: &str,
) -> DaxResult<Vec<usize>> {
    let table_key = normalize_ident(table);
    if filter.is_empty() {
        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
        return Ok((0..table_ref.row_count()).collect());
    }

    let sets = resolve_row_sets(model, filter)?;
    let Some(rows) = sets.get(&table_key) else {
        return Err(DaxError::UnknownTable(table.to_string()));
    };
    Ok(rows.iter_ones().collect())
}

/// Resolve the current filter context to a per-table set of allowed physical rows.
///
/// The algorithm is:
/// 1. Initialize each table's row set from the explicit column/row filters in [`FilterContext`].
/// 2. Repeatedly apply relationship constraints (`to_table → from_table`, plus the reverse
///    direction when `cross_filter_direction == Both`) until no row set changes.
///
/// The fixed-point iteration is required because relationships can form cycles (for example via
/// bidirectional filtering).
///
/// For [`Cardinality::ManyToMany`], propagation is based on the distinct set of keys visible on the
/// source side (a key is considered visible if *any* row with that key is allowed), rather than a
/// unique lookup row.
pub(crate) fn resolve_row_sets(
    model: &DataModel,
    filter: &FilterContext,
) -> DaxResult<HashMap<String, BitVec>> {
    if filter.is_empty() {
        return Ok(model
            .tables
            .iter()
            .map(|(name, table)| (name.clone(), BitVec::with_len_all_true(table.row_count())))
            .collect());
    }

    let mut sets: HashMap<String, BitVec> = HashMap::new();

    for (name, table) in model.tables.iter() {
        let row_count = table.row_count();
        let mut allowed = BitVec::with_len_all_true(row_count);
        if let Some(row_filter) = filter.row_filters.get(name) {
            match row_filter {
                RowFilter::All => {
                    // Keep `allowed` as all-true: all physical rows are permitted.
                }
                RowFilter::Rows(rows) => {
                    allowed = BitVec::with_len_all_false(row_count);
                    for &row in rows {
                        if row < row_count {
                            allowed.set(row, true);
                        }
                    }
                }
                RowFilter::Mask(mask) => {
                    // `PhysicalMask` row sets are sized to the physical row count, so this should
                    // match `row_count`. If it doesn't (e.g. due to a stale filter context), fall
                    // back to a safe rebuild.
                    if mask.len() == row_count {
                        allowed = mask.as_ref().clone();
                    } else {
                        allowed = BitVec::with_len_all_false(row_count);
                        for row in mask.iter_ones() {
                            if row < row_count {
                                allowed.set(row, true);
                            }
                        }
                    }
                }
            }
        }

        for ((t, c), values) in &filter.column_filters {
            if t != name {
                continue;
            }
            let idx = table.column_idx(c).ok_or_else(|| DaxError::UnknownColumn {
                table: t.clone(),
                column: c.clone(),
            })?;

            if values.is_empty() {
                allowed = BitVec::with_len_all_false(row_count);
                continue;
            }

            // Fast path: equality filter backed by a columnar dictionary scan.
            if values.len() == 1 {
                let value = values.iter().next().expect("len==1");
                if let Some(rows) = table.filter_eq(idx, value) {
                    let mut next = BitVec::with_len_all_false(row_count);
                    for row in rows {
                        if row < row_count && allowed.get(row) {
                            next.set(row, true);
                        }
                    }
                    allowed = next;
                    continue;
                }
            }

            if values.len() > 1 {
                let values_vec: Vec<Value> = values.iter().cloned().collect();
                if let Some(rows) = table.filter_in(idx, &values_vec) {
                    let mut next = BitVec::with_len_all_false(row_count);
                    for row in rows {
                        if row < row_count && allowed.get(row) {
                            next.set(row, true);
                        }
                    }
                    allowed = next;
                    continue;
                }
            }

            // Fallback: scan and check membership.
            let mut next = BitVec::with_len_all_false(row_count);
            for row in allowed.iter_ones() {
                let v = table.value_by_idx(row, idx).unwrap_or(Value::Blank);
                if values.contains(&v) {
                    next.set(row, true);
                }
            }
            allowed = next;

            if allowed.count_ones() == 0 {
                break;
            }
        }

        sets.insert(name.clone(), allowed);
    }

    let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
    for &idx in &filter.active_relationship_overrides {
        if let Some(rel) = model.relationships().get(idx) {
            override_pairs.insert((rel.from_table_key.as_str(), rel.to_table_key.as_str()));
        }
    }

    // `blank_row_allowed` determines whether unmatched keys on the fact side should be treated as
    // belonging to the relationship-generated blank/unknown member on the lookup side. In
    // snowflake schemas, filters on higher-level dimensions can exclude BLANK and should cascade
    // down to intermediate dimensions, making their virtual blank members invisible as well.
    //
    // Example: Categories(filter excludes BLANK) -> Products(blank member filtered out) ->
    // Sales(unmatched ProductId rows excluded).
    let blank_allowed = compute_blank_row_allowed_map(model, filter, &override_pairs);

    let trace_enabled = resolve_row_sets_trace_enabled();
    let mut iterations = 0usize;
    let mut propagate_calls = 0usize;
    let mut propagate_changes = 0usize;

    let mut changed = true;
    while changed {
        if trace_enabled {
            iterations += 1;
        }
        changed = false;
        for (idx, relationship) in model.relationships().iter().enumerate() {
            let pair = (
                relationship.from_table_key.as_str(),
                relationship.to_table_key.as_str(),
            );
            let is_active = if override_pairs.contains(&pair) {
                filter.active_relationship_overrides.contains(&idx)
            } else {
                relationship.rel.is_active
            };

            // CROSSFILTER can disable a relationship for the duration of the evaluation.
            let override_state = filter.cross_filter_overrides.get(&idx).copied();

            if !is_active || matches!(override_state, Some(RelationshipOverride::Disabled)) {
                continue;
            }

            match override_state {
                Some(RelationshipOverride::OneWayReverse) => {
                    if trace_enabled {
                        propagate_calls += 1;
                    }
                    let changed_to_one = propagate_filter(
                        model,
                        &mut sets,
                        relationship,
                        Direction::ToOne,
                        &blank_allowed,
                    )?;
                    if trace_enabled && changed_to_one {
                        propagate_changes += 1;
                    }
                    changed |= changed_to_one;
                }
                Some(RelationshipOverride::Active(dir)) => {
                    if trace_enabled {
                        propagate_calls += 1;
                    }
                    let changed_to_many = propagate_filter(
                        model,
                        &mut sets,
                        relationship,
                        Direction::ToMany,
                        &blank_allowed,
                    )?;
                    if trace_enabled && changed_to_many {
                        propagate_changes += 1;
                    }
                    changed |= changed_to_many;
                    if dir == CrossFilterDirection::Both {
                        if trace_enabled {
                            propagate_calls += 1;
                        }
                        let changed_to_one = propagate_filter(
                            model,
                            &mut sets,
                            relationship,
                            Direction::ToOne,
                            &blank_allowed,
                        )?;
                        if trace_enabled && changed_to_one {
                            propagate_changes += 1;
                        }
                        changed |= changed_to_one;
                    }
                }
                Some(RelationshipOverride::Disabled) => unreachable!("checked above"),
                None => {
                    if trace_enabled {
                        propagate_calls += 1;
                    }
                    let changed_to_many = propagate_filter(
                        model,
                        &mut sets,
                        relationship,
                        Direction::ToMany,
                        &blank_allowed,
                    )?;
                    if trace_enabled && changed_to_many {
                        propagate_changes += 1;
                    }
                    changed |= changed_to_many;
                    if relationship.rel.cross_filter_direction == CrossFilterDirection::Both {
                        if trace_enabled {
                            propagate_calls += 1;
                        }
                        let changed_to_one = propagate_filter(
                            model,
                            &mut sets,
                            relationship,
                            Direction::ToOne,
                            &blank_allowed,
                        )?;
                        if trace_enabled && changed_to_one {
                            propagate_changes += 1;
                        }
                        changed |= changed_to_one;
                    }
                }
            }
        }
    }

    if trace_enabled {
        maybe_trace_resolve_row_sets(
            model,
            filter,
            &sets,
            iterations,
            propagate_calls,
            propagate_changes,
        );
    }

    Ok(sets)
}

fn resolve_row_sets_trace_enabled() -> bool {
    static ENABLED: OnceLock<bool> = OnceLock::new();
    *ENABLED.get_or_init(|| std::env::var_os("FORMULA_DAX_RELATIONSHIP_TRACE").is_some())
}

fn maybe_trace_resolve_row_sets(
    model: &DataModel,
    filter: &FilterContext,
    sets: &HashMap<String, BitVec>,
    iterations: usize,
    propagate_calls: usize,
    propagate_changes: usize,
) {
    static EMITTED: AtomicBool = AtomicBool::new(false);
    if EMITTED.swap(true, AtomicOrdering::Relaxed) {
        return;
    }

    let mut table_counts: Vec<(&str, usize, usize)> = sets
        .iter()
        .map(|(name, allowed)| (name.as_str(), allowed.count_ones(), allowed.len()))
        .collect();
    table_counts.sort_by_key(|(name, _, _)| *name);
    let table_counts = table_counts
        .into_iter()
        .map(|(name, allowed, total)| format!("{name}={allowed}/{total}"))
        .collect::<Vec<_>>()
        .join(", ");

    eprintln!(
        "formula-dax resolve_row_sets: tables={} relationships={} filters(col={}, row={}) iterations={} propagate_calls={} propagate_changes={} sets=[{}]",
        model.tables.len(),
        model.relationships().len(),
        filter.column_filters.len(),
        filter.row_filters.len(),
        iterations,
        propagate_calls,
        propagate_changes,
        table_counts
    );
}

enum Direction {
    ToMany,
    ToOne,
}

fn propagate_filter(
    model: &DataModel,
    sets: &mut HashMap<String, BitVec>,
    relationship: &RelationshipInfo,
    direction: Direction,
    blank_allowed: &HashMap<String, bool>,
) -> DaxResult<bool> {
    match direction {
        Direction::ToMany => {
            // Propagate allowed keys from `to_table` to `from_table`.
            //
            // This is the default relationship direction in Tabular/PowerPivot. For 1:* and 1:1 it
            // corresponds to one → many propagation; for *:* it still uses key-set propagation
            // based on the distinct set of visible keys on the `to_table` side.
            let to_table_name = relationship.to_table_key.as_str();
            let from_table_name = relationship.from_table_key.as_str();

            let to_set = sets
                .get(to_table_name)
                .ok_or_else(|| DaxError::UnknownTable(to_table_name.to_string()))?;

            let blank_row_allowed = blank_allowed.get(to_table_name).copied().unwrap_or(true);

            // If `to_table` is unfiltered (including the relationship's implicit blank/unknown
            // member), it should not restrict `from_table`.
            if blank_row_allowed && to_set.all_true() {
                return Ok(false);
            }

            // Collect the set of relationship keys that are visible on the `to_table` side under
            // the current filter context.
            //
            // For many-to-many relationships, a key can correspond to multiple `to_table` rows. For
            // in-memory fact tables we may have a fully materialized `to_index` (FK -> row set) and
            // can use it for the visible key set. For large columnar `to_table`s, we may store only
            // a key set (to avoid allocating per-key row vectors); in that case, extract the
            // distinct visible values directly from the backend.
            let mut allowed_keys: Vec<Value> = match &relationship.to_index {
                // If all physical rows on the `to_table` side are visible, we can reuse the
                // relationship's precomputed distinct key set rather than scanning the table.
                ToIndex::RowSets { map, .. } if to_set.all_true() => {
                    map.keys().filter(|key| !key.is_blank()).cloned().collect()
                }
                ToIndex::KeySet { keys, .. } if to_set.all_true() => keys.iter().cloned().collect(),
                _ => match (&relationship.to_index, relationship.from_index.as_ref()) {
                    (ToIndex::RowSets { map, .. }, Some(_)) => map
                        .iter()
                        // Fact rows whose FK is BLANK always belong to the relationship-generated
                        // blank member, even if a physical BLANK key exists on the dimension side.
                        // Therefore, do not treat BLANK as a matchable key during propagation.
                        .filter(|(key, _)| !key.is_blank())
                        .filter_map(|(key, rows)| rows.any_allowed(to_set).then_some(key.clone()))
                        .collect(),
                    _ => {
                        let to_table = model
                            .table(to_table_name)
                            .ok_or_else(|| DaxError::UnknownTable(to_table_name.to_string()))?;
                        let to_idx = relationship.to_idx;
                        let all_visible = to_set.all_true();
                        if all_visible {
                            to_table
                                .distinct_values_filtered(to_idx, None)
                                .unwrap_or_else(|| {
                                    let mut seen = HashSet::new();
                                    let mut out = Vec::new();
                                    for row in 0..to_table.row_count() {
                                        let v = to_table
                                            .value_by_idx(row, to_idx)
                                            .unwrap_or(Value::Blank);
                                        if seen.insert(v.clone()) {
                                            out.push(v);
                                        }
                                    }
                                    out
                                })
                        } else {
                            let physical_rows = to_table.row_count();
                            let mut visible_count = to_set.count_ones();
                            if to_set.len() > physical_rows {
                                // Exclude any "virtual blank member" rows that live beyond the
                                // physical table length.
                                for idx in physical_rows..to_set.len() {
                                    if to_set.get(idx) {
                                        visible_count = visible_count.saturating_sub(1);
                                    }
                                }
                            }

                            if visible_count == 0 {
                                Vec::new()
                            } else if visible_count == physical_rows {
                                // All physical rows are visible (but the relationship's virtual
                                // blank member may be filtered out). Extract keys from the full
                                // physical table without materializing row indices.
                                to_table
                                    .distinct_values_filtered(to_idx, None)
                                    .unwrap_or_else(|| {
                                        let mut seen = HashSet::new();
                                        let mut out = Vec::new();
                                        for row in 0..physical_rows {
                                            let v = to_table
                                                .value_by_idx(row, to_idx)
                                                .unwrap_or(Value::Blank);
                                            if seen.insert(v.clone()) {
                                                out.push(v);
                                            }
                                        }
                                        out
                                    })
                            } else if to_set.len() == physical_rows {
                                // Prefer backends that can compute distinct values directly from a
                                // bitmap mask (avoids allocating a potentially huge `Vec<usize>`).
                                to_table
                                    .distinct_values_filtered_mask(to_idx, Some(to_set))
                                    .unwrap_or_else(|| {
                                        // Fall back to the generic implementation below.
                                        let sparse_to_dense_threshold = physical_rows / 64;
                                        if visible_count > sparse_to_dense_threshold {
                                            let mut seen = HashSet::new();
                                            let mut out = Vec::new();
                                            for row in to_set.iter_ones() {
                                                let v = to_table
                                                    .value_by_idx(row, to_idx)
                                                    .unwrap_or(Value::Blank);
                                                if seen.insert(v.clone()) {
                                                    out.push(v);
                                                }
                                            }
                                            out
                                        } else {
                                            let visible_rows: Vec<usize> =
                                                to_set.iter_ones().collect();
                                            to_table
                                                .distinct_values_filtered(
                                                    to_idx,
                                                    Some(visible_rows.as_slice()),
                                                )
                                                .unwrap_or_else(|| {
                                                    let mut seen = HashSet::new();
                                                    let mut out = Vec::new();
                                                    for &row in &visible_rows {
                                                        let v = to_table
                                                            .value_by_idx(row, to_idx)
                                                            .unwrap_or(Value::Blank);
                                                        if seen.insert(v.clone()) {
                                                            out.push(v);
                                                        }
                                                    }
                                                    out
                                                })
                                        }
                                    })
                            } else {
                                // `distinct_values_filtered` takes a row index list. For large row
                                // sets, that list can be prohibitively expensive (8 bytes/row).
                                // Prefer scanning the allowed `BitVec` directly once it would be
                                // cheaper than materializing row indices (same heuristic as
                                // `UnmatchedFactRowsBuilder`).
                                let sparse_to_dense_threshold = physical_rows / 64;
                                if visible_count > sparse_to_dense_threshold {
                                    let mut seen = HashSet::new();
                                    let mut out = Vec::new();
                                    for row in to_set.iter_ones().filter(|&idx| idx < physical_rows)
                                    {
                                        let v = to_table
                                            .value_by_idx(row, to_idx)
                                            .unwrap_or(Value::Blank);
                                        if seen.insert(v.clone()) {
                                            out.push(v);
                                        }
                                    }
                                    out
                                } else {
                                    let visible_rows: Vec<usize> = to_set
                                        .iter_ones()
                                        .filter(|&idx| idx < physical_rows)
                                        .collect();
                                    to_table
                                        .distinct_values_filtered(
                                            to_idx,
                                            Some(visible_rows.as_slice()),
                                        )
                                        .unwrap_or_else(|| {
                                            let mut seen = HashSet::new();
                                            let mut out = Vec::new();
                                            for &row in &visible_rows {
                                                let v = to_table
                                                    .value_by_idx(row, to_idx)
                                                    .unwrap_or(Value::Blank);
                                                if seen.insert(v.clone()) {
                                                    out.push(v);
                                                }
                                            }
                                            out
                                        })
                                }
                            }
                        }
                    }
                },
            };
            allowed_keys.retain(|v| !v.is_blank());
            let from_set = sets
                .get(from_table_name)
                .ok_or_else(|| DaxError::UnknownTable(from_table_name.to_string()))?;
            let mut next = BitVec::with_len_all_false(from_set.len());

            if let Some(from_index) = relationship.from_index.as_ref() {
                // Fast path: in-memory fact tables use a precomputed FK -> row list index.
                for key in &allowed_keys {
                    // Tabular's relationship-generated blank member is distinct from a *physical*
                    // BLANK key on the dimension side. Fact-side BLANK foreign keys should never
                    // match a physical BLANK dimension row; they belong to the virtual blank
                    // member and are only included when `blank_row_allowed` is true (handled
                    // below).
                    if key.is_blank() {
                        continue;
                    }
                    if let Some(rows) = from_index.get(key) {
                        for &row in rows {
                            if row < from_set.len() && from_set.get(row) {
                                next.set(row, true);
                            }
                        }
                    }
                }

                if blank_row_allowed {
                    // Include `from_table` rows whose key is BLANK or does not match any key in
                    // `to_table`. Tabular models treat those rows as belonging to a virtual
                    // blank/unknown member on the `to_table` side.
                    if let Some(unmatched) = relationship.unmatched_fact_rows.as_ref() {
                        unmatched.for_each_row(|row| {
                            if row < from_set.len() && from_set.get(row) {
                                next.set(row, true);
                            }
                        });
                    }
                }
            } else {
                // Columnar fact tables: avoid storing per-key row vectors. Instead, use backend
                // filter primitives.
                let from_table = model
                    .table(from_table_name)
                    .ok_or_else(|| DaxError::UnknownTable(from_table_name.to_string()))?;

                if !allowed_keys.is_empty() {
                    if let Some(rows) = from_table.filter_in(relationship.from_idx, &allowed_keys) {
                        for row in rows {
                            if row < from_set.len() && from_set.get(row) {
                                next.set(row, true);
                            }
                        }
                    } else {
                        // Fallback: scan and check membership.
                        let allowed_set: HashSet<Value> = allowed_keys.iter().cloned().collect();
                        for row in from_set.iter_ones() {
                            let v = from_table
                                .value_by_idx(row, relationship.from_idx)
                                .unwrap_or(Value::Blank);
                            if allowed_set.contains(&v) {
                                next.set(row, true);
                            }
                        }
                    }
                }

                if blank_row_allowed {
                    // Include `from_table` rows whose key is BLANK or does not match any key in
                    // `to_table`. Tabular models treat those rows as belonging to a virtual
                    // blank/unknown member on the `to_table` side.
                    if let Some(unmatched) = relationship.unmatched_fact_rows.as_ref() {
                        unmatched.for_each_row(|row| {
                            if row < from_set.len() && from_set.get(row) {
                                next.set(row, true);
                            }
                        });
                    }
                }
            }

            let changed = bitvec_any_removed(from_set, &next);
            if changed {
                sets.insert(from_table_name.to_string(), next);
            }
            Ok(changed)
        }
        Direction::ToOne => {
            // Propagate allowed keys from `from_table` back to `to_table`.
            //
            // When `cross_filter_direction == Both`, this enables bidirectional filtering for both
            // 1:* / 1:1 and *:* relationships.
            let to_table_name = relationship.to_table_key.as_str();
            let from_table_name = relationship.from_table_key.as_str();

            let from_set = sets
                .get(from_table_name)
                .ok_or_else(|| DaxError::UnknownTable(from_table_name.to_string()))?;
            let to_set = sets
                .get(to_table_name)
                .ok_or_else(|| DaxError::UnknownTable(to_table_name.to_string()))?;

            // If `from_table` isn't filtered, it should not restrict `to_table`. In particular,
            // bidirectional relationships should not remove `to_table` rows that simply have no
            // matching `from_table` rows.
            if from_set.all_true() {
                return Ok(false);
            }

            let mut next = BitVec::with_len_all_false(to_set.len());

            // Derive distinct *non-blank* FK values from the allowed fact rows.
            //
            // Fact-side BLANK foreign keys always belong to the relationship-generated blank member
            // and must not match a physical BLANK key on the dimension side.
            let keys: Vec<Value> = if let Some(from_index) = relationship.from_index.as_ref() {
                let mut out = Vec::new();
                for (key, rows) in from_index {
                    if key.is_blank() {
                        continue;
                    }
                    if rows
                        .iter()
                        .any(|row| *row < from_set.len() && from_set.get(*row))
                    {
                        out.push(key.clone());
                    }
                }
                out
            } else {
                let from_table = model
                    .table(from_table_name)
                    .ok_or_else(|| DaxError::UnknownTable(from_table_name.to_string()))?;

                let visible_count = from_set.count_ones();
                if visible_count == 0 {
                    Vec::new()
                } else if let Some(keys) =
                    from_table.distinct_values_filtered_mask(relationship.from_idx, Some(from_set))
                {
                    keys.into_iter().filter(|k| !k.is_blank()).collect()
                } else {
                    // `distinct_values_filtered` takes a row index list. For large row sets, that
                    // list can be prohibitively expensive (8 bytes/row). Prefer scanning the
                    // allowed `BitVec` directly once it would be cheaper than materializing row
                    // indices (same heuristic as `UnmatchedFactRowsBuilder`).
                    let sparse_to_dense_threshold = from_table.row_count() / 64;
                    if visible_count > sparse_to_dense_threshold {
                        let mut keys: HashSet<Value> = HashSet::new();
                        for row in from_set.iter_ones() {
                            let v = from_table
                                .value_by_idx(row, relationship.from_idx)
                                .unwrap_or(Value::Blank);
                            if !v.is_blank() {
                                keys.insert(v);
                            }
                        }
                        keys.into_iter().collect()
                    } else {
                        let rows: Vec<usize> = from_set.iter_ones().collect();
                        let keys = from_table
                            .distinct_values_filtered(relationship.from_idx, Some(rows.as_slice()))
                            .unwrap_or_else(|| {
                                let mut seen = HashSet::new();
                                let mut out = Vec::new();
                                for &row in &rows {
                                    let v = from_table
                                        .value_by_idx(row, relationship.from_idx)
                                        .unwrap_or(Value::Blank);
                                    if seen.insert(v.clone()) {
                                        out.push(v);
                                    }
                                }
                                out
                            });
                        keys.into_iter().filter(|k| !k.is_blank()).collect()
                    }
                }
            };

            match &relationship.to_index {
                ToIndex::RowSets { map, .. } => {
                    for key in &keys {
                        let Some(to_rows) = map.get(key) else {
                            continue;
                        };
                        to_rows.for_each_row(|to_row| {
                            if to_row < to_set.len() && to_set.get(to_row) {
                                next.set(to_row, true);
                            }
                        });
                    }
                }
                ToIndex::KeySet { .. } => {
                    let to_table = model
                        .table(to_table_name)
                        .ok_or_else(|| DaxError::UnknownTable(to_table_name.to_string()))?;

                    if !keys.is_empty() {
                        if let Some(rows) = to_table.filter_in(relationship.to_idx, &keys) {
                            for to_row in rows {
                                if to_row < to_set.len() && to_set.get(to_row) {
                                    next.set(to_row, true);
                                }
                            }
                        } else {
                            let key_set: HashSet<Value> = keys.iter().cloned().collect();
                            for to_row in to_set.iter_ones() {
                                let v = to_table
                                    .value_by_idx(to_row, relationship.to_idx)
                                    .unwrap_or(Value::Blank);
                                if key_set.contains(&v) {
                                    next.set(to_row, true);
                                }
                            }
                        }
                    }
                }
            }

            let changed = bitvec_any_removed(to_set, &next);
            if changed {
                sets.insert(to_table_name.to_string(), next);
            }
            Ok(changed)
        }
    }
}

fn compute_blank_row_allowed_map(
    model: &DataModel,
    filter: &FilterContext,
    override_pairs: &HashSet<(&str, &str)>,
) -> HashMap<String, bool> {
    // Start with the explicit/lexical filter checks (row filters and column filters on the table
    // itself).
    let mut allowed: HashMap<String, bool> = model
        .tables
        .iter()
        .map(|(name, _)| (name.clone(), blank_row_allowed(filter, name)))
        .collect();

    // Any table that has BLANK explicitly filtered out makes its relationship-generated blank
    // member invisible. That exclusion cascades down snowflake chains because the blank member of
    // a fact/dimension table belongs to the blank member of its lookup tables.
    let mut queue: VecDeque<String> = allowed
        .iter()
        .filter_map(|(table, is_allowed)| (!is_allowed).then_some(table.clone()))
        .collect();

    while let Some(to_table) = queue.pop_front() {
        for (idx, relationship) in model.relationships().iter().enumerate() {
            if relationship.to_table_key != to_table {
                continue;
            }

            let pair = (
                relationship.from_table_key.as_str(),
                relationship.to_table_key.as_str(),
            );
            let is_active = if override_pairs.contains(&pair) {
                filter.active_relationship_overrides.contains(&idx)
            } else {
                relationship.rel.is_active
            };

            if !is_active
                || matches!(
                    filter.cross_filter_overrides.get(&idx).copied(),
                    Some(RelationshipOverride::Disabled)
                )
            {
                continue;
            }

            // If CROSSFILTER forces the relationship to propagate only in the reverse direction,
            // filters on `to_table` do not restrict `from_table`, so BLANK exclusion should not
            // cascade.
            if matches!(
                filter.cross_filter_overrides.get(&idx).copied(),
                Some(RelationshipOverride::OneWayReverse)
            ) {
                continue;
            }

            let from_table = relationship.from_table_key.clone();
            if allowed.get(&from_table).copied().unwrap_or(true) {
                allowed.insert(from_table.clone(), false);
                queue.push_back(from_table);
            }
        }
    }

    allowed
}

fn bitvec_any_removed(prev: &BitVec, next: &BitVec) -> bool {
    if prev.len() != next.len() {
        return true;
    }
    prev.as_words()
        .iter()
        .zip(next.as_words())
        .any(|(p, n)| (p & !n) != 0)
}

fn column_is_dax_numeric(table: &dyn TableBackend, idx: usize) -> Option<bool> {
    use formula_columnar::ColumnType;

    let columnar = table.columnar_table()?;
    let column_type = columnar.schema().get(idx)?.column_type;
    Some(matches!(
        column_type,
        ColumnType::Number
            | ColumnType::DateTime
            | ColumnType::Currency { .. }
            | ColumnType::Percentage { .. }
    ))
}

fn coerce_number(value: &Value) -> DaxResult<f64> {
    match value {
        Value::Number(n) => Ok(n.0),
        Value::Boolean(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Value::Blank => Ok(0.0),
        Value::Text(_) => Err(DaxError::Type(format!("cannot coerce {value} to number"))),
    }
}

fn coerce_text(value: &Value) -> Cow<'_, str> {
    match value {
        Value::Text(s) => Cow::Borrowed(s.as_ref()),
        // DAX has nuanced formatting semantics. For now we use Rust's default formatting.
        Value::Number(n) => Cow::Owned(n.0.to_string()),
        // In DAX, BLANK coerces to the empty string for text operations like concatenation.
        Value::Blank => Cow::Borrowed(""),
        // DAX displays boolean values as TRUE/FALSE.
        Value::Boolean(b) => Cow::Borrowed(if *b { "TRUE" } else { "FALSE" }),
    }
}

fn cmp_text_case_insensitive(a: &str, b: &str) -> Ordering {
    if a.is_ascii() && b.is_ascii() {
        return cmp_ascii_case_insensitive(a, b);
    }

    // Compare using Unicode-aware uppercasing so semantics match Excel-like case-insensitive
    // ordering for non-ASCII text (e.g. ß -> SS).
    let mut a_iter = a.chars().flat_map(|c| c.to_uppercase());
    let mut b_iter = b.chars().flat_map(|c| c.to_uppercase());
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(ac), Some(bc)) => match ac.cmp(&bc) {
                Ordering::Equal => continue,
                ord => return ord,
            },
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}

fn cmp_ascii_case_insensitive(a: &str, b: &str) -> Ordering {
    let mut a_iter = a.as_bytes().iter();
    let mut b_iter = b.as_bytes().iter();
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(&ac), Some(&bc)) => {
                let ac = ac.to_ascii_uppercase();
                let bc = bc.to_ascii_uppercase();
                match ac.cmp(&bc) {
                    Ordering::Equal => continue,
                    ord => return ord,
                }
            }
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}

fn compare_values(op: &BinaryOp, left: &Value, right: &Value) -> DaxResult<bool> {
    let cmp = match (left, right) {
        // Text comparisons (BLANK coerces to empty string).
        (Value::Text(l), Value::Text(r)) => Some(l.as_ref().cmp(r.as_ref())),
        (Value::Text(l), Value::Blank) => Some(l.as_ref().cmp("")),
        (Value::Blank, Value::Text(r)) => Some("".cmp(r.as_ref())),
        (Value::Text(_), _) | (_, Value::Text(_)) => {
            return Err(DaxError::Type(format!(
                "cannot compare {left} and {right} with {op:?}"
            )))
        }
        // Numeric comparisons (BLANK coerces to 0, TRUE/FALSE to 1/0).
        _ => {
            let l = coerce_number(left)?;
            let r = coerce_number(right)?;
            Some(
                l.partial_cmp(&r)
                    .ok_or_else(|| DaxError::Eval("comparison failed".into()))?,
            )
        }
    };

    let cmp = cmp.expect("always set");
    Ok(match op {
        BinaryOp::Equals => cmp == std::cmp::Ordering::Equal,
        BinaryOp::NotEquals => cmp != std::cmp::Ordering::Equal,
        BinaryOp::Less => cmp == std::cmp::Ordering::Less,
        BinaryOp::LessEquals => cmp != std::cmp::Ordering::Greater,
        BinaryOp::Greater => cmp == std::cmp::Ordering::Greater,
        BinaryOp::GreaterEquals => cmp != std::cmp::Ordering::Less,
        _ => unreachable!("unexpected comparison operator {op:?}"),
    })
}

fn distinct_rows_by_all_columns(model: &DataModel, base: &TableResult) -> DaxResult<TableResult> {
    match base {
        TableResult::Physical {
            table,
            rows,
            visible_cols,
        } => {
            let table_ref = model
                .table(table)
                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;

            let mut seen: HashSet<Vec<Value>> = HashSet::new();
            let mut out_rows = Vec::new();
            for &row in rows {
                let indices: Box<dyn Iterator<Item = usize>> = match visible_cols {
                    Some(cols) => Box::new(cols.iter().copied()),
                    None => Box::new(0..table_ref.columns().len()),
                };
                let key: Vec<Value> = indices
                    .map(|idx| table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank))
                    .collect();
                if seen.insert(key) {
                    out_rows.push(row);
                }
            }

            Ok(TableResult::Physical {
                table: table.clone(),
                rows: out_rows,
                visible_cols: visible_cols.clone(),
            })
        }
        TableResult::PhysicalAll {
            table,
            row_count,
            visible_cols,
        } => {
            let table_ref = model
                .table(table)
                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;

            let mut seen: HashSet<Vec<Value>> = HashSet::new();
            let mut out_rows = Vec::new();
            for row in 0..*row_count {
                let indices: Box<dyn Iterator<Item = usize>> = match visible_cols {
                    Some(cols) => Box::new(cols.iter().copied()),
                    None => Box::new(0..table_ref.columns().len()),
                };
                let key: Vec<Value> = indices
                    .map(|idx| table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank))
                    .collect();
                if seen.insert(key) {
                    out_rows.push(row);
                }
            }

            Ok(TableResult::Physical {
                table: table.clone(),
                rows: out_rows,
                visible_cols: visible_cols.clone(),
            })
        }
        TableResult::PhysicalMask {
            table,
            mask,
            visible_cols,
        } => {
            let table_ref = model
                .table(table)
                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;

            let mut seen: HashSet<Vec<Value>> = HashSet::new();
            let mut out_rows = Vec::new();
            for row in mask.iter_ones() {
                let indices: Box<dyn Iterator<Item = usize>> = match visible_cols {
                    Some(cols) => Box::new(cols.iter().copied()),
                    None => Box::new(0..table_ref.columns().len()),
                };
                let key: Vec<Value> = indices
                    .map(|idx| table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank))
                    .collect();
                if seen.insert(key) {
                    out_rows.push(row);
                }
            }

            Ok(TableResult::Physical {
                table: table.clone(),
                rows: out_rows,
                visible_cols: visible_cols.clone(),
            })
        }
        TableResult::Virtual { columns, rows } => {
            let mut seen: HashSet<Vec<Value>> = HashSet::new();
            let mut out_rows = Vec::new();
            for row in rows.iter().cloned() {
                if seen.insert(row.clone()) {
                    out_rows.push(row);
                }
            }
            Ok(TableResult::Virtual {
                columns: columns.clone(),
                rows: out_rows,
            })
        }
    }
}

fn blank_row_allowed(filter: &FilterContext, table: &str) -> bool {
    let table = normalize_ident(table);
    // Row filters represent explicit row sets (e.g. FILTER(table, ...)). Those filters do not
    // include the relationship's implicit blank row, so unmatched foreign keys should be
    // excluded whenever a row filter is present.
    if filter.row_filters.contains_key(&table) {
        return false;
    }

    for ((t, _), values) in &filter.column_filters {
        if t == &table && !values.contains(&Value::Blank) {
            return false;
        }
    }

    true
}

fn virtual_blank_row_exists(
    model: &DataModel,
    filter: &FilterContext,
    table: &str,
    sets: Option<&HashMap<String, BitVec>>,
) -> DaxResult<bool> {
    let table = normalize_ident(table);
    // Tabular models materialize an "unknown" (blank) row on the one-side of relationships when
    // there are fact-side rows whose foreign key is BLANK or has no match in the dimension. We
    // model that row virtually (at `row_count()`), so we need to know whether it exists (and is
    // visible) for a given table under the currently active relationship set (including
    // `USERELATIONSHIP`).

    let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
    for &idx in &filter.active_relationship_overrides {
        if let Some(rel) = model.relationships().get(idx) {
            override_pairs.insert((rel.from_table_key.as_str(), rel.to_table_key.as_str()));
        }
    }

    // If BLANK is filtered out (directly or indirectly through snowflake relationships), the
    // relationship-generated blank member should not be considered visible even if unmatched rows
    // exist further down the chain.
    let blank_allowed = compute_blank_row_allowed_map(model, filter, &override_pairs);
    if !blank_allowed.get(&table).copied().unwrap_or(true) {
        return Ok(false);
    }

    let computed_sets;
    let sets = if filter.is_empty() {
        None
    } else if let Some(sets) = sets {
        Some(sets)
    } else {
        computed_sets = resolve_row_sets(model, filter)?;
        Some(&computed_sets)
    };

    let is_relationship_active = |idx: usize, rel: &RelationshipInfo| {
        let pair = (rel.from_table_key.as_str(), rel.to_table_key.as_str());
        let is_active = if override_pairs.contains(&pair) {
            filter.active_relationship_overrides.contains(&idx)
        } else {
            rel.rel.is_active
        };
        is_active
            && !matches!(
                filter.cross_filter_overrides.get(&idx).copied(),
                Some(RelationshipOverride::Disabled)
            )
    };

    // Compute virtual blank member existence for all tables under the current filter context.
    //
    // This is a fixed-point computation because the blank member can cascade across snowflake
    // relationships:
    //   Sales (unmatched ProductId) -> Products(blank member) -> Categories(blank member)
    //
    // We seed the set with tables that have any *currently visible* unmatched fact rows and then
    // repeatedly propagate blank-member existence from `from_table -> to_table` along active
    // relationships, treating the from-table's virtual blank row as a BLANK foreign key to the
    // related to-table.
    let mut exists: HashMap<String, bool> =
        model.tables.keys().map(|t| (t.clone(), false)).collect();

    // 1) Seed from direct unmatched physical rows.
    for (idx, rel) in model.relationships().iter().enumerate() {
        if !is_relationship_active(idx, rel) {
            continue;
        }

        if !blank_allowed
            .get(rel.to_table_key.as_str())
            .copied()
            .unwrap_or(true)
        {
            continue;
        }

        let has_unmatched = if filter.is_empty() {
            matches!(
                rel.unmatched_fact_rows.as_ref(),
                Some(unmatched) if !unmatched.is_empty()
            )
        } else {
            let sets = sets.expect("row sets are computed when filter is not empty");
            let from_set = sets
                .get(rel.from_table_key.as_str())
                .ok_or_else(|| DaxError::UnknownTable(rel.rel.from_table.clone()))?;
            matches!(
                rel.unmatched_fact_rows.as_ref(),
                Some(unmatched) if unmatched.any_row_allowed(from_set)
            )
        };

        if has_unmatched {
            exists.insert(rel.to_table_key.clone(), true);
        }
    }

    // 2) Propagate across snowflake hops.
    let mut changed = true;
    while changed {
        changed = false;
        for (idx, rel) in model.relationships().iter().enumerate() {
            if !is_relationship_active(idx, rel) {
                continue;
            }

            if !blank_allowed
                .get(rel.to_table_key.as_str())
                .copied()
                .unwrap_or(true)
            {
                continue;
            }

            if !blank_allowed
                .get(rel.from_table_key.as_str())
                .copied()
                .unwrap_or(true)
            {
                continue;
            }

            if exists
                .get(rel.from_table_key.as_str())
                .copied()
                .unwrap_or(false)
                && !exists
                    .get(rel.to_table_key.as_str())
                    .copied()
                    .unwrap_or(false)
            {
                exists.insert(rel.to_table_key.clone(), true);
                changed = true;
            }
        }
    }

    Ok(exists.get(table.as_str()).copied().unwrap_or(false))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{DataModel, Table};

    #[test]
    fn resolve_table_rows_multi_column_filters() {
        let mut model = DataModel::new();
        let mut t = Table::new("T", vec!["A", "B"]);
        t.push_row(vec![1.into(), Value::from("x")]).unwrap();
        t.push_row(vec![1.into(), Value::from("y")]).unwrap();
        t.push_row(vec![2.into(), Value::from("x")]).unwrap();
        model.add_table(t).unwrap();

        let filter = FilterContext::empty()
            .with_column_equals("T", "A", 1.into())
            .with_column_equals("T", "B", Value::from("x"));
        let rows = resolve_table_rows(&model, &filter, "T").unwrap();
        assert_eq!(rows, vec![0]);
    }

    #[test]
    fn resolve_table_rows_finds_table_under_single_column_filter() {
        // Regression guard: `resolve_table_rows` should always return a row set for every table in
        // the model (even under filters), keyed by normalized table name.
        let mut model = DataModel::new();
        let mut fact = Table::new("Fact", vec!["Group", "Amount"]);
        fact.push_row(vec![Value::from("A"), 12.34.into()]).unwrap();
        fact.push_row(vec![Value::from("B"), 2.0.into()]).unwrap();
        model.add_table(fact).unwrap();

        let filter = FilterContext::empty().with_column_equals("Fact", "Amount", 12.34.into());
        let rows = resolve_table_rows(&model, &filter, "Fact").unwrap();
        assert_eq!(rows, vec![0]);
    }
}
