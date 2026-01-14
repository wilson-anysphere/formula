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
    Cardinality, CrossFilterDirection, DataModel, RelationshipInfo, RelationshipPathDirection,
    RowSet,
};
use crate::parser::{BinaryOp, Expr, UnaryOp};
use crate::value::Value;
use ordered_float::OrderedFloat;
use std::borrow::Cow;
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

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
    Disabled,
}

#[derive(Clone, Debug, Default)]
pub struct FilterContext {
    column_filters: HashMap<(String, String), HashSet<Value>>,
    row_filters: HashMap<String, HashSet<usize>>,
    active_relationship_overrides: HashSet<usize>,
    cross_filter_overrides: HashMap<usize, RelationshipOverride>,
    suppress_implicit_measure_context_transition: bool,
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
        self.column_filters.insert(
            (table.to_string(), column.to_string()),
            HashSet::from([value]),
        );
    }

    pub fn set_column_in(
        &mut self,
        table: &str,
        column: &str,
        values: impl IntoIterator<Item = Value>,
    ) {
        self.column_filters.insert(
            (table.to_string(), column.to_string()),
            values.into_iter().collect(),
        );
    }

    pub fn clear_column_filter_public(&mut self, table: &str, column: &str) {
        self.clear_column_filter(table, column);
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
        self.column_filters.retain(|(t, _), _| t.as_str() != table);
        self.row_filters.remove(table);
    }

    fn clear_column_filter(&mut self, table: &str, column: &str) {
        self.column_filters
            .remove(&(table.to_string(), column.to_string()));
    }

    fn set_row_filter(&mut self, table: &str, rows: HashSet<usize>) {
        self.row_filters.insert(table.to_string(), rows);
    }
}

#[derive(Clone, Debug, Default)]
pub struct RowContext {
    stack: Vec<RowFrame>,
}

impl RowContext {
    /// Push a full row context frame for `table`/`row` (all columns visible).
    pub fn push(&mut self, table: &str, row: usize) {
        self.push_with_visible_cols(table, row, None);
    }

    pub(crate) fn push_with_visible_cols(
        &mut self,
        table: &str,
        row: usize,
        visible_cols: Option<Vec<usize>>,
    ) {
        self.stack.push(RowFrame {
            table: table.to_string(),
            row,
            visible_cols,
        });
    }

    pub fn pop(&mut self) {
        self.stack.pop();
    }

    /// Update the row index for the innermost (top-of-stack) row context.
    ///
    /// This is useful in hot loops (e.g. calculated-column evaluation) where we want to reuse a
    /// single [`RowContext`] and avoid allocating a new table name string for each row.
    pub fn set_current_row(&mut self, row: usize) {
        if let Some(frame) = self.stack.last_mut() {
            frame.row = row;
        }
    }

    fn current_frame(&self) -> Option<&RowFrame> {
        self.stack.last()
    }

    fn current_table(&self) -> Option<&str> {
        self.current_frame().map(|f| f.table.as_str())
    }

    fn frame_for_level(&self, table: &str, level_from_inner: usize) -> Option<&RowFrame> {
        self.stack
            .iter()
            .rev()
            .filter(|frame| frame.table == table)
            .nth(level_from_inner)
    }

    fn frame_for(&self, table: &str) -> Option<&RowFrame> {
        self.frame_for_level(table, 0)
    }

    fn frame_for_outermost(&self, table: &str) -> Option<&RowFrame> {
        self.stack.iter().find(|frame| frame.table == table)
    }

    fn row_for_level(&self, table: &str, level_from_inner: usize) -> Option<usize> {
        self.frame_for_level(table, level_from_inner).map(|f| f.row)
    }

    fn row_for(&self, table: &str) -> Option<usize> {
        self.row_for_level(table, 0)
    }

    fn row_for_outermost(&self, table: &str) -> Option<usize> {
        self.frame_for_outermost(table).map(|f| f.row)
    }

    fn is_column_visible(&self, table: &str, col_idx: usize) -> bool {
        match self
            .frame_for(table)
            .and_then(|f| f.visible_cols.as_deref())
        {
            None => true,
            Some(cols) => cols.contains(&col_idx),
        }
    }

    fn tables_with_current_rows(&self) -> impl Iterator<Item = (&str, usize, Option<&[usize]>)> {
        let mut seen = HashSet::new();
        self.stack.iter().rev().filter_map(move |frame| {
            if seen.insert(frame.table.as_str()) {
                Some((
                    frame.table.as_str(),
                    frame.row,
                    frame.visible_cols.as_deref(),
                ))
            } else {
                None
            }
        })
    }
}

#[derive(Clone, Debug)]
struct RowFrame {
    table: String,
    row: usize,
    /// When set, this row context frame only exposes the listed column indices.
    ///
    /// This is used to model iterators over virtual one-column (or grouped) tables like
    /// `VALUES(Table[Column])`, where row context and implicit context transition should only
    /// consider the columns present in the virtual table.
    visible_cols: Option<Vec<usize>>,
}

#[derive(Clone, Debug)]
enum VarValue {
    Scalar(Value),
    Table(TableResult),
}

#[derive(Clone, Debug, Default)]
struct VarEnv {
    scopes: Vec<HashMap<String, VarValue>>,
}

impl VarEnv {
    fn normalize_name(name: &str) -> String {
        name.trim().to_ascii_uppercase()
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
            Expr::Measure(name) => {
                let normalized = DataModel::normalize_measure_name(name).to_string();
                if let Some(measure) = model.measures().get(&normalized) {
                    // In DAX, evaluating a measure inside a row context implicitly performs a
                    // context transition (equivalent to `CALCULATE([Measure])`).
                    let eval_filter = if row_ctx.current_table().is_some()
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
                let Some(current_table) = row_ctx.current_table() else {
                    return Err(DaxError::UnknownMeasure(name.clone()));
                };
                let row = row_ctx
                    .row_for(current_table)
                    .ok_or_else(|| DaxError::Eval(format!("no row context for [{normalized}]")))?;
                let table_ref = model
                    .table(current_table)
                    .ok_or_else(|| DaxError::UnknownTable(current_table.to_string()))?;
                let Some(col_idx) = table_ref.column_idx(&normalized) else {
                    return Err(DaxError::Eval(format!(
                        "unknown measure [{normalized}] and no column {current_table}[{normalized}]"
                    )));
                };
                if !row_ctx.is_column_visible(current_table, col_idx) {
                    return Err(DaxError::Eval(format!(
                        "column {current_table}[{normalized}] is not available in the current row context"
                    )));
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
                let row = row_ctx.row_for(table).ok_or_else(|| {
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
                if !row_ctx.is_column_visible(table, idx) {
                    return Err(DaxError::Eval(format!(
                        "column {table}[{column}] is not available in the current row context"
                    )));
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
                let lhs = self.eval_scalar(model, left, filter, row_ctx, env)?;
                let rhs_values =
                    self.eval_one_column_table_literal(model, right, filter, row_ctx, env)?;
                for candidate in rhs_values {
                    if compare_values(&BinaryOp::Equals, &lhs, &candidate)? {
                        return Ok(Value::Boolean(true));
                    }
                }
                Ok(Value::Boolean(false))
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
                Some(VarValue::Table(_)) => Err(DaxError::Type(format!(
                    "table variable {name} used in scalar context"
                ))),
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
        let Expr::TableLiteral { rows } = expr else {
            return Err(DaxError::Type(format!(
                "expected a one-column table constructor, got {expr:?}"
            )));
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
                Ok(Value::from(table_result.rows.len() as i64))
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

                    if search_table != result_table {
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
                let candidate_rows = resolve_table_rows(model, filter, result_table)?;
                let mut matched_rows = Vec::new();
                for row in candidate_rows {
                    let mut matches = true;
                    for (col_idx, search_value) in search_cols.iter().zip(search_values.iter()) {
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
                        Vec::with_capacity(table_result.rows.len());
                    let mut saw_text = false;
                    let mut saw_numeric = false;

                    for row in table_result.rows.iter().copied() {
                        let inner_ctx = table_result.push_row_ctx(row_ctx, row);
                        let value = self.eval_scalar(model, text_expr, filter, &inner_ctx, env)?;
                        let text = coerce_text(&value).into_owned();

                        let key =
                            self.eval_scalar(model, order_by_expr, filter, &inner_ctx, env)?;
                        match &key {
                            Value::Text(_) => saw_text = true,
                            Value::Number(_) | Value::Boolean(_) => saw_numeric = true,
                            Value::Blank => {}
                        }
                        keyed.push((key, text));
                    }

                    if saw_text && saw_numeric {
                        return Err(DaxError::Type(
                            "CONCATENATEX order_by_expr produced mixed text and numeric values"
                                .into(),
                        ));
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
                    for row in table_result.rows.iter().copied() {
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

                // MVP: only support one-column tables.
                // `CONTAINSROW` is commonly used with one-column tables like:
                //   - table constructors: {1, 2, 3}
                //   - VALUES(Table[Column])
                if value_exprs.len() != 1 {
                    return Err(DaxError::Eval(
                        "CONTAINSROW currently only supports one-column tables".into(),
                    ));
                }
                let needle = self.eval_scalar(model, &value_exprs[0], filter, row_ctx, env)?;

                // Table constructors are not (yet) represented as a `TableResult`, but we can
                // evaluate them directly as a one-column list of values.
                if matches!(table_expr, Expr::TableLiteral { .. }) {
                    let values =
                        self.eval_one_column_table_literal(model, table_expr, filter, row_ctx, env)?;
                    for value in values {
                        if compare_values(&BinaryOp::Equals, &value, &needle)? {
                            return Ok(Value::Boolean(true));
                        }
                    }
                    return Ok(Value::Boolean(false));
                }

                let table_result = self.eval_table(model, table_expr, filter, row_ctx, env)?;
                let table_ref = model
                    .table(&table_result.table)
                    .ok_or_else(|| DaxError::UnknownTable(table_result.table.clone()))?;

                let visible_cols: Vec<usize> = match table_result.visible_cols.as_deref() {
                    Some(cols) => cols.to_vec(),
                    None => (0..table_ref.columns().len()).collect(),
                };

                if visible_cols.len() != value_exprs.len() {
                    return Err(DaxError::Eval(format!(
                        "CONTAINSROW expected {} value arguments, got {}",
                        visible_cols.len(),
                        value_exprs.len()
                    )));
                }
                if visible_cols.len() != 1 {
                    return Err(DaxError::Eval(
                        "CONTAINSROW currently only supports one-column tables".into(),
                    ));
                }

                let col_idx = visible_cols[0];
                for row in table_result.rows.iter().copied() {
                    let value = table_ref.value_by_idx(row, col_idx).unwrap_or(Value::Blank);
                    if compare_values(&BinaryOp::Equals, &value, &needle)? {
                        return Ok(Value::Boolean(true));
                    }
                }
                Ok(Value::Boolean(false))
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

                let Some(row) = row_ctx.row_for_level(table, level_from_inner) else {
                    let available = row_ctx
                        .stack
                        .iter()
                        .filter(|frame| frame.table.as_str() == table.as_str())
                        .count();
                    return Err(DaxError::Eval(format!(
                        "EARLIER refers to an outer row context that does not exist for {table}[{column}] (requested level {level_from_inner}, available {available})"
                    )));
                };

                let table_ref = model
                    .table(table)
                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                if row >= table_ref.row_count() {
                    return Ok(Value::Blank);
                }
                let value =
                    table_ref
                        .value(row, column)
                        .ok_or_else(|| DaxError::UnknownColumn {
                            table: table.clone(),
                            column: column.clone(),
                        })?;
                Ok(value)
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

                let Some(row) = row_ctx.row_for_outermost(table) else {
                    return Err(DaxError::Eval(format!(
                        "EARLIEST requires row context for {table}[{column}]"
                    )));
                };

                let table_ref = model
                    .table(table)
                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                if row >= table_ref.row_count() {
                    return Ok(Value::Blank);
                }
                let value =
                    table_ref
                        .value(row, column)
                        .ok_or_else(|| DaxError::UnknownColumn {
                            table: table.clone(),
                            column: column.clone(),
                        })?;
                Ok(value)
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

        let rows = resolve_table_rows(model, filter, table)?;
        let mut sum = 0.0;
        let mut count = 0usize;
        for row in rows {
            if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                sum += n.0;
                count += 1;
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

        let rows = resolve_table_rows(model, filter, table)?;
        let mut sum = 0.0;
        let mut count = 0usize;
        for row in rows {
            if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                sum += n.0;
                count += 1;
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

        let rows = resolve_table_rows(model, filter, table)?;
        let mut best: Option<f64> = None;
        for row in rows {
            if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                best = Some(best.map_or(n.0, |current| current.max(n.0)));
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

        let rows = resolve_table_rows(model, filter, table)?;
        let mut best: Option<f64> = None;
        for row in rows {
            if let Some(Value::Number(n)) = table_ref.value_by_idx(row, idx) {
                best = Some(best.map_or(n.0, |current| current.min(n.0)));
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
        if filter.is_empty() {
            for row in 0..table_ref.row_count() {
                if matches!(table_ref.value_by_idx(row, idx), Some(Value::Number(_))) {
                    count += 1;
                }
            }
        } else {
            let rows = resolve_table_rows(model, filter, table)?;
            for row in rows {
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
        if filter.is_empty() {
            for row in 0..table_ref.row_count() {
                if !table_ref
                    .value_by_idx(row, idx)
                    .unwrap_or(Value::Blank)
                    .is_blank()
                {
                    count += 1;
                }
            }
        } else {
            let rows = resolve_table_rows(model, filter, table)?;
            for row in rows {
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
        let Some(rows_set) = sets.get(table) else {
            return Err(DaxError::UnknownTable(table.to_string()));
        };

        let include_virtual_blank = blank_row_allowed(filter, table)
            && virtual_blank_row_exists(model, filter, table, Some(&sets))?;

        let mut blanks = 0usize;
        for (row, allowed) in rows_set.iter().enumerate() {
            if !*allowed {
                continue;
            }
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

        for row in table_result.rows.iter().copied() {
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

        let rows = resolve_table_rows(model, filter, table)?;
        if let Some(values) = table_ref.distinct_values_filtered(idx, Some(rows.as_slice())) {
            let count = values.iter().filter(|v| !v.is_blank()).count();
            return Ok(Value::from(count as i64));
        }

        let mut out = HashSet::new();
        for row in rows {
            let value = table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
            if !value.is_blank() {
                out.insert(value);
            }
        }
        Ok(Value::from(out.len() as i64))
    }

    fn distinct_column_values(
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
        let Some(rows_set) = sets.get(table) else {
            return Err(DaxError::UnknownTable(table.to_string()));
        };

        let include_virtual_blank = blank_row_allowed(filter, table)
            && virtual_blank_row_exists(model, filter, table, Some(&sets))?;

        let rows: Vec<usize> = rows_set
            .iter()
            .enumerate()
            .filter_map(|(idx, allowed)| allowed.then_some(idx))
            .collect();

        if let Some(values) = table_ref.distinct_values_filtered(idx, Some(rows.as_slice())) {
            let mut out: HashSet<Value> = values.into_iter().collect();
            if include_virtual_blank {
                out.insert(Value::Blank);
            }
            return Ok(out);
        }

        let mut out = HashSet::new();
        for row in rows {
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
        for (table, row, visible_cols) in row_ctx.tables_with_current_rows() {
            let table_ref = model
                .table(table)
                .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
            let table_name = table.to_string();

            let columns = table_ref.columns();
            if let Some(visible_cols) = visible_cols {
                for &col_idx in visible_cols {
                    let Some(column) = columns.get(col_idx) else {
                        return Err(DaxError::Eval(format!(
                            "row context refers to out-of-bounds column index {col_idx} for table {table}"
                        )));
                    };
                    let value = table_ref.value_by_idx(row, col_idx).unwrap_or(Value::Blank);
                    let key = (table_name.clone(), column.clone());
                    match new_filter.column_filters.get_mut(&key) {
                        Some(existing) => {
                            existing.retain(|v| v == &value);
                        }
                        None => {
                            new_filter
                                .column_filters
                                .insert(key, HashSet::from([value]));
                        }
                    }
                }
            } else {
                for (col_idx, column) in columns.iter().enumerate() {
                    let value = table_ref.value_by_idx(row, col_idx).unwrap_or(Value::Blank);
                    let key = (table_name.clone(), column.clone());
                    match new_filter.column_filters.get_mut(&key) {
                        Some(existing) => {
                            existing.retain(|v| v == &value);
                        }
                        None => {
                            new_filter
                                .column_filters
                                .insert(key, HashSet::from([value]));
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
        for arg in filter_args {
            if let Expr::Call { name, args } = arg {
                if name.eq_ignore_ascii_case("USERELATIONSHIP") {
                    self.apply_userelationship(model, filter, args)?;
                    self.apply_userelationship(model, &mut eval_filter, args)?;
                } else if name.eq_ignore_ascii_case("CROSSFILTER") {
                    self.apply_crossfilter(model, filter, args)?;
                    self.apply_crossfilter(model, &mut eval_filter, args)?;
                }
            }
        }

        let mut clear_tables: HashSet<String> = HashSet::new();
        let mut clear_columns: HashSet<(String, String)> = HashSet::new();
        let mut row_filters: Vec<(String, HashSet<usize>)> = Vec::new();
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
                Expr::ColumnRef { table, column } => {
                    tables.insert(table.clone());
                    columns.insert((table.clone(), column.clone()));
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
            env: &mut VarEnv,
            clear_columns: &mut HashSet<(String, String)>,
            row_filters: &mut Vec<(String, HashSet<usize>)>,
        ) -> DaxResult<()> {
            let mut referenced_tables: HashSet<String> = HashSet::new();
            let mut referenced_columns: HashSet<(String, String)> = HashSet::new();
            collect_column_refs(expr, &mut referenced_tables, &mut referenced_columns);

            let table = if referenced_tables.len() == 1 {
                referenced_tables.into_iter().next().expect("len==1")
            } else {
                let mut tables: Vec<String> = referenced_tables.into_iter().collect();
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

            let candidate_rows = resolve_table_rows(model, &base_filter, &table)?;
            let mut allowed_rows = HashSet::new();
            for row in candidate_rows {
                let mut inner_ctx = row_ctx.clone();
                inner_ctx.push(&table, row);
                let pred = engine.eval_scalar(model, expr, &base_filter, &inner_ctx, env)?;
                if pred.truthy().map_err(|e| DaxError::Type(e.to_string()))? {
                    allowed_rows.insert(row);
                }
            }

            row_filters.push((table, allowed_rows));
            Ok(())
        }

        for arg in filter_args {
            // `KEEPFILTERS` wraps a normal filter argument, but changes its semantics from
            // replacement (clear existing filters on the target table/column) to intersection.
            // We implement that by evaluating the inner argument as usual, but skipping any
            // additions to `clear_tables` / `clear_columns`.
            let (arg, keep_filters) = match arg {
                Expr::Call { name, args } if name.eq_ignore_ascii_case("KEEPFILTERS") => {
                    let [inner] = args.as_slice() else {
                        return Err(DaxError::Eval(
                            "KEEPFILTERS expects exactly 1 argument".into(),
                        ));
                    };
                    (inner, true)
                }
                _ => (arg, false),
            };

            match arg {
                Expr::Call { name, .. } if name.eq_ignore_ascii_case("USERELATIONSHIP") => {}
                Expr::Call { name, .. } if name.eq_ignore_ascii_case("CROSSFILTER") => {}
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
                        }
                        Expr::ColumnRef { table, column } => {
                            if !keep_filters {
                                clear_columns.insert((table.clone(), column.clone()));
                            }
                        }
                        other => {
                            return Err(DaxError::Type(format!(
                                "{function_name} expects a table name or column reference, got {other:?}"
                            )))
                        }
                    }
                }
                Expr::Call { name, args } if name.eq_ignore_ascii_case("ALLNOBLANKROW") => {
                    let [inner] = args.as_slice() else {
                        return Err(DaxError::Eval("ALLNOBLANKROW expects 1 argument".into()));
                    };
                    match inner {
                        Expr::TableName(table) => {
                            if !keep_filters {
                                clear_tables.insert(table.clone());
                            }
                            let table_ref = model
                                .table(table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                            // Apply an explicit row filter containing all physical rows. This
                            // matches `ALL(Table)` while ensuring the relationship-generated blank
                            // member is excluded (`blank_row_allowed` is false when a row filter is
                            // present).
                            row_filters.push((
                                table.clone(),
                                (0..table_ref.row_count()).collect::<HashSet<_>>(),
                            ));
                        }
                        Expr::ColumnRef { table, column } => {
                            let key = (table.clone(), column.clone());
                            if !keep_filters {
                                clear_columns.insert(key.clone());
                            }

                            let mut base_filter = eval_filter.clone();
                            base_filter.clear_column_filter(table, column);
                            let mut values =
                                self.distinct_column_values(model, inner, &base_filter)?;
                            values.retain(|v| !v.is_blank());
                            column_filters.push((key, values));
                        }
                        other => {
                            return Err(DaxError::Type(format!(
                            "ALLNOBLANKROW expects a table name or column reference, got {other:?}"
                        )))
                        }
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
                    self,
                    model,
                    arg,
                    &eval_filter,
                    row_ctx,
                    env,
                    keep_filters,
                    env,
                    &mut clear_columns,
                    &mut row_filters,
                )?,
                Expr::Call { name, .. }
                    if name.eq_ignore_ascii_case("NOT")
                        || name.eq_ignore_ascii_case("AND")
                        || name.eq_ignore_ascii_case("OR") =>
                {
                    apply_boolean_filter_expr(
                        self,
                        model,
                        arg,
                        &eval_filter,
                        row_ctx,
                        env,
                        keep_filters,
                        env,
                        &mut clear_columns,
                        &mut row_filters,
                    )?
                }
                Expr::BinaryOp { op, left, right } => {
                    let Expr::ColumnRef { table, column } = left.as_ref() else {
                        return Err(DaxError::Eval(
                            "CALCULATE filter must be a column comparison".into(),
                        ));
                    };

                    let key = (table.clone(), column.clone());
                    if !keep_filters {
                        clear_columns.insert(key.clone());
                    }

                    match op {
                        BinaryOp::In => {
                            let values = self.eval_one_column_table_literal(
                                model,
                                right,
                                &eval_filter,
                                row_ctx,
                                env,
                            )?;
                            column_filters.push((key, values.into_iter().collect()));
                        }
                        BinaryOp::Equals => {
                            let rhs = self.eval_scalar(model, right, &eval_filter, row_ctx, env)?;
                            column_filters.push((key, HashSet::from([rhs])));
                        }
                        BinaryOp::NotEquals
                        | BinaryOp::Less
                        | BinaryOp::LessEquals
                        | BinaryOp::Greater
                        | BinaryOp::GreaterEquals => {
                            let rhs = self.eval_scalar(model, right, &eval_filter, row_ctx, env)?;
                            let mut base_filter = eval_filter.clone();
                            base_filter.column_filters.remove(&key);
                            let candidate_rows = resolve_table_rows(model, &base_filter, table)?;

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
                            for row in candidate_rows {
                                let lhs = table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                let keep = match self.eval_binary(op, lhs.clone(), rhs.clone())? {
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

                            column_filters.push((key, allowed));
                        }
                        _ => {
                            return Err(DaxError::Eval(format!(
                                "unsupported CALCULATE filter operator {op:?}"
                            )))
                        }
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

                    let key = (target_table.clone(), target_column.clone());
                    if !keep_filters {
                        clear_columns.insert(key.clone());
                    }
                    let values =
                        self.distinct_column_values(model, source_col_expr, &eval_filter)?;
                    column_filters.push((key, values));
                }
                Expr::Call { name, args }
                    if (name.eq_ignore_ascii_case("VALUES")
                        || name.eq_ignore_ascii_case("DISTINCT"))
                        && matches!(args.as_slice(), [Expr::ColumnRef { .. }]) =>
                {
                    let Expr::ColumnRef { table, column } = &args[0] else {
                        unreachable!("checked above");
                    };
                    let key = (table.clone(), column.clone());
                    if !keep_filters {
                        clear_columns.insert(key.clone());
                    }
                    let values = self.distinct_column_values(model, &args[0], &eval_filter)?;
                    column_filters.push((key, values));
                }
                Expr::Call { .. } | Expr::TableName(_) => {
                    let table_filter = self.eval_table(model, arg, &eval_filter, row_ctx, env)?;
                    if !keep_filters {
                        clear_tables.insert(table_filter.table.clone());
                    }
                    row_filters.push((table_filter.table, table_filter.rows.into_iter().collect()));
                }
                other => {
                    return Err(DaxError::Eval(format!(
                        "unsupported CALCULATE filter argument {other:?}"
                    )))
                }
            }
        }

        for table in clear_tables {
            filter.clear_table_filters(&table);
        }
        for (table, column) in clear_columns {
            filter.clear_column_filter(&table, &column);
        }

        for (table, rows) in row_filters {
            match filter.row_filters.get_mut(&table) {
                Some(existing) => existing.retain(|row| rows.contains(row)),
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
            model.find_relationship_index(left_table, left_column, right_table, right_column)
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
    ) -> DaxResult<()> {
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
        let override_dir = match direction.as_str() {
            "BOTH" => RelationshipOverride::Active(CrossFilterDirection::Both),
            // DAX uses `ONEWAY` but we'll accept the more explicit `SINGLE` as well.
            "ONEWAY" | "SINGLE" => RelationshipOverride::Active(CrossFilterDirection::Single),
            "NONE" => RelationshipOverride::Disabled,
            other => {
                return Err(DaxError::Eval(format!(
                    "unsupported CROSSFILTER direction {other}"
                )))
            }
        };

        let Some(rel_idx) =
            model.find_relationship_index(left_table, left_column, right_table, right_column)
        else {
            return Err(DaxError::Eval(format!(
                "no relationship found between {left_table}[{left_column}] and {right_table}[{right_column}]"
            )));
        };

        filter.cross_filter_overrides.insert(rel_idx, override_dir);
        Ok(())
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
        let current_row = row_ctx
            .row_for(current_table)
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
        for rel_idx in path {
            let rel_info = model
                .relationships()
                .get(rel_idx)
                .expect("relationship index from path");

            let from_table = model
                .table(&rel_info.rel.from_table)
                .ok_or_else(|| DaxError::UnknownTable(rel_info.rel.from_table.clone()))?;
            let key = from_table
                .value_by_idx(row, rel_info.from_column_idx)
                .unwrap_or(Value::Blank);
            if key.is_blank() {
                return Ok(Value::Blank);
            }

            let Some(to_row_set) = rel_info.to_index.get(&key) else {
                return Ok(Value::Blank);
            };
            let to_row = match to_row_set {
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
                None => Ok(TableResult {
                    table: name.clone(),
                    rows: resolve_table_rows(model, filter, name)?,
                    visible_cols: None,
                }),
            },
            Expr::Call { name, args } => match name.to_ascii_uppercase().as_str() {
                "FILTER" => {
                    let [table_expr, predicate] = args.as_slice() else {
                        return Err(DaxError::Eval("FILTER expects 2 arguments".into()));
                    };
                    let base = self.eval_table(model, table_expr, filter, row_ctx, env)?;
                    let mut rows = Vec::new();
                    for row in base.rows.iter().copied() {
                        let inner_ctx = base.push_row_ctx(row_ctx, row);
                        let pred = self.eval_scalar(model, predicate, filter, &inner_ctx, env)?;
                        if pred.truthy().map_err(|e| DaxError::Type(e.to_string()))? {
                            rows.push(row);
                        }
                    }
                    Ok(TableResult {
                        table: base.table,
                        rows,
                        visible_cols: base.visible_cols,
                    })
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
                            Ok(TableResult {
                                table: name.clone(),
                                rows: (0..table_ref.row_count()).collect(),
                                visible_cols: None,
                            })
                        }
                        Expr::ColumnRef { table, column } => {
                            // `ALL(Table[Column])` removes filters from the target column but
                            // preserves other filters on the same table.
                            let mut modified_filter = filter.clone();
                            modified_filter.clear_column_filter(table, column);
                            let candidate_rows =
                                resolve_table_rows(model, &modified_filter, table)?;

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
                            for row in candidate_rows {
                                let value =
                                    table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                if seen.insert(value) {
                                    rows.push(row);
                                }
                            }
                            if !seen.contains(&Value::Blank)
                                && blank_row_allowed(&modified_filter, table)
                                && virtual_blank_row_exists(model, &modified_filter, table, None)?
                            {
                                rows.push(table_ref.row_count());
                            }
                            Ok(TableResult {
                                table: table.clone(),
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
                            let table_ref = model
                                .table(name)
                                .ok_or_else(|| DaxError::UnknownTable(name.clone()))?;
                            Ok(TableResult {
                                table: name.clone(),
                                rows: (0..table_ref.row_count()).collect(),
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
                            let rows_in_ctx = resolve_table_rows(model, &modified_filter, table)?;

                            let mut seen = HashSet::new();
                            let mut rows = Vec::new();
                            for row in rows_in_ctx {
                                let value =
                                    table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                if value.is_blank() {
                                    continue;
                                }
                                if seen.insert(value) {
                                    rows.push(row);
                                }
                            }

                            Ok(TableResult {
                                table: table.clone(),
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

                            let (rows_in_ctx, sets) = if filter.is_empty() {
                                ((0..table_ref.row_count()).collect(), None)
                            } else {
                                let sets = resolve_row_sets(model, filter)?;
                                let Some(rows_set) = sets.get(table) else {
                                    return Err(DaxError::UnknownTable(table.to_string()));
                                };
                                let rows: Vec<usize> = rows_set
                                    .iter()
                                    .enumerate()
                                    .filter_map(|(idx, allowed)| allowed.then_some(idx))
                                    .collect();
                                (rows, Some(sets))
                            };

                            let mut seen = HashSet::new();
                            let mut rows = Vec::new();
                            for row in rows_in_ctx {
                                let value =
                                    table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                if seen.insert(value) {
                                    rows.push(row);
                                }
                            }
                            if !seen.contains(&Value::Blank)
                                && blank_row_allowed(filter, table)
                                && virtual_blank_row_exists(model, filter, table, sets.as_ref())?
                            {
                                rows.push(table_ref.row_count());
                            }
                            Ok(TableResult {
                                table: table.clone(),
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

                    let mut keep: HashSet<&str> = HashSet::new();
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
                        if col_table != table {
                            return Err(DaxError::Eval(format!(
                                "ALLEXCEPT column must belong to {table}, got {col_table}[{column}]",
                            )));
                        }
                        keep.insert(column.as_str());
                    }

                    let mut modified_filter = filter.clone();
                    modified_filter.clear_table_filters(table);
                    for ((t, c), values) in &filter.column_filters {
                        if t == table && keep.contains(c.as_str()) {
                            modified_filter
                                .column_filters
                                .insert((t.clone(), c.clone()), values.clone());
                        }
                    }

                    Ok(TableResult {
                        table: table.clone(),
                        rows: resolve_table_rows(model, &modified_filter, table)?,
                        visible_cols: None,
                    })
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
                    let table_ref = model
                        .table(&base.table)
                        .ok_or_else(|| DaxError::UnknownTable(base.table.clone()))?;

                    let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
                    for &idx in &filter.active_relationship_overrides {
                        if let Some(rel) = model.relationships().get(idx) {
                            override_pairs
                                .insert((rel.rel.from_table.as_str(), rel.rel.to_table.as_str()));
                        }
                    }

                    let is_relationship_active =
                        |idx: usize, rel: &RelationshipInfo, overrides: &HashSet<(&str, &str)>| {
                            let pair = (rel.rel.from_table.as_str(), rel.rel.to_table.as_str());
                            let is_active = if overrides.contains(&pair) {
                                filter.active_relationship_overrides.contains(&idx)
                            } else {
                                rel.rel.is_active
                            };
                            is_active && !filter.is_relationship_disabled(idx)
                        };

                    struct RelatedHop {
                        relationship_idx: usize,
                        from_idx: usize,
                    }

                    enum GroupAccessor {
                        BaseColumn(usize),
                        RelatedPath {
                            hops: Vec<RelatedHop>,
                            to_table: String,
                            to_col_idx: usize,
                        },
                    }

                    let mut accessors = Vec::with_capacity(group_exprs.len());
                    for expr in group_exprs {
                        let Expr::ColumnRef { table, column } = expr else {
                            return Err(DaxError::Type(
                                "SUMMARIZE currently only supports grouping by columns".into(),
                            ));
                        };
                        if table != &base.table {
                            let Some(path) = model.find_unique_active_relationship_path(
                                &base.table,
                                table,
                                RelationshipPathDirection::ManyToOne,
                                |idx, rel| is_relationship_active(idx, rel, &override_pairs),
                            )?
                            else {
                                return Err(DaxError::Eval(format!(
                                    "SUMMARIZE grouping column {table}[{column}] is not reachable from {}",
                                    base.table
                                )));
                            };

                            let mut hops = Vec::with_capacity(path.len());
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

                                hops.push(RelatedHop {
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

                            accessors.push(GroupAccessor::RelatedPath {
                                hops,
                                to_table: table.clone(),
                                to_col_idx,
                            });
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

                    let mut seen: HashSet<Vec<Value>> = HashSet::new();
                    let mut rows = Vec::new();

                    // Reuse buffers across base rows to avoid repeated allocations.
                    let mut per_col_values: Vec<Vec<Value>> =
                        (0..accessors.len()).map(|_| Vec::new()).collect();
                    let mut partials: Vec<Vec<Value>> = Vec::new();

                    for row in base.rows {
                        // Collect candidate values per grouping column for this base row.
                        for (out, accessor) in per_col_values.iter_mut().zip(accessors.iter()) {
                            out.clear();
                            match accessor {
                                GroupAccessor::BaseColumn(idx) => {
                                    out.push(
                                        table_ref.value_by_idx(row, *idx).unwrap_or(Value::Blank),
                                    );
                                }
                                GroupAccessor::RelatedPath {
                                    hops,
                                    to_table,
                                    to_col_idx,
                                } => {
                                    // Track all reachable row indices along the relationship path,
                                    // expanding many-to-many hops as needed.
                                    let mut current_rows: Vec<usize> = vec![row];
                                    for hop in hops {
                                        let rel_info = model
                                            .relationships()
                                            .get(hop.relationship_idx)
                                            .expect("valid relationship index");

                                        let from_table_ref = model
                                            .table(&rel_info.rel.from_table)
                                            .ok_or_else(|| {
                                                DaxError::UnknownTable(
                                                    rel_info.rel.from_table.clone(),
                                                )
                                            })?;

                                        let allowed_to = row_sets
                                            .get(rel_info.rel.to_table.as_str())
                                            .ok_or_else(|| {
                                                DaxError::UnknownTable(
                                                    rel_info.rel.to_table.clone(),
                                                )
                                            })?;

                                        let mut next_rows: HashSet<usize> = HashSet::new();
                                        for &current_row in &current_rows {
                                            let fk = from_table_ref
                                                .value_by_idx(current_row, hop.from_idx)
                                                .unwrap_or(Value::Blank);
                                            if fk.is_blank() {
                                                continue;
                                            }
                                            let Some(to_row_set) = rel_info.to_index.get(&fk)
                                            else {
                                                continue;
                                            };
                                            to_row_set.for_each_row(|to_row| {
                                                if allowed_to.get(to_row).copied().unwrap_or(false)
                                                {
                                                    next_rows.insert(to_row);
                                                }
                                            });
                                        }

                                        if next_rows.is_empty() {
                                            current_rows.clear();
                                            break;
                                        }

                                        current_rows = next_rows.into_iter().collect();
                                    }

                                    if current_rows.is_empty() {
                                        out.push(Value::Blank);
                                        continue;
                                    }

                                    let to_table_ref = model
                                        .table(to_table)
                                        .ok_or_else(|| DaxError::UnknownTable(to_table.clone()))?;

                                    let mut unique = HashSet::new();
                                    for &to_row in &current_rows {
                                        unique.insert(
                                            to_table_ref
                                                .value_by_idx(to_row, *to_col_idx)
                                                .unwrap_or(Value::Blank),
                                        );
                                    }

                                    if unique.is_empty() {
                                        out.push(Value::Blank);
                                    } else {
                                        out.extend(unique.into_iter());
                                    }
                                }
                            }
                        }

                        // Expand per-column value sets into full group key tuples.
                        partials.clear();
                        partials.push(Vec::new());
                        for values in &per_col_values {
                            let mut next: Vec<Vec<Value>> = Vec::new();
                            for partial in &partials {
                                for value in values {
                                    let mut key = partial.clone();
                                    key.push(value.clone());
                                    next.push(key);
                                }
                            }
                            partials = next;
                        }

                        for key in partials.drain(..) {
                            if seen.insert(key) {
                                rows.push(row);
                            }
                        }
                    }

                    Ok(TableResult {
                        table: base.table,
                        rows,
                        visible_cols: None,
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
                                group_tables.insert(table.clone());
                                group_cols.push((table.clone(), column.clone()));
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

                    #[derive(Clone)]
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

                    let mut base_rows = resolve_table_rows(model, &summarize_filter, &base_table)?;
                    if group_tables.len() == 1
                        && blank_row_allowed(&summarize_filter, &base_table)
                        && virtual_blank_row_exists(model, &summarize_filter, &base_table, None)?
                    {
                        base_rows.push(base_table_ref.row_count());
                    }

                    let mut seen: HashSet<Vec<Value>> = HashSet::new();
                    let mut rows = Vec::new();
                    for row in base_rows {
                        let mut key = Vec::with_capacity(accessors.len());
                        for accessor in &accessors {
                            match accessor {
                                GroupAccessor::BaseColumn(idx) => key.push(
                                    base_table_ref
                                        .value_by_idx(row, *idx)
                                        .unwrap_or(Value::Blank),
                                ),
                                GroupAccessor::RelatedColumn { hops, to_col_idx } => {
                                    let mut current_row = row;
                                    let mut current_table_ref = base_table_ref;
                                    let mut failed = false;
                                    for hop in hops {
                                        let fk = current_table_ref
                                            .value_by_idx(current_row, hop.from_idx)
                                            .unwrap_or(Value::Blank);
                                        if fk.is_blank() {
                                            failed = true;
                                            break;
                                        }
                                        let rel_info = model
                                            .relationships()
                                            .get(hop.relationship_idx)
                                            .expect("valid relationship idx");
                                        let Some(to_row_set) = rel_info.to_index.get(&fk) else {
                                            failed = true;
                                            break;
                                        };
                                        let to_row = match to_row_set {
                                            RowSet::One(row) => *row,
                                            RowSet::Many(rows) => {
                                                if rows.len() == 1 {
                                                    rows[0]
                                                } else {
                                                    return Err(DaxError::Eval(format!(
                                                        "SUMMARIZECOLUMNS grouping is ambiguous: key {fk} matches multiple rows in {} (relationship {})",
                                                        rel_info.rel.to_table, rel_info.rel.name
                                                    )));
                                                }
                                            }
                                        };
                                        current_row = to_row;
                                        current_table_ref = model
                                            .table(&rel_info.rel.to_table)
                                            .ok_or_else(|| {
                                                DaxError::UnknownTable(
                                                    rel_info.rel.to_table.clone(),
                                                )
                                            })?;
                                    }

                                    let value = if failed {
                                        Value::Blank
                                    } else {
                                        current_table_ref
                                            .value_by_idx(current_row, *to_col_idx)
                                            .unwrap_or(Value::Blank)
                                    };
                                    key.push(value);
                                }
                            }
                        }
                        if seen.insert(key) {
                            rows.push(row);
                        }
                    }

                    Ok(TableResult {
                        table: base_table,
                        rows,
                        visible_cols: None,
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

                    let current_row = row_ctx
                        .row_for(current_table)
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

                    // Fast path: direct relationship `target_table (many) -> current_table (one)`.
                    if path.len() == 1 {
                        let rel = model
                            .relationships()
                            .get(path[0])
                            .expect("relationship index from path");
                        let to_table_ref = model
                            .table(current_table)
                            .ok_or_else(|| DaxError::UnknownTable(current_table.to_string()))?;
                        let key = to_table_ref
                            .value_by_idx(current_row, rel.to_column_idx)
                            .unwrap_or(Value::Blank);

                        let sets = resolve_row_sets(model, filter)?;
                        let allowed = sets
                            .get(target_table)
                            .ok_or_else(|| DaxError::UnknownTable(target_table.to_string()))?;

                        let mut rows = Vec::new();
                        if key.is_blank() {
                            if let Some(unmatched) = rel.unmatched_fact_rows.as_deref() {
                                for &row in unmatched {
                                    if allowed.get(row).copied().unwrap_or(false) {
                                        rows.push(row);
                                    }
                                }
                            } else if let Some(from_index) = rel.from_index.as_ref() {
                                for (fk, candidates) in from_index {
                                    if fk.is_blank() || !rel.to_index.contains_key(fk) {
                                        for &row in candidates {
                                            if allowed.get(row).copied().unwrap_or(false) {
                                                rows.push(row);
                                            }
                                        }
                                    }
                                }
                            } else {
                                // Fallback: scan the fact table to preserve blank-member semantics.
                                let from_table_ref =
                                    model.table(target_table).ok_or_else(|| {
                                        DaxError::UnknownTable(target_table.to_string())
                                    })?;
                                for row in 0..from_table_ref.row_count() {
                                    if !allowed.get(row).copied().unwrap_or(false) {
                                        continue;
                                    }
                                    let v = from_table_ref
                                        .value_by_idx(row, rel.from_column_idx)
                                        .unwrap_or(Value::Blank);
                                    if v.is_blank() || !rel.to_index.contains_key(&v) {
                                        rows.push(row);
                                    }
                                }
                            }
                        } else if let Some(from_index) = rel.from_index.as_ref() {
                            if let Some(candidates) = from_index.get(&key) {
                                for &row in candidates {
                                    if allowed.get(row).copied().unwrap_or(false) {
                                        rows.push(row);
                                    }
                                }
                            }
                        } else {
                            let from_table_ref = model
                                .table(target_table)
                                .ok_or_else(|| DaxError::UnknownTable(target_table.to_string()))?;
                            if let Some(candidates) =
                                from_table_ref.filter_eq(rel.from_column_idx, &key)
                            {
                                for row in candidates {
                                    if allowed.get(row).copied().unwrap_or(false) {
                                        rows.push(row);
                                    }
                                }
                            } else {
                                // Fallback: scan and compare.
                                for row in 0..from_table_ref.row_count() {
                                    if !allowed.get(row).copied().unwrap_or(false) {
                                        continue;
                                    }
                                    let v = from_table_ref
                                        .value_by_idx(row, rel.from_column_idx)
                                        .unwrap_or(Value::Blank);
                                    if v == key {
                                        rows.push(row);
                                    }
                                }
                            }
                        }

                        return Ok(TableResult {
                            table: target_table.clone(),
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
                        for &to_row in &current_rows {
                            let key = to_table_ref
                                .value_by_idx(to_row, rel_info.to_column_idx)
                                .unwrap_or(Value::Blank);
                            if key.is_blank() {
                                include_blank = true;
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
                                for (fk, candidates) in from_index {
                                    if fk.is_blank() || !rel_info.to_index.contains_key(fk) {
                                        next_rows.extend(candidates.iter().copied());
                                    }
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
                                    from_table_ref.filter_in(rel_info.from_column_idx, &keys)
                                {
                                    next_rows.extend(rows);
                                } else {
                                    // Fallback: scan and check membership.
                                    let key_set: HashSet<Value> = keys.iter().cloned().collect();
                                    for row in 0..from_table_ref.row_count() {
                                        let v = from_table_ref
                                            .value_by_idx(row, rel_info.from_column_idx)
                                            .unwrap_or(Value::Blank);
                                        if key_set.contains(&v) {
                                            next_rows.push(row);
                                        }
                                    }
                                }
                            }

                            if include_blank {
                                if let Some(unmatched) = rel_info.unmatched_fact_rows.as_deref() {
                                    next_rows.extend(unmatched.iter().copied());
                                } else {
                                    for row in 0..from_table_ref.row_count() {
                                        let v = from_table_ref
                                            .value_by_idx(row, rel_info.from_column_idx)
                                            .unwrap_or(Value::Blank);
                                        if v.is_blank() || !rel_info.to_index.contains_key(&v) {
                                            next_rows.push(row);
                                        }
                                    }
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

                    let rows = if let Some(sets) = sets {
                        let allowed = sets
                            .get(target_table)
                            .ok_or_else(|| DaxError::UnknownTable(target_table.to_string()))?;
                        current_rows
                            .into_iter()
                            .filter(|row| allowed.get(*row).copied().unwrap_or(false))
                            .collect()
                    } else {
                        current_rows
                    };

                    Ok(TableResult {
                        table: target_table.clone(),
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
struct TableResult {
    table: String,
    rows: Vec<usize>,
    visible_cols: Option<Vec<usize>>,
}

impl TableResult {
    fn push_row_ctx(&self, base: &RowContext, row: usize) -> RowContext {
        let mut ctx = base.clone();
        ctx.push_with_visible_cols(&self.table, row, self.visible_cols.clone());
        ctx
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

pub(crate) fn resolve_table_rows(
    model: &DataModel,
    filter: &FilterContext,
    table: &str,
) -> DaxResult<Vec<usize>> {
    if filter.is_empty() {
        let table_ref = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
        return Ok((0..table_ref.row_count()).collect());
    }

    let sets = resolve_row_sets(model, filter)?;
    let Some(rows) = sets.get(table) else {
        return Err(DaxError::UnknownTable(table.to_string()));
    };
    Ok(rows
        .iter()
        .enumerate()
        .filter_map(|(idx, allowed)| allowed.then_some(idx))
        .collect())
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
fn resolve_row_sets(
    model: &DataModel,
    filter: &FilterContext,
) -> DaxResult<HashMap<String, Vec<bool>>> {
    if filter.is_empty() {
        return Ok(model
            .tables
            .iter()
            .map(|(name, table)| (name.clone(), vec![true; table.row_count()]))
            .collect());
    }

    let mut sets: HashMap<String, Vec<bool>> = HashMap::new();

    for (name, table) in model.tables.iter() {
        let row_count = table.row_count();
        let mut allowed = vec![true; row_count];
        if let Some(row_filter) = filter.row_filters.get(name) {
            allowed.fill(false);
            for &row in row_filter {
                if row < row_count {
                    allowed[row] = true;
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
                allowed.fill(false);
                continue;
            }

            // Fast path: equality filter backed by a columnar dictionary scan.
            if values.len() == 1 {
                let value = values.iter().next().expect("len==1");
                if let Some(rows) = table.filter_eq(idx, value) {
                    let mut next = vec![false; row_count];
                    for row in rows {
                        if row < row_count && allowed[row] {
                            next[row] = true;
                        }
                    }
                    allowed = next;
                    continue;
                }
            }

            if values.len() > 1 {
                let values_vec: Vec<Value> = values.iter().cloned().collect();
                if let Some(rows) = table.filter_in(idx, &values_vec) {
                    let mut next = vec![false; row_count];
                    for row in rows {
                        if row < row_count && allowed[row] {
                            next[row] = true;
                        }
                    }
                    allowed = next;
                    continue;
                }
            }

            // Fallback: scan and check membership.
            for row in 0..row_count {
                if !allowed[row] {
                    continue;
                }
                let v = table.value_by_idx(row, idx).unwrap_or(Value::Blank);
                if !values.contains(&v) {
                    allowed[row] = false;
                }
            }
        }

        sets.insert(name.clone(), allowed);
    }

    let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
    for &idx in &filter.active_relationship_overrides {
        if let Some(rel) = model.relationships().get(idx) {
            override_pairs.insert((rel.rel.from_table.as_str(), rel.rel.to_table.as_str()));
        }
    }

    let mut changed = true;
    while changed {
        changed = false;
        for (idx, relationship) in model.relationships().iter().enumerate() {
            let pair = (
                relationship.rel.from_table.as_str(),
                relationship.rel.to_table.as_str(),
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

            let cross_filter_direction = match override_state {
                Some(RelationshipOverride::Active(dir)) => dir,
                None => relationship.rel.cross_filter_direction,
                Some(RelationshipOverride::Disabled) => unreachable!("checked above"),
            };

            changed |= propagate_filter(model, &mut sets, relationship, Direction::ToMany, filter)?;
            if cross_filter_direction == CrossFilterDirection::Both {
                changed |=
                    propagate_filter(model, &mut sets, relationship, Direction::ToOne, filter)?;
            }
        }
    }

    Ok(sets)
}

enum Direction {
    ToMany,
    ToOne,
}

fn propagate_filter(
    model: &DataModel,
    sets: &mut HashMap<String, Vec<bool>>,
    relationship: &RelationshipInfo,
    direction: Direction,
    filter: &FilterContext,
) -> DaxResult<bool> {
    match direction {
        Direction::ToMany => {
            // Propagate allowed keys from the relationship's `to_table` to the `from_table`.
            //
            // This is the standard Tabular/PowerPivot propagation primitive for both 1:* and 1:1
            // relationships (for 1:1, the `from_table` keys are also unique, but we keep the same
            // index shape and propagation logic).
            let to_table_name = relationship.rel.to_table.as_str();
            let from_table_name = relationship.rel.from_table.as_str();

            let to_set = sets
                .get(to_table_name)
                .ok_or_else(|| DaxError::UnknownTable(to_table_name.to_string()))?;

            let blank_row_allowed = blank_row_allowed(filter, to_table_name);

            // If the one-side table is unfiltered (including the relationship's implicit blank
            // row), it should not restrict the many-side table.
            if blank_row_allowed && to_set.iter().all(|allowed| *allowed) {
                return Ok(false);
            }

            // Collect the set of relationship keys that are visible on the `to_table` side under
            // the current filter context. For many-to-many relationships, a key can correspond to
            // multiple `to_table` rows, hence the use of `RowSet::any_allowed`.
            let allowed_keys: Vec<Value> = relationship
                .to_index
                .iter()
                .filter_map(|(key, rows)| rows.any_allowed(to_set).then_some(key.clone()))
                .collect();
            let from_set = sets
                .get(from_table_name)
                .ok_or_else(|| DaxError::UnknownTable(from_table_name.to_string()))?;
            let mut next = vec![false; from_set.len()];

            if let Some(from_index) = relationship.from_index.as_ref() {
                // Fast path: in-memory fact tables use a precomputed FK -> row list index.
                for key in &allowed_keys {
                    if let Some(rows) = from_index.get(key) {
                        for &row in rows {
                            if from_set.get(row).copied().unwrap_or(false) {
                                next[row] = true;
                            }
                        }
                    }
                }

                if blank_row_allowed {
                    // Include fact rows whose foreign key does not match any key in the dimension.
                    // Tabular models treat those rows as belonging to a virtual blank dimension row.
                    for (key, rows) in from_index {
                        if key.is_blank() || !relationship.to_index.contains_key(key) {
                            for &row in rows {
                                if from_set.get(row).copied().unwrap_or(false) {
                                    next[row] = true;
                                }
                            }
                        }
                    }
                }
            } else {
                // Columnar fact tables: avoid storing per-key row vectors. Instead, use backend
                // filter primitives.
                let from_table = model
                    .table(from_table_name)
                    .ok_or_else(|| DaxError::UnknownTable(from_table_name.to_string()))?;

                if !allowed_keys.is_empty() {
                    if let Some(rows) =
                        from_table.filter_in(relationship.from_column_idx, &allowed_keys)
                    {
                        for row in rows {
                            if from_set.get(row).copied().unwrap_or(false) {
                                next[row] = true;
                            }
                        }
                    } else {
                        // Fallback: scan and check membership.
                        let allowed_set: HashSet<Value> = allowed_keys.iter().cloned().collect();
                        for row in 0..from_set.len() {
                            if !from_set[row] {
                                continue;
                            }
                            let v = from_table
                                .value_by_idx(row, relationship.from_column_idx)
                                .unwrap_or(Value::Blank);
                            if allowed_set.contains(&v) {
                                next[row] = true;
                            }
                        }
                    }
                }

                if blank_row_allowed {
                    if let Some(unmatched) = relationship.unmatched_fact_rows.as_deref() {
                        for &row in unmatched {
                            if from_set.get(row).copied().unwrap_or(false) {
                                next[row] = true;
                            }
                        }
                    } else {
                        // Shouldn't happen for columnar relationships, but keep semantics by
                        // scanning if needed.
                        for row in 0..from_set.len() {
                            if !from_set[row] {
                                continue;
                            }
                            let v = from_table
                                .value_by_idx(row, relationship.from_column_idx)
                                .unwrap_or(Value::Blank);
                            if v.is_blank() || !relationship.to_index.contains_key(&v) {
                                next[row] = true;
                            }
                        }
                    }
                }
            }

            let changed = from_set
                .iter()
                .zip(&next)
                .any(|(prev, next)| *prev && !*next);
            if changed {
                sets.insert(from_table_name.to_string(), next);
            }
            Ok(changed)
        }
        Direction::ToOne => {
            // Propagate allowed keys from the relationship's `from_table` back to the `to_table`.
            //
            // When `cross_filter_direction == Both`, this enables bidirectional filtering for both
            // 1:* and 1:1 relationships.
            let to_table_name = relationship.rel.to_table.as_str();
            let from_table_name = relationship.rel.from_table.as_str();

            let from_set = sets
                .get(from_table_name)
                .ok_or_else(|| DaxError::UnknownTable(from_table_name.to_string()))?;
            let to_set = sets
                .get(to_table_name)
                .ok_or_else(|| DaxError::UnknownTable(to_table_name.to_string()))?;

            // Similarly, if the many-side table isn't filtered, it should not restrict the
            // one-side table. In particular, bidirectional relationships should not remove
            // dimension members that simply have no matching fact rows.
            if from_set.iter().all(|allowed| *allowed) {
                return Ok(false);
            }

            let mut next = vec![false; to_set.len()];

            if let Some(from_index) = relationship.from_index.as_ref() {
                for (key, rows) in from_index {
                    if !rows
                        .iter()
                        .any(|row| from_set.get(*row).copied().unwrap_or(false))
                    {
                        continue;
                    }

                    let Some(to_rows) = relationship.to_index.get(key) else {
                        continue;
                    };
                    to_rows.for_each_row(|to_row| {
                        if to_set.get(to_row).copied().unwrap_or(false) {
                            next[to_row] = true;
                        }
                    });
                }
            } else {
                // Columnar fact tables: derive distinct FK values from the allowed fact rows.
                let from_table = model
                    .table(from_table_name)
                    .ok_or_else(|| DaxError::UnknownTable(from_table_name.to_string()))?;

                let rows: Vec<usize> = from_set
                    .iter()
                    .enumerate()
                    .filter_map(|(idx, allowed)| allowed.then_some(idx))
                    .collect();

                let keys = from_table
                    .distinct_values_filtered(relationship.from_column_idx, Some(rows.as_slice()))
                    .unwrap_or_else(|| {
                        let mut seen = HashSet::new();
                        let mut out = Vec::new();
                        for &row in &rows {
                            let v = from_table
                                .value_by_idx(row, relationship.from_column_idx)
                                .unwrap_or(Value::Blank);
                            if seen.insert(v.clone()) {
                                out.push(v);
                            }
                        }
                        out
                    });

                for key in keys {
                    let Some(to_rows) = relationship.to_index.get(&key) else {
                        continue;
                    };
                    to_rows.for_each_row(|to_row| {
                        if to_set.get(to_row).copied().unwrap_or(false) {
                            next[to_row] = true;
                        }
                    });
                }
            }

            let changed = to_set.iter().zip(&next).any(|(prev, next)| *prev && !*next);
            if changed {
                sets.insert(to_table_name.to_string(), next);
            }
            Ok(changed)
        }
    }
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
    let table_ref = model
        .table(&base.table)
        .ok_or_else(|| DaxError::UnknownTable(base.table.clone()))?;

    let column_indices: Vec<usize> = match base.visible_cols.as_deref() {
        Some(cols) => cols.to_vec(),
        None => (0..table_ref.columns().len()).collect(),
    };

    let mut seen: HashSet<Vec<Value>> = HashSet::new();
    let mut rows = Vec::new();
    for row in base.rows.iter().copied() {
        let key: Vec<Value> = column_indices
            .iter()
            .map(|idx| table_ref.value_by_idx(row, *idx).unwrap_or(Value::Blank))
            .collect();
        if seen.insert(key) {
            rows.push(row);
        }
    }

    Ok(TableResult {
        table: base.table.clone(),
        rows,
        visible_cols: base.visible_cols.clone(),
    })
}

fn blank_row_allowed(filter: &FilterContext, table: &str) -> bool {
    // Row filters represent explicit row sets (e.g. FILTER(table, ...)). Those filters do not
    // include the relationship's implicit blank row, so unmatched foreign keys should be
    // excluded whenever a row filter is present.
    if filter.row_filters.contains_key(table) {
        return false;
    }

    for ((t, _), values) in &filter.column_filters {
        if t == table && !values.contains(&Value::Blank) {
            return false;
        }
    }

    true
}

fn virtual_blank_row_exists(
    model: &DataModel,
    filter: &FilterContext,
    table: &str,
    sets: Option<&HashMap<String, Vec<bool>>>,
) -> DaxResult<bool> {
    // Tabular models materialize an "unknown" (blank) row on the one-side of relationships when
    // there are fact-side rows whose foreign key is BLANK or has no match in the dimension. We
    // model that row virtually (at `row_count()`), so we need to know whether it exists for a
    // given table under the currently active relationship set (including `USERELATIONSHIP`).

    let mut override_pairs: HashSet<(&str, &str)> = HashSet::new();
    for &idx in &filter.active_relationship_overrides {
        if let Some(rel) = model.relationships().get(idx) {
            override_pairs.insert((rel.rel.from_table.as_str(), rel.rel.to_table.as_str()));
        }
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

    for (idx, rel) in model.relationships().iter().enumerate() {
        if rel.rel.to_table != table {
            continue;
        }

        let pair = (rel.rel.from_table.as_str(), rel.rel.to_table.as_str());
        let is_active = if override_pairs.contains(&pair) {
            filter.active_relationship_overrides.contains(&idx)
        } else {
            rel.rel.is_active
        };

        if !is_active
            || matches!(
                filter.cross_filter_overrides.get(&idx).copied(),
                Some(RelationshipOverride::Disabled)
            )
        {
            continue;
        }

        // A virtual blank row exists if the relationship has any *currently visible* fact-side
        // row whose foreign key is BLANK or has no match in the dimension.
        if filter.is_empty() {
            if let Some(unmatched) = rel.unmatched_fact_rows.as_deref() {
                if !unmatched.is_empty() {
                    return Ok(true);
                }
            } else if let Some(from_index) = rel.from_index.as_ref() {
                if from_index
                    .keys()
                    .any(|key| key.is_blank() || !rel.to_index.contains_key(key))
                {
                    return Ok(true);
                }
            }
            continue;
        }

        let Some(sets) = sets else {
            continue;
        };
        let from_set = sets
            .get(rel.rel.from_table.as_str())
            .ok_or_else(|| DaxError::UnknownTable(rel.rel.from_table.clone()))?;

        if let Some(unmatched) = rel.unmatched_fact_rows.as_deref() {
            if unmatched
                .iter()
                .copied()
                .any(|row| from_set.get(row).copied().unwrap_or(false))
            {
                return Ok(true);
            }
        } else if let Some(from_index) = rel.from_index.as_ref() {
            for (key, rows) in from_index {
                if !key.is_blank() && rel.to_index.contains_key(key) {
                    continue;
                }
                if rows
                    .iter()
                    .copied()
                    .any(|row| from_set.get(row).copied().unwrap_or(false))
                {
                    return Ok(true);
                }
            }
        }
    }

    Ok(false)
}
