use crate::backend::TableBackend;
use crate::model::{Cardinality, CrossFilterDirection, DataModel, RelationshipInfo};
use crate::parser::{BinaryOp, Expr, UnaryOp};
use crate::value::Value;
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

    #[error("type error: {0}")]
    Type(String),

    #[error("evaluation error: {0}")]
    Eval(String),
}

#[derive(Clone, Debug, Default)]
pub struct FilterContext {
    column_filters: HashMap<(String, String), HashSet<Value>>,
    row_filters: HashMap<String, HashSet<usize>>,
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

    pub fn set_column_equals(&mut self, table: &str, column: &str, value: Value) {
        self.column_filters.insert(
            (table.to_string(), column.to_string()),
            HashSet::from([value]),
        );
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
    stack: Vec<(String, usize)>,
}

impl RowContext {
    pub fn push(&mut self, table: &str, row: usize) {
        self.stack.push((table.to_string(), row));
    }

    pub fn pop(&mut self) {
        self.stack.pop();
    }

    fn current_table(&self) -> Option<&str> {
        self.stack.last().map(|(t, _)| t.as_str())
    }

    fn row_for(&self, table: &str) -> Option<usize> {
        self.stack
            .iter()
            .rev()
            .find(|(t, _)| t == table)
            .map(|(_, r)| *r)
    }

    fn tables_with_current_rows(&self) -> impl Iterator<Item = (&str, usize)> {
        let mut seen = HashSet::new();
        self.stack.iter().rev().filter_map(move |(t, r)| {
            if seen.insert(t.as_str()) {
                Some((t.as_str(), *r))
            } else {
                None
            }
        })
    }
}

#[derive(Clone, Debug, Default)]
pub struct DaxEngine;

impl DaxEngine {
    pub fn new() -> Self {
        Self
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
        self.eval_scalar(model, expr, filter, row_ctx)
    }

    fn eval_scalar(
        &self,
        model: &DataModel,
        expr: &Expr,
        filter: &FilterContext,
        row_ctx: &RowContext,
    ) -> DaxResult<Value> {
        match expr {
            Expr::Number(n) => Ok(Value::from(*n)),
            Expr::Text(s) => Ok(Value::from(s.clone())),
            Expr::Boolean(b) => Ok(Value::from(*b)),
            Expr::Measure(name) => {
                let normalized = DataModel::normalize_measure_name(name).to_string();
                if let Some(measure) = model.measures().get(&normalized) {
                    return self.eval_scalar(
                        model,
                        &measure.parsed,
                        filter,
                        &RowContext::default(),
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
                let value = table_ref.value(row, &normalized).ok_or_else(|| {
                    DaxError::Eval(format!(
                        "unknown measure [{normalized}] and no column {current_table}[{normalized}]"
                    ))
                })?;
                Ok(value)
            }
            Expr::ColumnRef { table, column } => {
                let row = row_ctx.row_for(table).ok_or_else(|| {
                    DaxError::Eval(format!("no row context for {table}[{column}]"))
                })?;
                let table_ref = model
                    .table(table)
                    .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                let value =
                    table_ref
                        .value(row, column)
                        .ok_or_else(|| DaxError::UnknownColumn {
                            table: table.clone(),
                            column: column.clone(),
                        })?;
                Ok(value)
            }
            Expr::UnaryOp { op, expr } => {
                let value = self.eval_scalar(model, expr, filter, row_ctx)?;
                match op {
                    UnaryOp::Negate => {
                        let n = value.as_f64().unwrap_or_else(|| 0.0);
                        Ok(Value::from(-n))
                    }
                }
            }
            Expr::BinaryOp { op, left, right } => {
                let left = self.eval_scalar(model, left, filter, row_ctx)?;
                let right = self.eval_scalar(model, right, filter, row_ctx)?;
                self.eval_binary(op, left, right)
            }
            Expr::Call { name, args } => self.eval_call_scalar(model, name, args, filter, row_ctx),
            Expr::TableName(name) => Err(DaxError::Type(format!(
                "table {name} used in scalar context"
            ))),
        }
    }

    fn eval_binary(&self, op: &BinaryOp, left: Value, right: Value) -> DaxResult<Value> {
        match op {
            BinaryOp::Add | BinaryOp::Subtract | BinaryOp::Multiply | BinaryOp::Divide => {
                let l = left.as_f64().unwrap_or(0.0);
                let r = right.as_f64().unwrap_or(0.0);
                let out = match op {
                    BinaryOp::Add => l + r,
                    BinaryOp::Subtract => l - r,
                    BinaryOp::Multiply => l * r,
                    BinaryOp::Divide => l / r,
                    _ => unreachable!(),
                };
                Ok(Value::from(out))
            }
            BinaryOp::Equals => Ok(Value::Boolean(left == right)),
            BinaryOp::NotEquals => Ok(Value::Boolean(left != right)),
            BinaryOp::Less | BinaryOp::LessEquals | BinaryOp::Greater | BinaryOp::GreaterEquals => {
                let out = match (left, right) {
                    (Value::Number(l), Value::Number(r)) => match op {
                        BinaryOp::Less => l < r,
                        BinaryOp::LessEquals => l <= r,
                        BinaryOp::Greater => l > r,
                        BinaryOp::GreaterEquals => l >= r,
                        _ => unreachable!(),
                    },
                    (Value::Text(l), Value::Text(r)) => match op {
                        BinaryOp::Less => l < r,
                        BinaryOp::LessEquals => l <= r,
                        BinaryOp::Greater => l > r,
                        BinaryOp::GreaterEquals => l >= r,
                        _ => unreachable!(),
                    },
                    (l, r) => {
                        return Err(DaxError::Type(format!(
                            "cannot compare {l} and {r} with {op:?}"
                        )))
                    }
                };
                Ok(Value::Boolean(out))
            }
            BinaryOp::And | BinaryOp::Or => {
                let l = left.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                let r = right.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                Ok(Value::Boolean(match op {
                    BinaryOp::And => l && r,
                    BinaryOp::Or => l || r,
                    _ => unreachable!(),
                }))
            }
        }
    }

    fn eval_call_scalar(
        &self,
        model: &DataModel,
        name: &str,
        args: &[Expr],
        filter: &FilterContext,
        row_ctx: &RowContext,
    ) -> DaxResult<Value> {
        match name.to_ascii_uppercase().as_str() {
            "TRUE" => Ok(Value::Boolean(true)),
            "FALSE" => Ok(Value::Boolean(false)),
            "BLANK" => Ok(Value::Blank),
            "IF" => {
                if args.len() < 2 || args.len() > 3 {
                    return Err(DaxError::Eval("IF expects 2 or 3 arguments".into()));
                }
                let cond = self.eval_scalar(model, &args[0], filter, row_ctx)?;
                let cond = cond.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                if cond {
                    self.eval_scalar(model, &args[1], filter, row_ctx)
                } else if args.len() == 3 {
                    self.eval_scalar(model, &args[2], filter, row_ctx)
                } else {
                    Ok(Value::Blank)
                }
            }
            "DIVIDE" => {
                if args.len() < 2 || args.len() > 3 {
                    return Err(DaxError::Eval("DIVIDE expects 2 or 3 arguments".into()));
                }
                let numerator = self.eval_scalar(model, &args[0], filter, row_ctx)?;
                let denominator = self.eval_scalar(model, &args[1], filter, row_ctx)?;
                let denominator = denominator.as_f64().unwrap_or(0.0);
                if denominator == 0.0 {
                    if args.len() == 3 {
                        self.eval_scalar(model, &args[2], filter, row_ctx)
                    } else {
                        Ok(Value::Blank)
                    }
                } else {
                    let numerator = numerator.as_f64().unwrap_or(0.0);
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
                    let value = self.eval_scalar(model, arg, filter, row_ctx)?;
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
                let value = self.eval_scalar(model, arg, filter, row_ctx)?;
                let b = value.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                Ok(Value::Boolean(!b))
            }
            "AND" => {
                let [left, right] = args else {
                    return Err(DaxError::Eval("AND expects 2 arguments".into()));
                };
                let left = self.eval_scalar(model, left, filter, row_ctx)?;
                let right = self.eval_scalar(model, right, filter, row_ctx)?;
                let left = left.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                let right = right.truthy().map_err(|e| DaxError::Type(e.to_string()))?;
                Ok(Value::Boolean(left && right))
            }
            "OR" => {
                let [left, right] = args else {
                    return Err(DaxError::Eval("OR expects 2 arguments".into()));
                };
                let left = self.eval_scalar(model, left, filter, row_ctx)?;
                let right = self.eval_scalar(model, right, filter, row_ctx)?;
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
                    self.eval_scalar(model, &args[1], filter, row_ctx)
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
                    IteratorKind::Min,
                )
            }
            "COUNTROWS" => {
                let [table_expr] = args else {
                    return Err(DaxError::Eval("COUNTROWS expects 1 argument".into()));
                };
                let table_result = self.eval_table(model, table_expr, filter, row_ctx)?;
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
                    IteratorKind::Count,
                )
            }
            "CALCULATE" => {
                if args.is_empty() {
                    return Err(DaxError::Eval(
                        "CALCULATE expects at least 1 argument".into(),
                    ));
                }
                self.eval_calculate(model, args, filter, row_ctx)
            }
            "RELATED" => {
                let [arg] = args else {
                    return Err(DaxError::Eval("RELATED expects 1 argument".into()));
                };
                self.eval_related(model, arg, row_ctx)
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

    fn eval_iterator(
        &self,
        model: &DataModel,
        table_expr: &Expr,
        value_expr: &Expr,
        filter: &FilterContext,
        row_ctx: &RowContext,
        kind: IteratorKind,
    ) -> DaxResult<Value> {
        let table_result = self.eval_table(model, table_expr, filter, row_ctx)?;
        let mut sum = 0.0;
        let mut count = 0usize;
        let mut best: Option<f64> = None;

        for row in table_result.rows {
            let mut inner_ctx = row_ctx.clone();
            inner_ctx.push(&table_result.table, row);
            let value = self.eval_scalar(model, value_expr, filter, &inner_ctx)?;
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
                    if table_ref.stats_has_blank(idx).unwrap_or(false) {
                        out += 1;
                    }
                    return Ok(Value::from(out));
                }
            }
        }

        let values = self.distinct_column_values(model, expr, filter)?;
        Ok(Value::from(values.len() as i64))
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
            if let Some(values) = table_ref.dictionary_values(idx) {
                let mut out: HashSet<Value> = values.into_iter().collect();
                if table_ref.stats_has_blank(idx).unwrap_or(false) {
                    out.insert(Value::Blank);
                }
                return Ok(out);
            }
        }

        let rows = resolve_table_rows(model, filter, table)?;
        if let Some(values) = table_ref.distinct_values_filtered(idx, Some(rows.as_slice())) {
            return Ok(values.into_iter().collect());
        }

        let mut out = HashSet::new();
        for row in rows {
            let value = table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
            out.insert(value);
        }
        Ok(out)
    }

    fn eval_calculate(
        &self,
        model: &DataModel,
        args: &[Expr],
        filter: &FilterContext,
        row_ctx: &RowContext,
    ) -> DaxResult<Value> {
        let (expr, filter_args) = args.split_first().expect("checked above");
        let mut new_filter = filter.clone();

        for (table, row) in row_ctx.tables_with_current_rows() {
            let table_ref = model
                .table(table)
                .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
            let table_name = table.to_string();

            for (col_idx, column) in table_ref.columns().iter().enumerate() {
                let value = table_ref
                    .value_by_idx(row, col_idx)
                    .unwrap_or(Value::Blank);
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

        for arg in filter_args {
            match arg {
                Expr::BinaryOp { op, left, right } => {
                    let Expr::ColumnRef { table, column } = left.as_ref() else {
                        return Err(DaxError::Eval(
                            "CALCULATE filter must be a column comparison".into(),
                        ));
                    };

                    let rhs = self.eval_scalar(model, right, &new_filter, row_ctx)?;
                    let key = (table.clone(), column.clone());

                    match op {
                        BinaryOp::Equals => {
                            new_filter.column_filters.insert(key, HashSet::from([rhs]));
                        }
                        BinaryOp::NotEquals
                        | BinaryOp::Less
                        | BinaryOp::LessEquals
                        | BinaryOp::Greater
                        | BinaryOp::GreaterEquals => {
                            let mut base_filter = new_filter.clone();
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
                                let lhs =
                                    table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);

                                let keep = match op {
                                    BinaryOp::NotEquals => lhs != rhs,
                                    BinaryOp::Less
                                    | BinaryOp::LessEquals
                                    | BinaryOp::Greater
                                    | BinaryOp::GreaterEquals => {
                                        let Some(l) = lhs.as_f64() else { continue };
                                        let Some(r) = rhs.as_f64() else { continue };
                                        match op {
                                            BinaryOp::Less => l < r,
                                            BinaryOp::LessEquals => l <= r,
                                            BinaryOp::Greater => l > r,
                                            BinaryOp::GreaterEquals => l >= r,
                                            _ => unreachable!(),
                                        }
                                    }
                                    _ => unreachable!(),
                                };

                                if keep {
                                    allowed.insert(lhs);
                                }
                            }

                            new_filter.column_filters.insert(key, allowed);
                        }
                        _ => {
                            return Err(DaxError::Eval(format!(
                                "unsupported CALCULATE filter operator {op:?}"
                            )))
                        }
                    }
                }
                Expr::Call { name, args } if name.eq_ignore_ascii_case("ALL") => {
                    let [inner] = args.as_slice() else {
                        return Err(DaxError::Eval("ALL expects 1 argument".into()));
                    };
                    match inner {
                        Expr::TableName(table) => {
                            new_filter.clear_table_filters(table);
                        }
                        Expr::ColumnRef { table, column } => {
                            new_filter.clear_column_filter(table, column);
                        }
                        other => {
                            return Err(DaxError::Type(format!(
                                "ALL expects a table name or column reference, got {other:?}"
                            )))
                        }
                    }
                }
                Expr::Call { .. } | Expr::TableName(_) => {
                    let table_filter = self.eval_table(model, arg, &new_filter, row_ctx)?;
                    new_filter.clear_table_filters(&table_filter.table);
                    new_filter.set_row_filter(
                        &table_filter.table,
                        table_filter.rows.into_iter().collect(),
                    );
                }
                other => {
                    return Err(DaxError::Eval(format!(
                        "unsupported CALCULATE filter argument {other:?}"
                    )))
                }
            }
        }

        self.eval_scalar(model, expr, &new_filter, row_ctx)
    }

    fn eval_related(
        &self,
        model: &DataModel,
        arg: &Expr,
        row_ctx: &RowContext,
    ) -> DaxResult<Value> {
        let Expr::ColumnRef { table, column } = arg else {
            return Err(DaxError::Type("RELATED expects a column reference".into()));
        };
        let Some(current_table) = row_ctx.current_table() else {
            return Err(DaxError::Eval("RELATED requires row context".into()));
        };

        let rel_info = model
            .relationships()
            .iter()
            .find(|rel| {
                rel.rel.is_active
                    && rel.rel.from_table == current_table
                    && rel.rel.to_table == *table
            })
            .ok_or_else(|| {
                DaxError::Eval(format!(
                    "no active relationship from {current_table} to {table} for RELATED"
                ))
            })?;

        let current_row = row_ctx
            .row_for(current_table)
            .ok_or_else(|| DaxError::Eval("missing row for current table".into()))?;

        let from_table = model
            .table(current_table)
            .ok_or_else(|| DaxError::UnknownTable(current_table.to_string()))?;
        let from_idx = from_table
            .column_idx(&rel_info.rel.from_column)
            .ok_or_else(|| DaxError::UnknownColumn {
                table: current_table.to_string(),
                column: rel_info.rel.from_column.clone(),
            })?;
        let key = from_table
            .value_by_idx(current_row, from_idx)
            .ok_or_else(|| DaxError::Eval("missing key value".into()))?
            ;
        if key.is_blank() {
            return Ok(Value::Blank);
        }

        let Some(to_row) = rel_info.to_index.get(&key).copied() else {
            return Ok(Value::Blank);
        };

        let to_table = model
            .table(table)
            .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
        let value = to_table
            .value(to_row, column)
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
    ) -> DaxResult<TableResult> {
        match expr {
            Expr::TableName(name) => Ok(TableResult {
                table: name.clone(),
                rows: resolve_table_rows(model, filter, name)?,
            }),
            Expr::Call { name, args } => match name.to_ascii_uppercase().as_str() {
                "FILTER" => {
                    let [table_expr, predicate] = args.as_slice() else {
                        return Err(DaxError::Eval("FILTER expects 2 arguments".into()));
                    };
                    let base = self.eval_table(model, table_expr, filter, row_ctx)?;
                    let mut rows = Vec::new();
                    for row in base.rows.iter().copied() {
                        let mut inner_ctx = row_ctx.clone();
                        inner_ctx.push(&base.table, row);
                        let pred = self.eval_scalar(model, predicate, filter, &inner_ctx)?;
                        if pred.truthy().map_err(|e| DaxError::Type(e.to_string()))? {
                            rows.push(row);
                        }
                    }
                    Ok(TableResult {
                        table: base.table,
                        rows,
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
                            })
                        }
                        Expr::ColumnRef { table, column } => {
                            let table_ref = model
                                .table(table)
                                .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                            let idx = table_ref
                                .column_idx(column)
                                .ok_or_else(|| DaxError::UnknownColumn {
                                    table: table.clone(),
                                    column: column.clone(),
                                })?;

                            let mut seen = HashSet::new();
                            let mut rows = Vec::new();
                            for row in 0..table_ref.row_count() {
                                let value =
                                    table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                                if seen.insert(value) {
                                    rows.push(row);
                                }
                            }
                            Ok(TableResult {
                                table: table.clone(),
                                rows,
                            })
                        }
                        other => Err(DaxError::Type(format!(
                            "ALL expects a table name or column reference, got {other:?}"
                        ))),
                    }
                }
                "VALUES" => {
                    let [arg] = args.as_slice() else {
                        return Err(DaxError::Eval("VALUES expects 1 argument".into()));
                    };
                    let Expr::ColumnRef { table, column } = arg else {
                        return Err(DaxError::Type(
                            "VALUES currently only supports a column reference".into(),
                        ));
                    };

                    let rows_in_ctx = resolve_table_rows(model, filter, table)?;
                    let table_ref = model
                        .table(table)
                        .ok_or_else(|| DaxError::UnknownTable(table.clone()))?;
                    let idx = table_ref
                        .column_idx(column)
                        .ok_or_else(|| DaxError::UnknownColumn {
                            table: table.clone(),
                            column: column.clone(),
                        })?;

                    let mut seen = HashSet::new();
                    let mut rows = Vec::new();
                    for row in rows_in_ctx {
                        let value = table_ref.value_by_idx(row, idx).unwrap_or(Value::Blank);
                        if seen.insert(value) {
                            rows.push(row);
                        }
                    }
                    Ok(TableResult {
                        table: table.clone(),
                        rows,
                    })
                }
                "SUMMARIZE" => {
                    let (table_expr, group_exprs) =
                        args.split_first().ok_or_else(|| DaxError::Eval(
                            "SUMMARIZE expects at least 2 arguments".into(),
                        ))?;
                    if group_exprs.is_empty() {
                        return Err(DaxError::Eval(
                            "SUMMARIZE expects at least 2 arguments".into(),
                        ));
                    }

                    let base = self.eval_table(model, table_expr, filter, row_ctx)?;
                    let table_ref = model
                        .table(&base.table)
                        .ok_or_else(|| DaxError::UnknownTable(base.table.clone()))?;

                    let mut group_idxs = Vec::with_capacity(group_exprs.len());
                    for expr in group_exprs {
                        let Expr::ColumnRef { table, column } = expr else {
                            return Err(DaxError::Type(
                                "SUMMARIZE currently only supports grouping by columns".into(),
                            ));
                        };
                        if table != &base.table {
                            return Err(DaxError::Eval(format!(
                                "SUMMARIZE grouping column must belong to {}, got {table}[{column}]",
                                base.table
                            )));
                        }
                        let idx = table_ref
                            .column_idx(column)
                            .ok_or_else(|| DaxError::UnknownColumn {
                                table: table.clone(),
                                column: column.clone(),
                            })?;
                        group_idxs.push(idx);
                    }

                    let mut seen: HashSet<Vec<Value>> = HashSet::new();
                    let mut rows = Vec::new();
                    for row in base.rows {
                        let key: Vec<Value> = group_idxs
                            .iter()
                            .map(|idx| table_ref.value_by_idx(row, *idx).unwrap_or(Value::Blank))
                            .collect();
                        if seen.insert(key) {
                            rows.push(row);
                        }
                    }

                    Ok(TableResult {
                        table: base.table,
                        rows,
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

                    let rel = model
                        .relationships()
                        .iter()
                        .find(|rel| {
                            rel.rel.is_active
                                && rel.rel.from_table == *target_table
                                && rel.rel.to_table == current_table
                        })
                        .ok_or_else(|| {
                            DaxError::Eval(format!(
                                "no active relationship between {current_table} and {target_table}"
                            ))
                        })?;

                    let current_row = row_ctx
                        .row_for(current_table)
                        .ok_or_else(|| DaxError::Eval("missing current row".into()))?;

                    let to_table_ref = model
                        .table(current_table)
                        .ok_or_else(|| DaxError::UnknownTable(current_table.to_string()))?;
                    let to_idx = to_table_ref.column_idx(&rel.rel.to_column).ok_or_else(|| {
                        DaxError::UnknownColumn {
                            table: current_table.to_string(),
                            column: rel.rel.to_column.clone(),
                        }
                    })?;
                    let key = to_table_ref
                        .value_by_idx(current_row, to_idx)
                        .ok_or_else(|| DaxError::Eval("missing key".into()))?
                        ;

                    if key.is_blank() {
                        return Ok(TableResult {
                            table: target_table.clone(),
                            rows: Vec::new(),
                        });
                    }

                    // `RELATEDTABLE` is frequently used inside iterators. Use the relationship
                    // index to fetch candidate fact rows, and only then apply the existing filter
                    // context (including relationship propagation).
                    let sets = resolve_row_sets(model, filter)?;
                    let allowed = sets.get(target_table).ok_or_else(|| {
                        DaxError::UnknownTable(target_table.to_string())
                    })?;

                    let mut rows = Vec::new();
                    if let Some(candidates) = rel.from_index.get(&key) {
                        for &row in candidates {
                            if allowed.get(row).copied().unwrap_or(false) {
                                rows.push(row);
                            }
                        }
                    }

                    Ok(TableResult {
                        table: target_table.clone(),
                        rows,
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

    let mut changed = true;
    while changed {
        changed = false;
        for relationship in model.relationships() {
            if !relationship.rel.is_active {
                continue;
            }
            changed |= propagate_filter(&mut sets, relationship, Direction::ToMany)?;
            if relationship.rel.cross_filter_direction == CrossFilterDirection::Both {
                changed |= propagate_filter(&mut sets, relationship, Direction::ToOne)?;
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
    sets: &mut HashMap<String, Vec<bool>>,
    relationship: &RelationshipInfo,
    direction: Direction,
) -> DaxResult<bool> {
    match direction {
        Direction::ToMany => {
            // Propagate allowed keys from the "one" side to the "many" side:
            //   to_table (one) -> from_table (many)
            let to_table_name = relationship.rel.to_table.as_str();
            let from_table_name = relationship.rel.from_table.as_str();

            let to_set = sets
                .get(to_table_name)
                .ok_or_else(|| DaxError::UnknownTable(to_table_name.to_string()))?;

            let allowed_keys: Vec<&Value> = relationship
                .to_index
                .iter()
                .filter_map(|(key, &row)| to_set.get(row).copied().unwrap_or(false).then_some(key))
                .collect();

            let from_set = sets
                .get(from_table_name)
                .ok_or_else(|| DaxError::UnknownTable(from_table_name.to_string()))?;
            let mut next = vec![false; from_set.len()];

            for key in allowed_keys {
                if let Some(rows) = relationship.from_index.get(key) {
                    for &row in rows {
                        if from_set.get(row).copied().unwrap_or(false) {
                            next[row] = true;
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
            // Propagate allowed keys from the "many" side to the "one" side:
            //   from_table (many) -> to_table (one)
            let to_table_name = relationship.rel.to_table.as_str();
            let from_table_name = relationship.rel.from_table.as_str();

            let from_set = sets
                .get(from_table_name)
                .ok_or_else(|| DaxError::UnknownTable(from_table_name.to_string()))?;
            let to_set = sets
                .get(to_table_name)
                .ok_or_else(|| DaxError::UnknownTable(to_table_name.to_string()))?;

            let mut next = vec![false; to_set.len()];
            for (key, rows) in &relationship.from_index {
                let mut any_allowed = false;
                for &row in rows {
                    if from_set.get(row).copied().unwrap_or(false) {
                        any_allowed = true;
                        break;
                    }
                }
                if !any_allowed {
                    continue;
                }
                let Some(&to_row) = relationship.to_index.get(key) else {
                    continue;
                };
                if to_set.get(to_row).copied().unwrap_or(false) {
                    next[to_row] = true;
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
