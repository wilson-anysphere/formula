use std::collections::HashMap;

use crate::functions::math::criteria::Criteria;
use crate::functions::{ArgValue, FunctionContext};
use crate::value::{casefold, Array, ErrorKind, Value};

#[derive(Debug, Clone)]
pub struct DatabaseTable {
    pub(crate) array: Array,
    /// Case-folded database field names mapped to their column indices.
    pub(crate) header_map: HashMap<String, usize>,
}

impl DatabaseTable {
    fn cell(&self, row: usize, col: usize) -> &Value {
        // Bounds are validated when constructing the table; callers should only request valid
        // coordinates. Default to blank for defensive robustness.
        self.array.get(row, col).unwrap_or(&Value::Blank)
    }

    fn data_row_count(&self) -> usize {
        self.array.rows.saturating_sub(1)
    }
}

#[derive(Debug, Clone)]
pub struct CriteriaClause {
    pub(crate) conditions: Vec<(usize, Criteria)>,
}

#[derive(Debug, Clone)]
pub struct DatabaseQuery {
    pub(crate) table: DatabaseTable,
    pub(crate) field_index: usize,
    pub(crate) criteria: Vec<CriteriaClause>,
}

impl DatabaseQuery {
    /// Iterates matching database **data** rows (excluding the header row).
    ///
    /// Yields row indices in the underlying table array (so the first record row is `1`).
    pub fn iter_matching_rows(&self) -> impl Iterator<Item = usize> + '_ {
        (1..self.table.array.rows).filter(|&row| row_matches(&self.table, row, &self.criteria))
    }

    pub fn field_value(&self, row: usize) -> &Value {
        self.table.cell(row, self.field_index)
    }

    pub fn record_count(&self) -> usize {
        self.table.data_row_count()
    }
}

pub fn parse_query(
    ctx: &dyn FunctionContext,
    database: ArgValue,
    field: Value,
    criteria: ArgValue,
) -> Result<DatabaseQuery, ErrorKind> {
    let table = parse_database_range(ctx, database)?;
    let field_index = resolve_field(ctx, &table, &field)?;
    let criteria = parse_criteria_range(ctx, &table, criteria)?;
    Ok(DatabaseQuery {
        table,
        field_index,
        criteria,
    })
}

fn parse_database_range(ctx: &dyn FunctionContext, database: ArgValue) -> Result<DatabaseTable, ErrorKind> {
    let array = arg_to_array(ctx, database)?;
    if array.rows == 0 || array.cols == 0 {
        return Err(ErrorKind::Value);
    }
    if array.rows < 1 {
        return Err(ErrorKind::Value);
    }

    let mut header_map: HashMap<String, usize> = HashMap::new();
    let mut saw_header = false;
    for col in 0..array.cols {
        let label = header_label(ctx, array.get(0, col).unwrap_or(&Value::Blank))?;
        if let Some(label) = label {
            saw_header = true;
            header_map.entry(casefold(label.trim())).or_insert(col);
        }
    }

    if !saw_header {
        // A database range without any labels is treated as invalid (missing headers).
        return Err(ErrorKind::Value);
    }

    Ok(DatabaseTable { array, header_map })
}

fn resolve_field(
    ctx: &dyn FunctionContext,
    table: &DatabaseTable,
    field: &Value,
) -> Result<usize, ErrorKind> {
    match field {
        Value::Error(e) => Err(*e),
        Value::Text(s) => {
            let key = casefold(s.trim());
            if key.is_empty() {
                return Err(ErrorKind::Value);
            }
            table
                .header_map
                .get(&key)
                .copied()
                .ok_or(ErrorKind::Value)
        }
        other => {
            let idx = other.coerce_to_i64_with_ctx(ctx)?;
            if idx <= 0 {
                return Err(ErrorKind::Value);
            }
            let idx = (idx - 1) as usize;
            if idx >= table.array.cols {
                return Err(ErrorKind::Value);
            }
            Ok(idx)
        }
    }
}

fn parse_criteria_range(
    ctx: &dyn FunctionContext,
    table: &DatabaseTable,
    criteria: ArgValue,
) -> Result<Vec<CriteriaClause>, ErrorKind> {
    let array = arg_to_array(ctx, criteria)?;
    if array.rows < 2 || array.cols == 0 {
        // Excel requires a header row plus at least one criteria row.
        return Err(ErrorKind::Value);
    }

    // Map each criteria column to a database column index.
    let mut col_map: Vec<Option<usize>> = Vec::with_capacity(array.cols);
    for col in 0..array.cols {
        let header_cell = array.get(0, col).unwrap_or(&Value::Blank);
        let label = header_label(ctx, header_cell)?;
        let Some(label) = label else {
            // Excel uses blank criteria headers for "computed criteria" (criteria formulas).
            // We do not currently implement that behavior. Allow entirely blank columns (no
            // criteria values) so callers can pass wider ranges without error.
            let mut any_nonblank = false;
            for row in 1..array.rows {
                if !matches!(array.get(row, col).unwrap_or(&Value::Blank), Value::Blank) {
                    any_nonblank = true;
                    break;
                }
            }
            if any_nonblank {
                return Err(ErrorKind::Value);
            }
            col_map.push(None);
            continue;
        };

        let key = casefold(label.trim());
        if key.is_empty() {
            return Err(ErrorKind::Value);
        }
        let db_col = table.header_map.get(&key).copied().ok_or(ErrorKind::Value)?;
        col_map.push(Some(db_col));
    }

    let mut clauses = Vec::with_capacity(array.rows - 1);
        for row in 1..array.rows {
            let mut conditions = Vec::new();
            for (col, mapping) in col_map.iter().enumerate() {
            let Some(db_col) = *mapping else {
                continue;
            };

            let crit_cell = array.get(row, col).unwrap_or(&Value::Blank);
            if matches!(crit_cell, Value::Blank) {
                continue;
            }

            let crit = Criteria::parse_with_date_system_and_locales(
                crit_cell,
                ctx.date_system(),
                ctx.value_locale(),
                ctx.now_utc(),
                ctx.locale_config(),
            )?;
            conditions.push((db_col, crit));
        }
        clauses.push(CriteriaClause { conditions });
    }

    Ok(clauses)
}

fn row_matches(table: &DatabaseTable, row: usize, criteria: &[CriteriaClause]) -> bool {
    if criteria.is_empty() {
        return true;
    }

    for clause in criteria {
        let mut ok = true;
        for (col, crit) in &clause.conditions {
            let v = table.cell(row, *col);
            if !crit.matches(v) {
                ok = false;
                break;
            }
        }
        if ok {
            return true;
        }
    }

    false
}

fn header_label(ctx: &dyn FunctionContext, value: &Value) -> Result<Option<String>, ErrorKind> {
    match value {
        Value::Blank => Ok(None),
        Value::Text(s) => {
            if s.trim().is_empty() {
                Ok(None)
            } else {
                Ok(Some(s.clone()))
            }
        }
        Value::Number(_) | Value::Bool(_) | Value::Entity(_) | Value::Record(_) => {
            let s = value.coerce_to_string_with_ctx(ctx)?;
            if s.trim().is_empty() {
                Ok(None)
            } else {
                Ok(Some(s))
            }
        }
        Value::Error(_) => Ok(None),
        // Header cells that evaluate to these internal runtime values are not considered valid
        // field names. Treat them as missing.
        Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Array(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => Ok(None),
    }
}

fn arg_to_array(ctx: &dyn FunctionContext, arg: ArgValue) -> Result<Array, ErrorKind> {
    match arg {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => Ok(arr),
            Value::Error(e) => Err(e),
            _ => Err(ErrorKind::Value),
        },
        ArgValue::Reference(r) => reference_to_array(ctx, r),
        ArgValue::ReferenceUnion(_) => Err(ErrorKind::Value),
    }
}

fn reference_to_array(ctx: &dyn FunctionContext, reference: crate::functions::Reference) -> Result<Array, ErrorKind> {
    let r = reference.normalized();
    let rows = (r.end.row - r.start.row + 1) as usize;
    let cols = (r.end.col - r.start.col + 1) as usize;
    let total = rows.saturating_mul(cols);
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Err(ErrorKind::Num);
    }
    values.resize(total, Value::Blank);

    // Use `iter_reference_cells` to allow sparse backends while still producing a dense array.
    for addr in ctx.iter_reference_cells(&r) {
        let row = (addr.row - r.start.row) as usize;
        let col = (addr.col - r.start.col) as usize;
        if row >= rows || col >= cols {
            continue;
        }
        let idx = row * cols + col;
        values[idx] = ctx.get_cell_value(&r.sheet_id, addr);
    }

    Ok(Array::new(rows, cols, values))
}
