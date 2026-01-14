use std::collections::HashMap;

use crate::eval::{compile_canonical_expr, is_valid_external_sheet_key, split_external_sheet_key};
use crate::functions::math::criteria::Criteria;
use crate::functions::{ArgValue, FunctionContext};
use crate::value::{casefold, Array, ErrorKind, Value};
use crate::{CellAddr as ParserCellAddr, ParseOptions, ReferenceStyle};
use formula_model::{EXCEL_MAX_COLS, EXCEL_MAX_ROWS};

#[derive(Debug, Clone)]
pub struct DatabaseTable {
    pub(crate) array: Array,
    /// Case-folded database field names mapped to their column indices.
    pub(crate) header_map: HashMap<String, usize>,
    /// Original database range reference when the input was a reference.
    pub(crate) source_ref: Option<crate::functions::Reference>,
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
pub struct ComputedCriteria {
    pub(crate) expr: crate::Expr,
}

#[derive(Debug, Clone)]
pub struct CriteriaClause {
    pub(crate) conditions: Vec<(usize, Criteria)>,
    pub(crate) computed: Vec<ComputedCriteria>,
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
    pub fn iter_matching_rows<'a>(
        &'a self,
        ctx: &'a dyn FunctionContext,
    ) -> impl Iterator<Item = Result<usize, ErrorKind>> + 'a {
        MatchingRowsIter::new(ctx, self)
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

fn parse_database_range(
    ctx: &dyn FunctionContext,
    database: ArgValue,
) -> Result<DatabaseTable, ErrorKind> {
    let (array, source_ref) = match database {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => (arr, None),
            Value::Error(e) => return Err(e),
            _ => return Err(ErrorKind::Value),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            if let crate::functions::SheetId::External(key) = &r.sheet_id {
                // Database functions (DSUM/DGET/DCOUNT/...) require the database range to be a
                // single-sheet 2D rectangle. Even though the evaluator can expand external-workbook
                // 3D spans like `[Book.xlsx]Sheet1:Sheet3!A1:D4` into a multi-area reference,
                // Excel treats that form as an invalid database range.
                if !is_valid_external_sheet_key(key) {
                    return Err(ErrorKind::Value);
                }
            }
            let array = reference_to_array(ctx, r.clone())?;
            (array, Some(r))
        }
        ArgValue::ReferenceUnion(_) => return Err(ErrorKind::Value),
    };
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

    Ok(DatabaseTable {
        array,
        header_map,
        source_ref,
    })
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
            table.header_map.get(&key).copied().ok_or(ErrorKind::Value)
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
    let (array, criteria_ref) = match criteria {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => (arr, None),
            Value::Error(e) => return Err(e),
            _ => return Err(ErrorKind::Value),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            let array = reference_to_array(ctx, r.clone())?;
            (array, Some(r))
        }
        ArgValue::ReferenceUnion(_) => return Err(ErrorKind::Value),
    };
    if array.rows < 2 || array.cols == 0 {
        // Excel requires a header row plus at least one criteria row.
        return Err(ErrorKind::Value);
    }

    #[derive(Debug, Clone, Copy)]
    enum CriteriaColumn {
        Standard(usize),
        Computed,
    }

    // Map each criteria column to either a database column index (standard criteria) or a
    // computed-criteria column (formula criteria).
    //
    // Excel supports "computed criteria" by placing a formula in the criteria range whose header
    // does *not* match any database field name. A blank header is the most common pattern, but
    // any non-matching header label should be treated as computed criteria.
    let mut col_map: Vec<CriteriaColumn> = Vec::with_capacity(array.cols);
    for col in 0..array.cols {
        let header_cell = array.get(0, col).unwrap_or(&Value::Blank);
        let label = header_label(ctx, header_cell)?;
        if let Some(label) = label {
            let key = casefold(label.trim());
            if key.is_empty() {
                return Err(ErrorKind::Value);
            }
            if let Some(db_col) = table.header_map.get(&key).copied() {
                col_map.push(CriteriaColumn::Standard(db_col));
            } else {
                col_map.push(CriteriaColumn::Computed);
            }
        } else {
            // Excel uses blank criteria headers for "computed criteria" (criteria formulas), but
            // also allows any label that doesn't match a database field name.
            col_map.push(CriteriaColumn::Computed);
        }
    }

    let mut clauses = Vec::with_capacity(array.rows - 1);
    for row in 1..array.rows {
        let mut conditions = Vec::new();
        let mut computed = Vec::new();
        for (col, mapping) in col_map.iter().enumerate() {
            match mapping {
                CriteriaColumn::Standard(db_col) => {
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
                    conditions.push((*db_col, crit));
                }
                CriteriaColumn::Computed => {
                    let Some(criteria_ref) = criteria_ref.as_ref() else {
                        // No reference => no formulas are possible. Treat any non-blank cell as invalid.
                        if !matches!(array.get(row, col).unwrap_or(&Value::Blank), Value::Blank) {
                            return Err(ErrorKind::Value);
                        }
                        continue;
                    };

                    let addr = crate::eval::CellAddr {
                        row: criteria_ref
                            .start
                            .row
                            .checked_add(row as u32)
                            .ok_or(ErrorKind::Value)?,
                        col: criteria_ref
                            .start
                            .col
                            .checked_add(col as u32)
                            .ok_or(ErrorKind::Value)?,
                    };

                    if let Some(formula_text) = ctx.get_cell_formula(&criteria_ref.sheet_id, addr) {
                        let db_ref = table.source_ref.as_ref().ok_or(ErrorKind::Value)?;
                        let first_row = db_ref.start.row.checked_add(1).ok_or(ErrorKind::Value)?;
                        let origin = ParserCellAddr::new(first_row, db_ref.start.col);
                        let ast = crate::parse_formula(
                            formula_text,
                            ParseOptions {
                                locale: ctx.locale_config(),
                                reference_style: ReferenceStyle::A1,
                                normalize_relative_to: Some(origin),
                            },
                        )
                        .map_err(|_| ErrorKind::Value)?;
                        let mut expr = ast.expr;
                        if let crate::functions::SheetId::External(key) = &db_ref.sheet_id {
                            if !is_valid_external_sheet_key(key) {
                                return Err(ErrorKind::Value);
                            }
                            let (workbook, sheet) =
                                split_external_sheet_key(key).ok_or(ErrorKind::Value)?;
                            expr = qualify_unprefixed_sheet_references(&expr, workbook, sheet);
                        }
                        computed.push(ComputedCriteria { expr });
                        continue;
                    }

                    // Non-blank computed-criteria cells must contain a formula.
                    let crit_cell = array.get(row, col).unwrap_or(&Value::Blank);
                    if !matches!(crit_cell, Value::Blank) {
                        return Err(ErrorKind::Value);
                    }
                }
            }
        }
        clauses.push(CriteriaClause {
            conditions,
            computed,
        });
    }

    Ok(clauses)
}

struct MatchingRowsIter<'a> {
    ctx: &'a dyn FunctionContext,
    query: &'a DatabaseQuery,
    next_row: usize,
    done: bool,
}

impl<'a> MatchingRowsIter<'a> {
    fn new(ctx: &'a dyn FunctionContext, query: &'a DatabaseQuery) -> Self {
        Self {
            ctx,
            query,
            next_row: 1,
            done: false,
        }
    }
}

impl Iterator for MatchingRowsIter<'_> {
    type Item = Result<usize, ErrorKind>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        while self.next_row < self.query.table.array.rows {
            let row = self.next_row;
            self.next_row += 1;

            match row_matches(self.ctx, &self.query.table, row, &self.query.criteria) {
                Ok(true) => return Some(Ok(row)),
                Ok(false) => continue,
                Err(e) => {
                    self.done = true;
                    return Some(Err(e));
                }
            }
        }
        None
    }
}

fn row_matches(
    ctx: &dyn FunctionContext,
    table: &DatabaseTable,
    row: usize,
    criteria: &[CriteriaClause],
) -> Result<bool, ErrorKind> {
    if criteria.is_empty() {
        return Ok(true);
    }

    let mut any_clause = false;
    for clause in criteria {
        let mut ok = true;
        for (col, crit) in &clause.conditions {
            let v = table.cell(row, *col);
            if !crit.matches(v) {
                ok = false;
                break;
            }
        }

        // Evaluate computed criteria formulas, if any.
        if !clause.computed.is_empty() {
            let db_ref = table.source_ref.as_ref().ok_or(ErrorKind::Value)?;
            let current_sheet_for_compile = match db_ref.sheet_id {
                crate::functions::SheetId::Local(id) => id,
                crate::functions::SheetId::External(_) => ctx.current_sheet_id(),
            };
            let row_u32 = u32::try_from(row).map_err(|_| ErrorKind::Value)?;
            let origin = crate::eval::CellAddr {
                row: db_ref
                    .start
                    .row
                    .checked_add(row_u32)
                    .ok_or(ErrorKind::Value)?,
                col: db_ref.start.col,
            };

            for comp in &clause.computed {
                let mut resolve_sheet = |name: &str| ctx.resolve_sheet_name(name);
                let mut sheet_dimensions = |_sheet_id: usize| (EXCEL_MAX_ROWS, EXCEL_MAX_COLS);
                let compiled = compile_canonical_expr(
                    &comp.expr,
                    current_sheet_for_compile,
                    origin,
                    &mut resolve_sheet,
                    &mut sheet_dimensions,
                );
                let value = ctx.eval_formula(&compiled);
                let b = match value {
                    Value::Error(e) => return Err(e),
                    other => other.coerce_to_bool_with_ctx(ctx)?,
                };
                if !b {
                    ok = false;
                }
            }
        }

        if ok {
            any_clause = true;
        }
    }

    Ok(any_clause)
}

fn qualify_unprefixed_sheet_references(expr: &crate::Expr, workbook: &str, sheet: &str) -> crate::Expr {
    match expr {
        crate::Expr::Number(v) => crate::Expr::Number(v.clone()),
        crate::Expr::String(v) => crate::Expr::String(v.clone()),
        crate::Expr::Boolean(v) => crate::Expr::Boolean(*v),
        crate::Expr::Error(v) => crate::Expr::Error(v.clone()),
        crate::Expr::Missing => crate::Expr::Missing,
        crate::Expr::NameRef(n) => crate::Expr::NameRef(n.clone()),
        crate::Expr::StructuredRef(r) => crate::Expr::StructuredRef(r.clone()),
        crate::Expr::FieldAccess(access) => crate::Expr::FieldAccess(crate::FieldAccessExpr {
            base: Box::new(qualify_unprefixed_sheet_references(&access.base, workbook, sheet)),
            field: access.field.clone(),
        }),
        crate::Expr::CellRef(r) => {
            let mut r = r.clone();
            if r.workbook.is_none() && r.sheet.is_none() {
                r.workbook = Some(workbook.to_string());
                r.sheet = Some(crate::SheetRef::Sheet(sheet.to_string()));
            }
            crate::Expr::CellRef(r)
        }
        crate::Expr::ColRef(r) => {
            let mut r = r.clone();
            if r.workbook.is_none() && r.sheet.is_none() {
                r.workbook = Some(workbook.to_string());
                r.sheet = Some(crate::SheetRef::Sheet(sheet.to_string()));
            }
            crate::Expr::ColRef(r)
        }
        crate::Expr::RowRef(r) => {
            let mut r = r.clone();
            if r.workbook.is_none() && r.sheet.is_none() {
                r.workbook = Some(workbook.to_string());
                r.sheet = Some(crate::SheetRef::Sheet(sheet.to_string()));
            }
            crate::Expr::RowRef(r)
        }
        crate::Expr::Array(arr) => crate::Expr::Array(crate::ArrayLiteral {
            rows: arr
                .rows
                .iter()
                .map(|row| {
                    row.iter()
                        .map(|el| qualify_unprefixed_sheet_references(el, workbook, sheet))
                        .collect()
                })
                .collect(),
        }),
        crate::Expr::FunctionCall(call) => crate::Expr::FunctionCall(crate::FunctionCall {
            name: call.name.clone(),
            args: call
                .args
                .iter()
                .map(|arg| qualify_unprefixed_sheet_references(arg, workbook, sheet))
                .collect(),
        }),
        crate::Expr::Call(call) => crate::Expr::Call(crate::CallExpr {
            callee: Box::new(qualify_unprefixed_sheet_references(
                &call.callee,
                workbook,
                sheet,
            )),
            args: call
                .args
                .iter()
                .map(|arg| qualify_unprefixed_sheet_references(arg, workbook, sheet))
                .collect(),
        }),
        crate::Expr::Unary(u) => crate::Expr::Unary(crate::UnaryExpr {
            op: u.op,
            expr: Box::new(qualify_unprefixed_sheet_references(&u.expr, workbook, sheet)),
        }),
        crate::Expr::Postfix(p) => crate::Expr::Postfix(crate::PostfixExpr {
            op: p.op,
            expr: Box::new(qualify_unprefixed_sheet_references(&p.expr, workbook, sheet)),
        }),
        crate::Expr::Binary(b) => crate::Expr::Binary(crate::BinaryExpr {
            op: b.op,
            left: Box::new(qualify_unprefixed_sheet_references(&b.left, workbook, sheet)),
            right: Box::new(qualify_unprefixed_sheet_references(&b.right, workbook, sheet)),
        }),
    }
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

fn reference_to_array(
    ctx: &dyn FunctionContext,
    reference: crate::functions::Reference,
) -> Result<Array, ErrorKind> {
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
