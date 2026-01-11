use crate::engine::{DaxResult, FilterContext, RowContext};
use crate::parser::Expr;
use crate::{DaxEngine, DataModel, Value};
use std::collections::HashSet;

/// A group-by column used by the pivot engine.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct GroupByColumn {
    pub table: String,
    pub column: String,
}

impl GroupByColumn {
    pub fn new(table: impl Into<String>, column: impl Into<String>) -> Self {
        Self {
            table: table.into(),
            column: column.into(),
        }
    }
}

/// A measure expression to evaluate for each pivot group.
#[derive(Clone, Debug)]
pub struct PivotMeasure {
    pub name: String,
    pub expression: String,
    pub(crate) parsed: Expr,
}

impl PivotMeasure {
    pub fn new(name: impl Into<String>, expression: impl Into<String>) -> DaxResult<Self> {
        let name = name.into();
        let expression = expression.into();
        let parsed = crate::parser::parse(&expression)?;
        Ok(Self {
            name,
            expression,
            parsed,
        })
    }
}

/// The result of a pivot/group-by query.
#[derive(Clone, Debug, PartialEq)]
pub struct PivotResult {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<Value>>,
}

/// Compute a grouped table suitable for rendering a pivot table.
///
/// This API is intentionally small: it takes a base table (typically the fact table),
/// a set of group-by columns, and a list of measure expressions.
pub fn pivot(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<PivotResult> {
    let engine = DaxEngine::new();

    let base_rows = crate::engine::resolve_table_rows(model, filter, base_table)?;
    let mut seen: HashSet<Vec<Value>> = HashSet::new();
    let mut groups: Vec<Vec<Value>> = Vec::new();
    let group_exprs: Vec<Expr> = group_by
        .iter()
        .map(|col| {
            if col.table == base_table {
                Expr::ColumnRef {
                    table: col.table.clone(),
                    column: col.column.clone(),
                }
            } else {
                Expr::Call {
                    name: "RELATED".to_string(),
                    args: vec![Expr::ColumnRef {
                        table: col.table.clone(),
                        column: col.column.clone(),
                    }],
                }
            }
        })
        .collect();

    // Build the set of groups by scanning the base table rows. This ensures we only create
    // groups that actually exist in the fact table under the current filter context.
    for row in base_rows {
        let mut row_ctx = RowContext::default();
        row_ctx.push(base_table, row);

        let mut key = Vec::with_capacity(group_by.len());
        for expr in &group_exprs {
            let value = engine.evaluate_expr(model, expr, filter, &row_ctx)?;
            key.push(value);
        }

        if seen.insert(key.clone()) {
            groups.push(key);
        }
    }

    let mut columns: Vec<String> = group_by
        .iter()
        .map(|c| format!("{}[{}]", c.table, c.column))
        .collect();
    columns.extend(measures.iter().map(|m| m.name.clone()));

    let mut rows_out = Vec::with_capacity(groups.len());
    for key in groups {
        let mut group_filter = filter.clone();
        for (col, value) in group_by.iter().zip(key.iter()) {
            group_filter.set_column_equals(&col.table, &col.column, value.clone());
        }

        let mut row = key;
        for measure in measures {
            let value = engine.evaluate_expr(model, &measure.parsed, &group_filter, &RowContext::default())?;
            row.push(value);
        }
        rows_out.push(row);
    }

    Ok(PivotResult {
        columns,
        rows: rows_out,
    })
}
