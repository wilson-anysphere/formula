use crate::backend::{AggregationKind, AggregationSpec, TableBackend};
use crate::engine::{DaxError, DaxResult, FilterContext, RowContext};
use crate::parser::Expr;
use crate::{DaxEngine, DataModel, Value};
use std::cmp::Ordering;
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

fn cmp_value(a: &Value, b: &Value) -> Ordering {
    match (a, b) {
        (Value::Blank, Value::Blank) => Ordering::Equal,
        (Value::Blank, _) => Ordering::Less,
        (_, Value::Blank) => Ordering::Greater,
        (Value::Boolean(a), Value::Boolean(b)) => a.cmp(b),
        (Value::Boolean(_), _) => Ordering::Less,
        (_, Value::Boolean(_)) => Ordering::Greater,
        (Value::Number(a), Value::Number(b)) => a.cmp(b),
        (Value::Number(_), _) => Ordering::Less,
        (_, Value::Number(_)) => Ordering::Greater,
        (Value::Text(a), Value::Text(b)) => a.as_ref().cmp(b.as_ref()),
    }
}

fn cmp_key(a: &[Value], b: &[Value]) -> Ordering {
    for (a, b) in a.iter().zip(b.iter()) {
        let ord = cmp_value(a, b);
        if ord != Ordering::Equal {
            return ord;
        }
    }
    a.len().cmp(&b.len())
}

fn extract_aggregation(
    model: &DataModel,
    base_table: &crate::model::Table,
    base_table_name: &str,
    expr: &Expr,
    depth: usize,
) -> DaxResult<Option<AggregationSpec>> {
    if depth > 16 {
        return Ok(None);
    }

    match expr {
        Expr::Measure(name) => {
            let normalized = DataModel::normalize_measure_name(name);
            let measure = model
                .measures()
                .get(normalized)
                .ok_or_else(|| DaxError::UnknownMeasure(name.clone()))?;
            extract_aggregation(
                model,
                base_table,
                base_table_name,
                &measure.parsed,
                depth + 1,
            )
        }
        Expr::Call { name, args } => match name.to_ascii_uppercase().as_str() {
            "SUM" | "AVERAGE" | "MIN" | "MAX" | "DISTINCTCOUNT" => {
                let [arg] = args.as_slice() else {
                    return Ok(None);
                };
                let Expr::ColumnRef { table, column } = arg else {
                    return Ok(None);
                };
                if table != base_table_name {
                    return Ok(None);
                }
                let idx = base_table.column_idx(column).ok_or_else(|| DaxError::UnknownColumn {
                    table: table.clone(),
                    column: column.clone(),
                })?;
                let kind = match name.to_ascii_uppercase().as_str() {
                    "SUM" => AggregationKind::Sum,
                    "AVERAGE" => AggregationKind::Average,
                    "MIN" => AggregationKind::Min,
                    "MAX" => AggregationKind::Max,
                    "DISTINCTCOUNT" => AggregationKind::DistinctCount,
                    _ => unreachable!(),
                };
                Ok(Some(AggregationSpec {
                    kind,
                    column_idx: Some(idx),
                }))
            }
            "COUNTROWS" => {
                let [arg] = args.as_slice() else {
                    return Ok(None);
                };
                let Expr::TableName(table) = arg else {
                    return Ok(None);
                };
                if table != base_table_name {
                    return Ok(None);
                }
                Ok(Some(AggregationSpec {
                    kind: AggregationKind::CountRows,
                    column_idx: None,
                }))
            }
            _ => Ok(None),
        },
        _ => Ok(None),
    }
}

fn pivot_columnar_group_by(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<Option<PivotResult>> {
    if group_by.iter().any(|c| c.table != base_table) {
        return Ok(None);
    }

    let table_ref = model
        .table(base_table)
        .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;

    let mut group_idxs = Vec::with_capacity(group_by.len());
    for col in group_by {
        let idx = table_ref.column_idx(&col.column).ok_or_else(|| DaxError::UnknownColumn {
            table: base_table.to_string(),
            column: col.column.clone(),
        })?;
        group_idxs.push(idx);
    }

    let mut aggs = Vec::with_capacity(measures.len());
    for measure in measures {
        match extract_aggregation(model, table_ref, base_table, &measure.parsed, 0)? {
            Some(spec) => aggs.push(spec),
            None => return Ok(None),
        }
    }

    let rows_buffer;
    let rows = if filter.is_empty() {
        None
    } else {
        rows_buffer = crate::engine::resolve_table_rows(model, filter, base_table)?;
        Some(rows_buffer.as_slice())
    };

    let Some(mut grouped_rows) = table_ref.group_by_aggregations(&group_idxs, &aggs, rows) else {
        return Ok(None);
    };

    let key_len = group_idxs.len();
    grouped_rows.sort_by(|a, b| cmp_key(&a[..key_len], &b[..key_len]));

    let mut columns: Vec<String> = group_by
        .iter()
        .map(|c| format!("{}[{}]", c.table, c.column))
        .collect();
    columns.extend(measures.iter().map(|m| m.name.clone()));

    Ok(Some(PivotResult {
        columns,
        rows: grouped_rows,
    }))
}

fn pivot_row_scan(
    model: &DataModel,
    base_table: &str,
    group_by: &[GroupByColumn],
    measures: &[PivotMeasure],
    filter: &FilterContext,
) -> DaxResult<PivotResult> {
    let engine = DaxEngine::new();

    let table_ref = model
        .table(base_table)
        .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;
    let base_rows = (!filter.is_empty())
        .then(|| crate::engine::resolve_table_rows(model, filter, base_table))
        .transpose()?;
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
    let mut process_row = |row: usize| -> DaxResult<()> {
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
        Ok(())
    };

    if let Some(rows) = base_rows {
        for row in rows {
            process_row(row)?;
        }
    } else {
        for row in 0..table_ref.row_count() {
            process_row(row)?;
        }
    }

    groups.sort_by(|a, b| cmp_key(a, b));

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
            let value =
                engine.evaluate_expr(model, &measure.parsed, &group_filter, &RowContext::default())?;
            row.push(value);
        }
        rows_out.push(row);
    }

    Ok(PivotResult {
        columns,
        rows: rows_out,
    })
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
    if let Some(result) = pivot_columnar_group_by(model, base_table, group_by, measures, filter)? {
        return Ok(result);
    }

    pivot_row_scan(model, base_table, group_by, measures, filter)
}

#[cfg(test)]
mod tests {
    use super::*;
    use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
    use std::sync::Arc;
    use std::time::Instant;

    #[test]
    fn pivot_benchmark_old_vs_new_columnar() {
        if std::env::var_os("FORMULA_DAX_PIVOT_BENCH").is_none() {
            return;
        }

        let rows = 1_000_000usize;
        let schema = vec![
            ColumnSchema {
                name: "Group".to_string(),
                column_type: ColumnType::String,
            },
            ColumnSchema {
                name: "Amount".to_string(),
                column_type: ColumnType::Number,
            },
        ];
        let options = TableOptions {
            page_size_rows: 65_536,
            cache: PageCacheConfig { max_entries: 8 },
        };
        let mut builder = ColumnarTableBuilder::new(schema, options);
        let groups = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];
        for i in 0..rows {
            builder.append_row(&[
                formula_columnar::Value::String(Arc::<str>::from(groups[i % groups.len()])),
                formula_columnar::Value::Number((i % 100) as f64),
            ]);
        }

        let mut model = DataModel::new();
        model
            .add_table(crate::Table::from_columnar("Fact", builder.finalize()))
            .unwrap();
        model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

        let measures = vec![PivotMeasure::new("Total", "[Total]").unwrap()];
        let group_by = vec![GroupByColumn::new("Fact", "Group")];
        let filter = FilterContext::empty();

        let start = Instant::now();
        let scan = pivot_row_scan(&model, "Fact", &group_by, &measures, &filter).unwrap();
        let scan_elapsed = start.elapsed();

        let start = Instant::now();
        let fast = pivot(&model, "Fact", &group_by, &measures, &filter).unwrap();
        let fast_elapsed = start.elapsed();

        assert_eq!(scan, fast);

        println!(
            "pivot row-scan: {:?}, columnar group-by: {:?} ({:.2}x speedup)",
            scan_elapsed,
            fast_elapsed,
            scan_elapsed.as_secs_f64() / fast_elapsed.as_secs_f64()
        );
    }
}
