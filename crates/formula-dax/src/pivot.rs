use crate::backend::{AggregationKind, AggregationSpec, TableBackend};
use crate::engine::{DaxError, DaxResult, FilterContext, RowContext};
use crate::parser::{BinaryOp, Expr, UnaryOp};
use crate::{DaxEngine, DataModel, Value};
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

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

#[derive(Clone, Debug)]
enum PlannedExpr {
    Const(Value),
    AggRef(usize),
    Negate(Box<PlannedExpr>),
    Binary {
        op: BinaryOp,
        left: Box<PlannedExpr>,
        right: Box<PlannedExpr>,
    },
    Divide {
        numerator: Box<PlannedExpr>,
        denominator: Box<PlannedExpr>,
        alternate: Option<Box<PlannedExpr>>,
    },
    Coalesce(Vec<PlannedExpr>),
}

fn eval_planned(expr: &PlannedExpr, agg_values: &[Value]) -> Value {
    match expr {
        PlannedExpr::Const(v) => v.clone(),
        PlannedExpr::AggRef(idx) => agg_values.get(*idx).cloned().unwrap_or(Value::Blank),
        PlannedExpr::Negate(inner) => {
            let v = eval_planned(inner, agg_values);
            Value::from(-v.as_f64().unwrap_or(0.0))
        }
        PlannedExpr::Binary { op, left, right } => {
            let l = eval_planned(left, agg_values).as_f64().unwrap_or(0.0);
            let r = eval_planned(right, agg_values).as_f64().unwrap_or(0.0);
            let out = match op {
                BinaryOp::Add => l + r,
                BinaryOp::Subtract => l - r,
                BinaryOp::Multiply => l * r,
                BinaryOp::Divide => l / r,
                _ => return Value::Blank,
            };
            Value::from(out)
        }
        PlannedExpr::Divide {
            numerator,
            denominator,
            alternate,
        } => {
            let num = eval_planned(numerator, agg_values);
            let denom = eval_planned(denominator, agg_values);
            let denom = denom.as_f64().unwrap_or(0.0);
            if denom == 0.0 {
                alternate
                    .as_ref()
                    .map(|alt| eval_planned(alt, agg_values))
                    .unwrap_or(Value::Blank)
            } else {
                let num = num.as_f64().unwrap_or(0.0);
                Value::from(num / denom)
            }
        }
        PlannedExpr::Coalesce(args) => {
            for arg in args {
                let value = eval_planned(arg, agg_values);
                if !value.is_blank() {
                    return value;
                }
            }
            Value::Blank
        }
    }
}

fn ensure_agg(
    kind: AggregationKind,
    column_idx: Option<usize>,
    agg_specs: &mut Vec<AggregationSpec>,
    agg_map: &mut HashMap<(AggregationKind, Option<usize>), usize>,
) -> usize {
    let key = (kind, column_idx);
    if let Some(&idx) = agg_map.get(&key) {
        return idx;
    }
    let idx = agg_specs.len();
    agg_specs.push(AggregationSpec { kind, column_idx });
    agg_map.insert(key, idx);
    idx
}

fn plan_pivot_expr(
    model: &DataModel,
    base_table: &crate::model::Table,
    base_table_name: &str,
    expr: &Expr,
    depth: usize,
    agg_specs: &mut Vec<AggregationSpec>,
    agg_map: &mut HashMap<(AggregationKind, Option<usize>), usize>,
) -> DaxResult<Option<PlannedExpr>> {
    if depth > 32 {
        return Ok(None);
    }

    match expr {
        Expr::Number(n) => Ok(Some(PlannedExpr::Const(Value::from(*n)))),
        Expr::Text(s) => Ok(Some(PlannedExpr::Const(Value::from(s.clone())))),
        Expr::Boolean(b) => Ok(Some(PlannedExpr::Const(Value::from(*b)))),
        Expr::Measure(name) => {
            let normalized = DataModel::normalize_measure_name(name);
            let measure = model
                .measures()
                .get(normalized)
                .ok_or_else(|| DaxError::UnknownMeasure(name.clone()))?;
            plan_pivot_expr(
                model,
                base_table,
                base_table_name,
                &measure.parsed,
                depth + 1,
                agg_specs,
                agg_map,
            )
        }
        Expr::UnaryOp { op, expr } => match op {
            UnaryOp::Negate => {
                let planned = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    expr,
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?;
                Ok(planned.map(|inner| PlannedExpr::Negate(Box::new(inner))))
            }
        },
        Expr::BinaryOp { op, left, right } => match op {
            BinaryOp::Add | BinaryOp::Subtract | BinaryOp::Multiply | BinaryOp::Divide => {
                let Some(left) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    left,
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                let Some(right) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    right,
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                Ok(Some(PlannedExpr::Binary {
                    op: *op,
                    left: Box::new(left),
                    right: Box::new(right),
                }))
            }
            _ => Ok(None),
        },
        Expr::Call { name, args } => match name.to_ascii_uppercase().as_str() {
            "BLANK" if args.is_empty() => Ok(Some(PlannedExpr::Const(Value::Blank))),
            "TRUE" if args.is_empty() => Ok(Some(PlannedExpr::Const(Value::from(true)))),
            "FALSE" if args.is_empty() => Ok(Some(PlannedExpr::Const(Value::from(false)))),
            "COALESCE" => {
                if args.is_empty() {
                    return Ok(None);
                }
                let mut planned_args = Vec::with_capacity(args.len());
                for arg in args {
                    let Some(planned) = plan_pivot_expr(
                        model,
                        base_table,
                        base_table_name,
                        arg,
                        depth + 1,
                        agg_specs,
                        agg_map,
                    )?
                    else {
                        return Ok(None);
                    };
                    planned_args.push(planned);
                }
                Ok(Some(PlannedExpr::Coalesce(planned_args)))
            }
            "DIVIDE" => {
                if args.len() < 2 || args.len() > 3 {
                    return Ok(None);
                }
                let Some(numerator) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    &args[0],
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                let Some(denominator) = plan_pivot_expr(
                    model,
                    base_table,
                    base_table_name,
                    &args[1],
                    depth + 1,
                    agg_specs,
                    agg_map,
                )?
                else {
                    return Ok(None);
                };
                let alternate = if args.len() == 3 {
                    let Some(alt) = plan_pivot_expr(
                        model,
                        base_table,
                        base_table_name,
                        &args[2],
                        depth + 1,
                        agg_specs,
                        agg_map,
                    )?
                    else {
                        return Ok(None);
                    };
                    Some(Box::new(alt))
                } else {
                    None
                };
                Ok(Some(PlannedExpr::Divide {
                    numerator: Box::new(numerator),
                    denominator: Box::new(denominator),
                    alternate,
                }))
            }
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
                let agg_idx = ensure_agg(kind, Some(idx), agg_specs, agg_map);
                Ok(Some(PlannedExpr::AggRef(agg_idx)))
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
                let agg_idx = ensure_agg(AggregationKind::CountRows, None, agg_specs, agg_map);
                Ok(Some(PlannedExpr::AggRef(agg_idx)))
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

    let mut agg_specs: Vec<AggregationSpec> = Vec::new();
    let mut agg_map: HashMap<(AggregationKind, Option<usize>), usize> = HashMap::new();
    let mut plans: Vec<PlannedExpr> = Vec::with_capacity(measures.len());
    for measure in measures {
        let Some(plan) = plan_pivot_expr(
            model,
            table_ref,
            base_table,
            &measure.parsed,
            0,
            &mut agg_specs,
            &mut agg_map,
        )?
        else {
            return Ok(None);
        };
        plans.push(plan);
    }

    let rows_buffer;
    let rows = if filter.is_empty() {
        None
    } else {
        rows_buffer = crate::engine::resolve_table_rows(model, filter, base_table)?;
        Some(rows_buffer.as_slice())
    };

    let Some(grouped_rows) = table_ref.group_by_aggregations(&group_idxs, &agg_specs, rows) else {
        return Ok(None);
    };

    let key_len = group_idxs.len();
    let mut rows_out: Vec<Vec<Value>> = Vec::with_capacity(grouped_rows.len());
    for mut row in grouped_rows {
        let agg_values = row.get(key_len..).unwrap_or(&[]);
        let mut measure_values = Vec::with_capacity(plans.len());
        for plan in &plans {
            measure_values.push(eval_planned(plan, agg_values));
        }
        row.truncate(key_len);
        row.extend(measure_values);
        rows_out.push(row);
    }

    rows_out.sort_by(|a, b| cmp_key(&a[..key_len], &b[..key_len]));

    let mut columns: Vec<String> = group_by
        .iter()
        .map(|c| format!("{}[{}]", c.table, c.column))
        .collect();
    columns.extend(measures.iter().map(|m| m.name.clone()));

    Ok(Some(PivotResult {
        columns,
        rows: rows_out,
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
