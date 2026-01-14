use crate::engine::{DaxError, DaxResult};
use crate::model::normalize_ident;
use crate::pivot::{GroupByColumn, PivotMeasure};
use crate::{DataModel, FilterContext};

use formula_model::pivots::{AggregationType, PivotConfig, PivotFieldRef, PivotSource};

/// Planned Data Model pivot query derived from `formula_model` pivot schema types.
#[derive(Clone, Debug)]
pub struct DataModelPivotPlan {
    pub base_table: String,
    pub group_by: Vec<GroupByColumn>,
    pub measures: Vec<PivotMeasure>,
    pub filter: FilterContext,
}

/// Build a `formula_dax::pivot` plan from a canonical `formula_model` pivot config + source.
///
/// This adapter is responsible for:
/// - Ensuring Data Model pivots reference columns/measures explicitly (no ambiguous strings).
/// - Producing `GroupByColumn` and `PivotMeasure` objects for the `formula_dax::pivot` engine.
/// - Failing with clear errors when references cannot be resolved against the `DataModel`.
pub fn build_data_model_pivot_plan(
    model: &DataModel,
    source: &PivotSource,
    cfg: &PivotConfig,
) -> DaxResult<DataModelPivotPlan> {
    let PivotSource::DataModel { table: base_table } = source else {
        return Err(DaxError::Eval(
            "build_data_model_pivot_plan requires PivotSource::DataModel".to_string(),
        ));
    };

    cfg.validate_for_source(source)
        .map_err(|e| DaxError::Eval(format!("invalid pivot config: {e}")))?;

    // Resolve base table up-front.
    model
        .table(base_table)
        .ok_or_else(|| DaxError::UnknownTable(base_table.clone()))?;

    let mut group_by = Vec::with_capacity(cfg.row_fields.len() + cfg.column_fields.len());
    for field in cfg.row_fields.iter().chain(cfg.column_fields.iter()) {
        let PivotFieldRef::DataModelColumn { table, column } = &field.source_field else {
            return Err(DaxError::Eval(
                "row/column fields must reference data model columns".to_string(),
            ));
        };
        resolve_column(model, table, column)?;
        group_by.push(GroupByColumn::new(table, column));
    }

    // TODO: Filter fields should map to `FilterContext` once a stable filter schema exists. For
    // now we validate the refs and return an empty filter context.
    for filter_field in &cfg.filter_fields {
        let PivotFieldRef::DataModelColumn { table, column } = &filter_field.source_field else {
            return Err(DaxError::Eval(
                "filter fields must reference data model columns".to_string(),
            ));
        };
        resolve_column(model, table, column)?;
    }

    let mut measures = Vec::with_capacity(cfg.value_fields.len());
    for value_field in &cfg.value_fields {
        let expr = match &value_field.source_field {
            PivotFieldRef::DataModelMeasure(measure) => {
                let key = normalize_ident(DataModel::normalize_measure_name(measure));
                if model.measures().get(&key).is_none() {
                    return Err(DaxError::UnknownMeasure(measure.clone()));
                }
                format_dax_measure_ref(measure)
            }
            PivotFieldRef::DataModelColumn { table, column } => {
                resolve_column(model, table, column)?;
                build_aggregation_expr(value_field.aggregation, table, column)?
            }
            PivotFieldRef::CacheFieldName(_) => {
                return Err(DaxError::Eval(
                    "value fields must reference data model measures or columns".to_string(),
                ))
            }
        };

        measures.push(PivotMeasure::new(value_field.name.clone(), expr)?);
    }

    Ok(DataModelPivotPlan {
        base_table: base_table.clone(),
        group_by,
        measures,
        filter: FilterContext::empty(),
    })
}

fn resolve_column(model: &DataModel, table: &str, column: &str) -> DaxResult<()> {
    let table_ref = model
        .table(table)
        .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
    if table_ref.column_idx(column).is_none() {
        return Err(DaxError::UnknownColumn {
            table: table.to_string(),
            column: column.to_string(),
        });
    }
    Ok(())
}

fn build_aggregation_expr(agg: AggregationType, table: &str, column: &str) -> DaxResult<String> {
    let col_ref = format_dax_column_ref(table, column);
    let expr = match agg {
        AggregationType::Sum => format!("SUM({col_ref})"),
        AggregationType::Average => format!("AVERAGE({col_ref})"),
        AggregationType::Min => format!("MIN({col_ref})"),
        AggregationType::Max => format!("MAX({col_ref})"),
        AggregationType::Count | AggregationType::CountNumbers => {
            // Best-effort mapping. In DAX, `COUNTX(Table, Table[Column])` counts non-blank values.
            let table_ref = format_dax_table_name(table);
            format!("COUNTX({table_ref}, {col_ref})")
        }
        other => {
            return Err(DaxError::Eval(format!(
                "aggregation {other:?} is not supported for data model pivots"
            )))
        }
    };
    Ok(expr)
}

fn format_dax_table_name(table: &str) -> String {
    // Always quote table names to avoid edge cases with spaces/reserved words.
    let escaped = table.replace('\'', "''");
    format!("'{escaped}'")
}

fn format_dax_column_ref(table: &str, column: &str) -> String {
    format!("{}[{}]", format_dax_table_name(table), column)
}

fn format_dax_measure_ref(measure: &str) -> String {
    let name = DataModel::normalize_measure_name(measure);
    format!("[{name}]")
}
