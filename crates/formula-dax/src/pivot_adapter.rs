use crate::engine::{DaxError, DaxResult};
use crate::ident::{format_dax_column_ref, format_dax_measure_ref, format_dax_table_name};
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
    let base_table = model
        .table(base_table)
        .map(|t| t.name().to_string())
        .ok_or_else(|| DaxError::UnknownTable(base_table.clone()))?;

    let mut group_by = Vec::new();
    let _ = group_by.try_reserve_exact(cfg.row_fields.len() + cfg.column_fields.len());
    for field in cfg.row_fields.iter().chain(cfg.column_fields.iter()) {
        let PivotFieldRef::DataModelColumn { table, column } = &field.source_field else {
            return Err(DaxError::Eval(
                "row/column fields must reference data model columns".to_string(),
            ));
        };
        let (table, column) = resolve_column_canonical(model, table, column)?;
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
        let _ = resolve_column_canonical(model, table, column)?;
    }

    let mut measures = Vec::new();
    let _ = measures.try_reserve_exact(cfg.value_fields.len());
    for value_field in &cfg.value_fields {
        let expr = match &value_field.source_field {
            PivotFieldRef::DataModelMeasure(measure) => {
                let key = normalize_ident(DataModel::normalize_measure_name(measure));
                let measure = model
                    .measures()
                    .get(&key)
                    .ok_or_else(|| DaxError::UnknownMeasure(measure.clone()))?;
                format_dax_measure_ref(&measure.name)
            }
            PivotFieldRef::DataModelColumn { table, column } => {
                let (table, column) = resolve_column_canonical(model, table, column)?;
                build_aggregation_expr(value_field.aggregation, &table, &column)?
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
        base_table,
        group_by,
        measures,
        filter: FilterContext::empty(),
    })
}

fn resolve_column_canonical(model: &DataModel, table: &str, column: &str) -> DaxResult<(String, String)> {
    let table_ref = model
        .table(table)
        .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
    let idx = table_ref.column_idx(column).ok_or_else(|| DaxError::UnknownColumn {
        table: table.to_string(),
        column: column.to_string(),
    })?;
    let table = table_ref.name().to_string();
    let column = table_ref
        .columns()
        .get(idx)
        .cloned()
        .unwrap_or_else(|| column.to_string());
    Ok((table, column))
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Cardinality, CrossFilterDirection, Relationship, Table, Value};

    use formula_model::pivots::{
        AggregationType, GrandTotals, Layout, PivotField, PivotFieldRef, PivotSource, SubtotalPosition,
        ValueField,
    };
    use pretty_assertions::assert_eq;

    #[test]
    fn build_data_model_pivot_plan_is_case_insensitive_and_canonicalizes_names() {
        let mut model = DataModel::new();
        let mut customers = Table::new("Customers", vec!["CustomerId", "Region"]);
        customers.push_row(vec![1.into(), "East".into()]).unwrap();
        customers.push_row(vec![2.into(), "West".into()]).unwrap();
        model.add_table(customers).unwrap();

        let mut orders = Table::new("Orders", vec!["OrderId", "CustomerId", "Amount"]);
        orders
            .push_row(vec![100.into(), 1.into(), 10.0.into()])
            .unwrap();
        orders
            .push_row(vec![101.into(), 1.into(), 20.0.into()])
            .unwrap();
        orders
            .push_row(vec![102.into(), 2.into(), 5.0.into()])
            .unwrap();
        model.add_table(orders).unwrap();

        model
            .add_relationship(Relationship {
                name: "Orders->Customers".into(),
                from_table: "Orders".into(),
                from_column: "CustomerId".into(),
                to_table: "Customers".into(),
                to_column: "CustomerId".into(),
                cardinality: Cardinality::OneToMany,
                cross_filter_direction: CrossFilterDirection::Single,
                is_active: true,
                enforce_referential_integrity: false,
            })
            .unwrap();

        model
            .add_measure("Total Sales", "SUM(Orders[Amount])")
            .unwrap();

        let source = PivotSource::DataModel {
            table: "orders".to_string(),
        };

        let row_field = PivotField {
            source_field: PivotFieldRef::DataModelColumn {
                table: "customers".to_string(),
                column: "region".to_string(),
            },
            sort_order: Default::default(),
            manual_sort: None,
        };

        let value_field = ValueField {
            source_field: PivotFieldRef::DataModelMeasure("total sales".to_string()),
            name: "Total Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        };

        let cfg = PivotConfig {
            row_fields: vec![row_field],
            column_fields: vec![],
            value_fields: vec![value_field],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals::default(),
        };

        let plan = build_data_model_pivot_plan(&model, &source, &cfg).unwrap();
        assert_eq!(plan.base_table, "Orders");
        assert_eq!(
            plan.group_by,
            vec![GroupByColumn::new("Customers", "Region")]
        );
        assert_eq!(plan.measures[0].expression, "[Total Sales]");

        let result = crate::pivot::pivot(
            &model,
            &plan.base_table,
            &plan.group_by,
            &plan.measures,
            &plan.filter,
        )
        .unwrap();
        assert_eq!(result.columns, vec!["Customers[Region]", "Total Sales"]);
        assert_eq!(
            result.rows,
            vec![
                vec![Value::from("East"), Value::from(30.0)],
                vec![Value::from("West"), Value::from(5.0)],
            ]
        );
    }

    #[test]
    fn build_data_model_pivot_plan_escapes_bracket_identifiers_in_generated_dax() {
        let mut model = DataModel::new();
        let mut orders = Table::new("Orders", vec!["Region]Name", "Amount]USD"]);
        orders
            .push_row(vec![Value::from("East"), Value::from(10.0)])
            .unwrap();
        orders
            .push_row(vec![Value::from("East"), Value::from(20.0)])
            .unwrap();
        orders
            .push_row(vec![Value::from("West"), Value::from(5.0)])
            .unwrap();
        model.add_table(orders).unwrap();

        model
            .add_measure("Total]USD", "SUM(Orders[Amount]]USD])")
            .unwrap();

        let source = PivotSource::DataModel {
            table: "orders".to_string(),
        };

        let row_field = PivotField {
            source_field: PivotFieldRef::DataModelColumn {
                table: "orders".to_string(),
                column: "region]name".to_string(),
            },
            sort_order: Default::default(),
            manual_sort: None,
        };

        let sum_amount_field = ValueField {
            source_field: PivotFieldRef::DataModelColumn {
                table: "orders".to_string(),
                column: "amount]usd".to_string(),
            },
            name: "Sum Amount".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        };

        let total_measure_field = ValueField {
            source_field: PivotFieldRef::DataModelMeasure("total]usd".to_string()),
            name: "Total Measure".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        };

        let cfg = PivotConfig {
            row_fields: vec![row_field],
            column_fields: vec![],
            value_fields: vec![sum_amount_field, total_measure_field],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals::default(),
        };

        let plan = build_data_model_pivot_plan(&model, &source, &cfg).unwrap();
        assert_eq!(plan.base_table, "Orders");
        assert_eq!(
            plan.group_by,
            vec![GroupByColumn::new("Orders", "Region]Name")]
        );
        assert_eq!(
            plan.measures
                .iter()
                .map(|m| m.expression.as_str())
                .collect::<Vec<_>>(),
            vec!["SUM('Orders'[Amount]]USD])", "[Total]]USD]"]
        );
    }
}
