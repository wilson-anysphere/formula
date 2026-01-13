//! Helpers for executing a `formula_model::pivots::PivotConfig` against a [`DataModel`].
//!
//! `formula-model` defines a canonical, serialization-friendly pivot configuration schema that is
//! shared by both worksheet pivots and Data Model pivots. `formula-dax` implements Data Model pivots
//! by executing a pivot/group-by query (`GroupByColumn` + `PivotMeasure`) under a `FilterContext`.
//!
//! This module bridges those worlds by converting a `PivotConfig` into:
//! - `row_fields` + `column_fields` as `Vec<GroupByColumn>`
//! - `value_fields` as `Vec<PivotMeasure>`
//! - `filter_fields` as a `FilterContext`
//!
//! # Field identifier parsing (MVP)
//! Pivot fields are identified by a string in one of two forms:
//!
//! - `Table[Column]` — explicit table and column.
//! - `Column` — shorthand for `base_table[Column]`.
//!
//! These rules apply to:
//! - `PivotConfig.row_fields[*].source_field`
//! - `PivotConfig.column_fields[*].source_field`
//! - `PivotConfig.filter_fields[*].source_field`
//!
//! Unknown tables/columns are validated eagerly and reported as `DaxError::UnknownTable` /
//! `DaxError::UnknownColumn`.

use crate::engine::{DaxError, DaxResult, FilterContext};
use crate::pivot::{pivot_crosstab, GroupByColumn, PivotMeasure, PivotResultGrid};
use crate::{DataModel, Value};
use formula_model::pivots::{AggregationType, PivotConfig, PivotField, PivotKeyPart, ValueField};

/// Inputs required by [`pivot_crosstab`] derived from a [`PivotConfig`].
#[derive(Clone, Debug)]
pub struct PivotInputs {
    pub row_fields: Vec<GroupByColumn>,
    pub column_fields: Vec<GroupByColumn>,
    pub measures: Vec<PivotMeasure>,
    pub filter: FilterContext,
}

/// Convert a `formula_model::pivots::PivotConfig` into `formula-dax` pivot inputs.
pub fn pivot_inputs_from_config(
    model: &DataModel,
    base_table: &str,
    cfg: &PivotConfig,
) -> DaxResult<PivotInputs> {
    // Validate the base table early so shorthand parsing errors are reported cleanly.
    model
        .table(base_table)
        .ok_or_else(|| DaxError::UnknownTable(base_table.to_string()))?;

    let row_fields = cfg
        .row_fields
        .iter()
        .map(|f| parse_group_by_field(model, base_table, f))
        .collect::<DaxResult<Vec<_>>>()?;

    let column_fields = cfg
        .column_fields
        .iter()
        .map(|f| parse_group_by_field(model, base_table, f))
        .collect::<DaxResult<Vec<_>>>()?;

    let measures = cfg
        .value_fields
        .iter()
        .map(|f| pivot_measure_from_value_field(model, base_table, f))
        .collect::<DaxResult<Vec<_>>>()?;

    let filter = filter_context_from_config(model, base_table, cfg)?;

    Ok(PivotInputs {
        row_fields,
        column_fields,
        measures,
        filter,
    })
}

/// Convenience helper: execute [`pivot_crosstab`] by converting a [`PivotConfig`].
pub fn pivot_crosstab_from_config(
    model: &DataModel,
    base_table: &str,
    cfg: &PivotConfig,
) -> DaxResult<PivotResultGrid> {
    let inputs = pivot_inputs_from_config(model, base_table, cfg)?;
    pivot_crosstab(
        model,
        base_table,
        &inputs.row_fields,
        &inputs.column_fields,
        &inputs.measures,
        &inputs.filter,
    )
}

fn parse_group_by_field(
    model: &DataModel,
    base_table: &str,
    field: &PivotField,
) -> DaxResult<GroupByColumn> {
    let (table, column) = parse_column_identifier(&field.source_field, base_table)?;
    validate_table_column(model, &table, &column)?;
    Ok(GroupByColumn { table, column })
}

fn filter_context_from_config(
    model: &DataModel,
    base_table: &str,
    cfg: &PivotConfig,
) -> DaxResult<FilterContext> {
    let mut filter = FilterContext::empty();
    for f in &cfg.filter_fields {
        let Some(allowed) = &f.allowed else {
            continue;
        };
        let (table, column) = parse_column_identifier(&f.source_field, base_table)?;
        validate_table_column(model, &table, &column)?;
        filter.set_column_in(&table, &column, allowed.iter().map(pivot_key_part_to_value));
    }
    Ok(filter)
}

fn pivot_measure_from_value_field(
    model: &DataModel,
    base_table: &str,
    field: &ValueField,
) -> DaxResult<PivotMeasure> {
    let source = field.source_field.trim();
    if is_measure_reference(source) {
        // Allow explicit measure references like `[Total Sales]`.
        return PivotMeasure::new(field.name.clone(), source.to_string());
    }

    let (table, column) = parse_column_identifier(source, base_table)?;
    validate_table_column(model, &table, &column)?;

    let expr = match field.aggregation {
        AggregationType::Sum => format!("SUM({table}[{column}])"),
        AggregationType::Count => format!("COUNTA({table}[{column}])"),
        AggregationType::CountNumbers => format!("COUNT({table}[{column}])"),
        AggregationType::Average => format!("AVERAGE({table}[{column}])"),
        AggregationType::Min => format!("MIN({table}[{column}])"),
        AggregationType::Max => format!("MAX({table}[{column}])"),
        other => {
            return Err(DaxError::Eval(format!(
                "unsupported aggregation {other:?} for pivot value field {}",
                field.name
            )))
        }
    };

    PivotMeasure::new(field.name.clone(), expr)
}

fn is_measure_reference(field: &str) -> bool {
    let field = field.trim();
    field.starts_with('[') && field.ends_with(']') && !field.contains("][")
}

fn validate_table_column(model: &DataModel, table: &str, column: &str) -> DaxResult<()> {
    let table_ref = model
        .table(table)
        .ok_or_else(|| DaxError::UnknownTable(table.to_string()))?;
    table_ref
        .column_idx(column)
        .ok_or_else(|| DaxError::UnknownColumn {
            table: table.to_string(),
            column: column.to_string(),
        })?;
    Ok(())
}

fn parse_column_identifier(field: &str, base_table: &str) -> DaxResult<(String, String)> {
    let field = field.trim();
    if field.is_empty() {
        return Err(DaxError::Parse("empty pivot field identifier".to_string()));
    }
    if field.starts_with('[') {
        return Err(DaxError::Parse(format!(
            "expected a column identifier (Table[Column] or Column), got {field}"
        )));
    }

    let Some(open) = field.find('[') else {
        return Ok((base_table.to_string(), field.to_string()));
    };
    if !field.ends_with(']') {
        return Err(DaxError::Parse(format!(
            "invalid pivot field identifier {field}: missing closing ]"
        )));
    }
    let table = field[..open].trim();
    let column = field[open + 1..field.len() - 1].trim();
    if table.is_empty() || column.is_empty() {
        return Err(DaxError::Parse(format!(
            "invalid pivot field identifier {field}: expected Table[Column]"
        )));
    }
    Ok((table.to_string(), column.to_string()))
}

fn pivot_key_part_to_value(part: &PivotKeyPart) -> Value {
    match part {
        PivotKeyPart::Blank => Value::Blank,
        PivotKeyPart::Number(bits) => Value::from(f64::from_bits(*bits)),
        PivotKeyPart::Date(d) => Value::from(d.to_string()),
        PivotKeyPart::Text(s) => Value::from(s.clone()),
        PivotKeyPart::Bool(b) => Value::from(*b),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{Cardinality, CrossFilterDirection, Relationship, Table};
    use formula_model::pivots::{
        FilterField, GrandTotals, Layout, PivotField, PivotKeyPart, SubtotalPosition,
    };
    use pretty_assertions::assert_eq;
    use std::collections::HashSet;

    fn sum_amount_value_field() -> ValueField {
        ValueField {
            source_field: "Amount".to_string(),
            name: "Sum of Amount".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }
    }

    #[test]
    fn base_table_group_by_shorthand_column() {
        let mut model = DataModel::new();
        let mut fact = Table::new("Fact", vec!["Category", "Amount"]);
        fact.push_row(vec![Value::from("A"), Value::from(10.0)])
            .unwrap();
        fact.push_row(vec![Value::from("B"), Value::from(5.0)])
            .unwrap();
        fact.push_row(vec![Value::from("A"), Value::from(2.0)])
            .unwrap();
        model.add_table(fact).unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Category")],
            column_fields: vec![],
            value_fields: vec![sum_amount_value_field()],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals::default(),
        };

        let grid = pivot_crosstab_from_config(&model, "Fact", &cfg).unwrap();
        assert_eq!(
            grid,
            PivotResultGrid {
                data: vec![
                    vec![Value::from("Fact[Category]"), Value::from("Sum of Amount")],
                    vec![Value::from("A"), Value::from(12.0)],
                    vec![Value::from("B"), Value::from(5.0)],
                ],
            }
        );
    }

    #[test]
    fn dimension_table_group_by_with_relationship() {
        let mut model = DataModel::new();
        let mut dim = Table::new("Dim", vec!["Id", "Name"]);
        dim.push_row(vec![Value::from(1_i64), Value::from("Alpha")])
            .unwrap();
        dim.push_row(vec![Value::from(2_i64), Value::from("Beta")])
            .unwrap();
        model.add_table(dim).unwrap();

        let mut fact = Table::new("Fact", vec!["DimId", "Amount"]);
        fact.push_row(vec![Value::from(1_i64), Value::from(10.0)])
            .unwrap();
        fact.push_row(vec![Value::from(2_i64), Value::from(5.0)])
            .unwrap();
        fact.push_row(vec![Value::from(1_i64), Value::from(2.0)])
            .unwrap();
        model.add_table(fact).unwrap();

        model
            .add_relationship(Relationship {
                name: "Fact->Dim".to_string(),
                from_table: "Fact".to_string(),
                from_column: "DimId".to_string(),
                to_table: "Dim".to_string(),
                to_column: "Id".to_string(),
                cardinality: Cardinality::OneToMany,
                cross_filter_direction: CrossFilterDirection::Single,
                is_active: true,
                enforce_referential_integrity: false,
            })
            .unwrap();

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Dim[Name]")],
            column_fields: vec![],
            value_fields: vec![sum_amount_value_field()],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals::default(),
        };

        let grid = pivot_crosstab_from_config(&model, "Fact", &cfg).unwrap();
        assert_eq!(
            grid,
            PivotResultGrid {
                data: vec![
                    vec![Value::from("Dim[Name]"), Value::from("Sum of Amount")],
                    vec![Value::from("Alpha"), Value::from(12.0)],
                    vec![Value::from("Beta"), Value::from(5.0)],
                ],
            }
        );
    }

    #[test]
    fn filter_field_multiple_allowed_values() {
        let mut model = DataModel::new();
        let mut fact = Table::new("Fact", vec!["Category", "Amount"]);
        fact.push_row(vec![Value::from("A"), Value::from(10.0)])
            .unwrap();
        fact.push_row(vec![Value::from("B"), Value::from(5.0)])
            .unwrap();
        fact.push_row(vec![Value::from("C"), Value::from(7.0)])
            .unwrap();
        model.add_table(fact).unwrap();

        let allowed = HashSet::from([
            PivotKeyPart::Text("A".to_string()),
            PivotKeyPart::Text("C".to_string()),
        ]);

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Category")],
            column_fields: vec![],
            value_fields: vec![sum_amount_value_field()],
            filter_fields: vec![FilterField {
                source_field: "Category".to_string(),
                allowed: Some(allowed),
            }],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals::default(),
        };

        let grid = pivot_crosstab_from_config(&model, "Fact", &cfg).unwrap();
        assert_eq!(
            grid,
            PivotResultGrid {
                data: vec![
                    vec![Value::from("Fact[Category]"), Value::from("Sum of Amount")],
                    vec![Value::from("A"), Value::from(10.0)],
                    vec![Value::from("C"), Value::from(7.0)],
                ],
            }
        );
    }
}
