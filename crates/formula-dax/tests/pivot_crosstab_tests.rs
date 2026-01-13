mod common;

use common::build_model;
use formula_dax::{
    pivot_crosstab, DataModel, FilterContext, GroupByColumn, PivotMeasure, Table, Value,
};
use pretty_assertions::assert_eq;

fn build_fact_model() -> DataModel {
    let mut model = DataModel::new();
    let mut fact = Table::new("Fact", vec!["Region", "Product", "Amount"]);
    fact.push_row(vec!["East".into(), "A".into(), 100.0.into()])
        .unwrap();
    fact.push_row(vec!["East".into(), "B".into(), 150.0.into()])
        .unwrap();
    fact.push_row(vec!["West".into(), "A".into(), 200.0.into()])
        .unwrap();
    fact.push_row(vec!["West".into(), "B".into(), 250.0.into()])
        .unwrap();
    model.add_table(fact).unwrap();
    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    model.add_measure("Double", "[Total] * 2").unwrap();
    model
}

#[test]
fn pivot_crosstab_single_row_single_col_single_measure() {
    let model = build_fact_model();

    let result = pivot_crosstab(
        &model,
        "Fact",
        &[GroupByColumn::new("Fact", "Region")],
        &[GroupByColumn::new("Fact", "Product")],
        &[PivotMeasure::new("Total", "[Total]").unwrap()],
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result.data,
        vec![
            vec![Value::from("Fact[Region]"), Value::from("A"), Value::from("B")],
            vec![Value::from("East"), 100.0.into(), 150.0.into()],
            vec![Value::from("West"), 200.0.into(), 250.0.into()],
        ]
    );
}

#[test]
fn pivot_crosstab_supports_multiple_measures() {
    let model = build_fact_model();

    let result = pivot_crosstab(
        &model,
        "Fact",
        &[GroupByColumn::new("Fact", "Region")],
        &[GroupByColumn::new("Fact", "Product")],
        &[
            PivotMeasure::new("Total", "[Total]").unwrap(),
            PivotMeasure::new("Double", "[Double]").unwrap(),
        ],
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result.data,
        vec![
            vec![
                Value::from("Fact[Region]"),
                Value::from("A - Total"),
                Value::from("A - Double"),
                Value::from("B - Total"),
                Value::from("B - Double"),
            ],
            vec![
                Value::from("East"),
                100.0.into(),
                200.0.into(),
                150.0.into(),
                300.0.into(),
            ],
            vec![
                Value::from("West"),
                200.0.into(),
                400.0.into(),
                250.0.into(),
                500.0.into(),
            ],
        ]
    );
}

#[test]
fn pivot_crosstab_with_no_column_fields_behaves_like_grouped_table() {
    let model = build_fact_model();

    let result = pivot_crosstab(
        &model,
        "Fact",
        &[GroupByColumn::new("Fact", "Region")],
        &[],
        &[
            PivotMeasure::new("Total", "[Total]").unwrap(),
            PivotMeasure::new("Double", "[Double]").unwrap(),
        ],
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result.data,
        vec![
            vec![Value::from("Fact[Region]"), Value::from("Total"), Value::from("Double")],
            vec![Value::from("East"), 250.0.into(), 500.0.into()],
            vec![Value::from("West"), 450.0.into(), 900.0.into()],
        ]
    );
}

#[test]
fn pivot_crosstab_renders_blank_for_missing_row_column_combinations() {
    let mut model = DataModel::new();
    let mut fact = Table::new("Fact", vec!["Region", "Product", "Amount"]);
    fact.push_row(vec!["East".into(), "A".into(), 10.0.into()])
        .unwrap();
    fact.push_row(vec!["East".into(), "B".into(), 5.0.into()])
        .unwrap();
    fact.push_row(vec!["West".into(), "A".into(), 7.0.into()])
        .unwrap();
    // Note: West/B is intentionally missing.
    model.add_table(fact).unwrap();
    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

    let result = pivot_crosstab(
        &model,
        "Fact",
        &[GroupByColumn::new("Fact", "Region")],
        &[GroupByColumn::new("Fact", "Product")],
        &[PivotMeasure::new("Total", "[Total]").unwrap()],
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result.data,
        vec![
            vec![Value::from("Fact[Region]"), Value::from("A"), Value::from("B")],
            vec![Value::from("East"), 10.0.into(), 5.0.into()],
            vec![Value::from("West"), 7.0.into(), Value::Blank],
        ]
    );
}

#[test]
fn pivot_crosstab_multiple_column_fields_join_with_slash() {
    let mut model = DataModel::new();
    let mut fact = Table::new("Fact", vec!["Region", "Year", "Quarter", "Amount"]);
    fact.push_row(vec!["East".into(), 2024.0.into(), "Q1".into(), 10.0.into()])
        .unwrap();
    fact.push_row(vec!["East".into(), 2024.0.into(), "Q2".into(), 20.0.into()])
        .unwrap();
    fact.push_row(vec!["West".into(), 2024.0.into(), "Q1".into(), 7.0.into()])
        .unwrap();
    // Note: West/Q2 intentionally missing.
    model.add_table(fact).unwrap();
    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

    let result = pivot_crosstab(
        &model,
        "Fact",
        &[GroupByColumn::new("Fact", "Region")],
        &[
            GroupByColumn::new("Fact", "Year"),
            GroupByColumn::new("Fact", "Quarter"),
        ],
        &[PivotMeasure::new("Total", "[Total]").unwrap()],
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result.data,
        vec![
            vec![
                Value::from("Fact[Region]"),
                Value::from("2024 / Q1"),
                Value::from("2024 / Q2"),
            ],
            vec![Value::from("East"), 10.0.into(), 20.0.into()],
            vec![Value::from("West"), 7.0.into(), Value::Blank],
        ]
    );
}

#[test]
fn pivot_crosstab_supports_related_dimension_row_and_column_fields() {
    let mut model = build_model();
    model.add_measure("Total", "SUM(Orders[Amount])").unwrap();

    let result = pivot_crosstab(
        &model,
        "Orders",
        &[GroupByColumn::new("Customers", "Region")],
        &[GroupByColumn::new("Customers", "Name")],
        &[PivotMeasure::new("Total", "[Total]").unwrap()],
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result.data,
        vec![
            vec![
                Value::from("Customers[Region]"),
                Value::from("Alice"),
                Value::from("Bob"),
                Value::from("Carol"),
            ],
            vec![Value::from("East"), 30.0.into(), Value::Blank, 8.0.into()],
            vec![Value::from("West"), Value::Blank, 5.0.into(), Value::Blank],
        ]
    );
}
