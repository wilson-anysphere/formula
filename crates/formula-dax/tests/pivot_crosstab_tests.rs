mod common;

use common::build_model;
use formula_dax::{
    pivot_crosstab, Cardinality, CrossFilterDirection, DataModel, FilterContext, GroupByColumn,
    PivotMeasure, Relationship, Table, Value,
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

#[test]
fn pivot_crosstab_resolves_unicode_identifiers_case_insensitively_and_preserves_model_casing() {
    let mut model = DataModel::new();
    let mut fact = Table::new("Straße", vec!["StraßenId", "Region", "Maß"]);
    fact.push_row(vec![1.into(), "East".into(), 10.0.into()])
        .unwrap();
    fact.push_row(vec![1.into(), "East".into(), 20.0.into()])
        .unwrap();
    fact.push_row(vec![2.into(), "West".into(), 5.0.into()])
        .unwrap();
    model.add_table(fact).unwrap();

    // Use ASCII-only identifiers to ensure ß/SS folding is applied for table and column names.
    model.add_measure("Total", "SUM('STRASSE'[MASS])").unwrap();

    let result = pivot_crosstab(
        &model,
        "strasse",
        &[GroupByColumn::new("STRASSE", "STRASSENID")],
        &[GroupByColumn::new("STRASSE", "REGION")],
        &[PivotMeasure::new("Total", "[TOTAL]").unwrap()],
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result.data,
        vec![
            vec![
                Value::from("Straße[StraßenId]"),
                Value::from("East"),
                Value::from("West")
            ],
            vec![1.into(), 30.0.into(), Value::Blank],
            vec![2.into(), Value::Blank, 5.0.into()],
        ]
    );
}

#[test]
fn pivot_crosstab_supports_unicode_related_dimension_row_and_column_fields() {
    let mut model = DataModel::new();

    let mut streets = Table::new("Straße", vec!["StraßenId", "StraßenName", "Region"]);
    streets
        .push_row(vec![1.into(), "A".into(), "East".into()])
        .unwrap();
    streets
        .push_row(vec![2.into(), "B".into(), "West".into()])
        .unwrap();
    streets
        .push_row(vec![3.into(), "C".into(), "East".into()])
        .unwrap();
    model.add_table(streets).unwrap();

    let mut orders = Table::new("Orders", vec!["OrderId", "StraßenId", "Amount"]);
    orders
        .push_row(vec![100.into(), 1.into(), 10.0.into()])
        .unwrap();
    orders
        .push_row(vec![101.into(), 1.into(), 20.0.into()])
        .unwrap();
    orders
        .push_row(vec![102.into(), 2.into(), 5.0.into()])
        .unwrap();
    orders
        .push_row(vec![103.into(), 3.into(), 8.0.into()])
        .unwrap();
    model.add_table(orders).unwrap();

    // Add relationship using mixed casing and ASCII-only identifiers that casefold to the Unicode
    // model names (`Straße` -> `STRASSE`, `StraßenId` -> `STRASSENID`).
    model
        .add_relationship(Relationship {
            name: "Orders->Straße".into(),
            from_table: "orders".into(),
            from_column: "straßenid".into(),
            to_table: "STRASSE".into(),
            to_column: "STRASSENID".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total", "SUM(Orders[Amount])").unwrap();

    let result = pivot_crosstab(
        &model,
        "Orders",
        &[GroupByColumn::new("STRASSE", "REGION")],
        &[GroupByColumn::new("STRASSE", "STRASSENNAME")],
        &[PivotMeasure::new("Total", "[TOTAL]").unwrap()],
        &FilterContext::empty(),
    )
    .unwrap();

    assert_eq!(
        result.data,
        vec![
            vec![
                Value::from("Straße[Region]"),
                Value::from("A"),
                Value::from("B"),
                Value::from("C"),
            ],
            vec![Value::from("East"), 30.0.into(), Value::Blank, 8.0.into()],
            vec![Value::from("West"), Value::Blank, 5.0.into(), Value::Blank],
        ]
    );
}
