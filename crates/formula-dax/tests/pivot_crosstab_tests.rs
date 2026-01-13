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

