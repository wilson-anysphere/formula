use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{pivot, DataModel, FilterContext, GroupByColumn, PivotMeasure, Table, Value};
use pretty_assertions::assert_eq;
use std::sync::Arc;

fn build_currency_models() -> (DataModel, DataModel) {
    let mut vec_model = DataModel::new();
    let mut vec_fact = Table::new("Fact", vec!["Group", "Amount"]);
    vec_fact
        .push_row(vec![Value::from("A"), Value::from(12.34)])
        .unwrap();
    vec_fact
        .push_row(vec![Value::from("B"), Value::from(2.0)])
        .unwrap();
    vec_model.add_table(vec_fact).unwrap();
    vec_model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    vec_model.add_measure("Min", "MIN(Fact[Amount])").unwrap();
    vec_model.add_measure("Max", "MAX(Fact[Amount])").unwrap();
    vec_model.add_measure("Avg", "AVERAGE(Fact[Amount])").unwrap();
    vec_model
        .add_measure("Distinct", "DISTINCTCOUNT(Fact[Amount])")
        .unwrap();

    let schema = vec![
        ColumnSchema {
            name: "Group".to_string(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Currency { scale: 2 },
        },
    ];
    let options = TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut builder = ColumnarTableBuilder::new(schema, options);
    builder.append_row(&[
        formula_columnar::Value::String(Arc::<str>::from("A")),
        formula_columnar::Value::Currency(1234),
    ]);
    builder.append_row(&[
        formula_columnar::Value::String(Arc::<str>::from("B")),
        formula_columnar::Value::Currency(200),
    ]);

    let mut col_model = DataModel::new();
    col_model
        .add_table(Table::from_columnar("Fact", builder.finalize()))
        .unwrap();
    col_model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    col_model.add_measure("Min", "MIN(Fact[Amount])").unwrap();
    col_model.add_measure("Max", "MAX(Fact[Amount])").unwrap();
    col_model.add_measure("Avg", "AVERAGE(Fact[Amount])").unwrap();
    col_model
        .add_measure("Distinct", "DISTINCTCOUNT(Fact[Amount])")
        .unwrap();

    (vec_model, col_model)
}

#[test]
fn columnar_currency_scale_respected_for_measures() {
    let (_vec_model, mut col_model) = build_currency_models();
    col_model
        .add_measure("Selected", "SELECTEDVALUE(Fact[Amount])")
        .unwrap();

    assert_eq!(
        col_model
            .evaluate_measure("Total", &FilterContext::empty())
            .unwrap(),
        14.34.into()
    );
    assert_eq!(
        col_model.evaluate_measure("Min", &FilterContext::empty()).unwrap(),
        2.0.into()
    );
    assert_eq!(
        col_model.evaluate_measure("Avg", &FilterContext::empty()).unwrap(),
        7.17.into()
    );
    assert_eq!(
        col_model
            .evaluate_measure("Distinct", &FilterContext::empty())
            .unwrap(),
        2.into()
    );

    // Non-empty filter forces row scanning and exercises `value_by_idx` conversion.
    let a_filter = FilterContext::empty().with_column_equals("Fact", "Group", "A".into());
    assert_eq!(col_model.evaluate_measure("Total", &a_filter).unwrap(), 12.34.into());

    // Filtering on the currency column itself exercises `filter_in` + typed conversion.
    let amount_filter = FilterContext::empty().with_column_equals("Fact", "Amount", 2.0.into());
    assert_eq!(
        col_model.evaluate_measure("Total", &amount_filter).unwrap(),
        2.0.into()
    );
    assert_eq!(
        col_model.evaluate_measure("Selected", &amount_filter).unwrap(),
        2.0.into()
    );
    assert_eq!(
        col_model
            .evaluate_measure("Selected", &FilterContext::empty())
            .unwrap(),
        Value::Blank
    );
}

#[test]
fn columnar_currency_scale_respected_in_pivot_group_by() {
    let (vec_model, col_model) = build_currency_models();

    let group_by = vec![GroupByColumn::new("Fact", "Group")];
    let measures = vec![
        PivotMeasure::new("Total", "[Total]").unwrap(),
        PivotMeasure::new("Min", "[Min]").unwrap(),
        PivotMeasure::new("Max", "[Max]").unwrap(),
        PivotMeasure::new("Avg", "[Avg]").unwrap(),
        PivotMeasure::new("Distinct", "[Distinct]").unwrap(),
    ];

    assert_eq!(
        pivot(&vec_model, "Fact", &group_by, &measures, &FilterContext::empty()).unwrap(),
        pivot(&col_model, "Fact", &group_by, &measures, &FilterContext::empty()).unwrap(),
    );

    // Also verify the optimized path under a filter that selects a single currency value.
    let amount_filter = FilterContext::empty().with_column_equals("Fact", "Amount", 12.34.into());
    assert_eq!(
        pivot(&vec_model, "Fact", &group_by, &measures, &amount_filter).unwrap(),
        pivot(&col_model, "Fact", &group_by, &measures, &amount_filter).unwrap(),
    );
}

