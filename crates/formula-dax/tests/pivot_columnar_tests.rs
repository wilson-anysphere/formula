use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{pivot, DataModel, FilterContext, GroupByColumn, PivotMeasure, Table, Value};
use pretty_assertions::assert_eq;
use std::sync::Arc;

fn build_models(rows: usize) -> (DataModel, DataModel) {
    let groups = ["A", "B", "C", "D"];

    let mut vec_model = DataModel::new();
    let mut vec_fact = Table::new("Fact", vec!["Group", "Amount", "Maybe"]);
    for i in 0..rows {
        let g = groups[i % groups.len()];
        let amount = (i % 100) as f64;
        let maybe = if i % 3 == 0 { Value::Blank } else { Value::from(amount) };
        vec_fact
            .push_row(vec![Value::from(g), Value::from(amount), maybe])
            .unwrap();
    }
    vec_model.add_table(vec_fact).unwrap();
    vec_model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    vec_model.add_measure("Double", "[Total] * 2").unwrap();
    vec_model
        .add_measure("Big Total", "IF([Total] > 20000, [Total], BLANK())")
        .unwrap();

    let schema = vec![
        ColumnSchema {
            name: "Group".to_string(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Maybe".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 1024,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut builder = ColumnarTableBuilder::new(schema, options);
    for i in 0..rows {
        let g = groups[i % groups.len()];
        let amount = (i % 100) as f64;
        let maybe = if i % 3 == 0 {
            formula_columnar::Value::Null
        } else {
            formula_columnar::Value::Number(amount)
        };
        builder.append_row(&[
            formula_columnar::Value::String(Arc::<str>::from(g)),
            formula_columnar::Value::Number(amount),
            maybe,
        ]);
    }

    let mut col_model = DataModel::new();
    col_model
        .add_table(Table::from_columnar("Fact", builder.finalize()))
        .unwrap();
    col_model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    col_model.add_measure("Double", "[Total] * 2").unwrap();
    col_model
        .add_measure("Big Total", "IF([Total] > 20000, [Total], BLANK())")
        .unwrap();

    (vec_model, col_model)
}

#[test]
fn pivot_matches_between_vec_and_columnar_backends() {
    let (vec_model, col_model) = build_models(10_000);

    let group_by = vec![GroupByColumn::new("Fact", "Group")];
    let measures = vec![
        PivotMeasure::new("Total", "[Total]").unwrap(),
        PivotMeasure::new("Double", "[Double]").unwrap(),
        PivotMeasure::new("Big Total", "[Big Total]").unwrap(),
        PivotMeasure::new("Rows", "COUNTROWS(Fact)").unwrap(),
        PivotMeasure::new("Count", "COUNT(Fact[Maybe])").unwrap(),
        PivotMeasure::new("CountA", "COUNTA(Fact[Maybe])").unwrap(),
        PivotMeasure::new("Blank Count", "COUNTBLANK(Fact[Maybe])").unwrap(),
        PivotMeasure::new("Avg", "AVERAGE(Fact[Amount])").unwrap(),
        PivotMeasure::new("Distinct Amount", "DISTINCTCOUNT(Fact[Amount])").unwrap(),
    ];

    let vec_result = pivot(
        &vec_model,
        "Fact",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    let col_result = pivot(
        &col_model,
        "Fact",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(vec_result, col_result);

    let a_filter = FilterContext::empty().with_column_equals("Fact", "Group", "A".into());
    let vec_result = pivot(&vec_model, "Fact", &group_by, &measures, &a_filter).unwrap();
    let col_result = pivot(&col_model, "Fact", &group_by, &measures, &a_filter).unwrap();
    assert_eq!(vec_result, col_result);

    let amount_filter = FilterContext::empty().with_column_equals("Fact", "Amount", 42.0.into());
    let vec_result = pivot(&vec_model, "Fact", &group_by, &measures, &amount_filter).unwrap();
    let col_result = pivot(&col_model, "Fact", &group_by, &measures, &amount_filter).unwrap();
    assert_eq!(vec_result, col_result);
}
