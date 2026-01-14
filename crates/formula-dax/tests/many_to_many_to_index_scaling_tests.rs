use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, DaxEngine, FilterContext, GroupByColumn,
    PivotMeasure, Relationship, RowContext, Table, Value,
};
use pretty_assertions::assert_eq;
use std::sync::Arc;

fn options() -> TableOptions {
    TableOptions {
        page_size_rows: 64,
        cache: PageCacheConfig { max_entries: 2 },
    }
}

#[test]
fn many_to_many_columnar_to_table_with_duplicates_is_correct() {
    // This test exercises the scalable many-to-many `to_table` lookup path:
    // - `to_table` is columnar
    // - `to_table` contains duplicate keys
    // - `from_table` is in-memory (so we still build `from_index`)
    //
    // The implementation should not need to eagerly materialize `Vec<usize>` per key on the
    // `to_table` side, but semantics must remain unchanged.
    let mut model = DataModel::new();

    // In-memory fact table.
    let mut orders = Table::new("Orders", vec!["OrderId", "ProductId", "Amount"]);
    orders.push_row(vec![1.into(), 1.into(), 10.0.into()]).unwrap();
    orders.push_row(vec![2.into(), 2.into(), 20.0.into()]).unwrap();
    model.add_table(orders).unwrap();

    // Columnar dimension table with duplicate keys.
    let schema = vec![
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Category".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut products = ColumnarTableBuilder::new(schema, options());
    products.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("A")),
    ]);
    products.append_row(&[
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::String(Arc::<str>::from("B")),
    ]);
    products.append_row(&[
        formula_columnar::Value::Number(2.0),
        formula_columnar::Value::String(Arc::<str>::from("C")),
    ]);
    model
        .add_table(Table::from_columnar("Products", products.finalize()))
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "Orders_Products".into(),
            from_table: "Orders".into(),
            from_column: "ProductId".into(),
            to_table: "Products".into(),
            to_column: "ProductId".into(),
            cardinality: Cardinality::ManyToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_measure("Total Amount", "SUM(Orders[Amount])")
        .unwrap();

    // Filter propagation should use key-set semantics: both A and B map to ProductId=1 and should
    // return the same fact rows.
    let filter_a =
        FilterContext::empty().with_column_equals("Products", "Category", Value::from("A"));
    let filter_b =
        FilterContext::empty().with_column_equals("Products", "Category", Value::from("B"));
    let filter_c =
        FilterContext::empty().with_column_equals("Products", "Category", Value::from("C"));

    assert_eq!(
        model.evaluate_measure("Total Amount", &filter_a).unwrap(),
        10.0.into()
    );
    assert_eq!(
        model.evaluate_measure("Total Amount", &filter_b).unwrap(),
        10.0.into()
    );
    assert_eq!(
        model.evaluate_measure("Total Amount", &filter_c).unwrap(),
        20.0.into()
    );

    // Group expansion must include all reachable related rows/values.
    let engine = DaxEngine::new();
    assert_eq!(
        engine
            .evaluate(
                &model,
                "COUNTROWS(SUMMARIZE(Orders, Products[Category]))",
                &FilterContext::empty(),
                &RowContext::default(),
            )
            .unwrap(),
        Value::from(3_i64)
    );

    let measures = vec![PivotMeasure::new("Total Amount", "SUM(Orders[Amount])").unwrap()];
    let group_by = vec![GroupByColumn::new("Products", "Category")];
    let pivoted = pivot(&model, "Orders", &group_by, &measures, &FilterContext::empty()).unwrap();
    assert_eq!(
        pivoted.rows,
        vec![
            vec![Value::from("A"), 10.0.into()],
            vec![Value::from("B"), 10.0.into()],
            vec![Value::from("C"), 20.0.into()],
        ]
    );
}

