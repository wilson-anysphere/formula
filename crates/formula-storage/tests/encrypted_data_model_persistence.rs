use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, FilterContext, GroupByColumn, PivotMeasure,
    Relationship, Table,
};
use formula_storage::{InMemoryKeyProvider, Storage};
use std::sync::Arc;
use tempfile::tempdir;

#[test]
fn encrypted_data_model_round_trip_columnar() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("workbook.formula");
    let key_provider = Arc::new(InMemoryKeyProvider::default());

    let storage1 = Storage::open_encrypted_path(&path, key_provider.clone()).expect("open storage");
    let workbook = storage1
        .create_workbook("Book", None)
        .expect("create workbook");

    let options = TableOptions {
        page_size_rows: 2,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let dim_schema = vec![
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Category".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut dim_builder = ColumnarTableBuilder::new(dim_schema, options);
    dim_builder.append_row(&[Value::Number(1.0), Value::String(Arc::<str>::from("A"))]);
    dim_builder.append_row(&[Value::Number(2.0), Value::String(Arc::<str>::from("A"))]);
    dim_builder.append_row(&[Value::Number(3.0), Value::String(Arc::<str>::from("B"))]);
    let dim = dim_builder.finalize();

    let fact_schema = vec![
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut fact_builder = ColumnarTableBuilder::new(fact_schema, options);
    fact_builder.append_row(&[Value::Number(1.0), Value::Number(10.0)]);
    fact_builder.append_row(&[Value::Number(1.0), Value::Number(20.0)]);
    fact_builder.append_row(&[Value::Number(2.0), Value::Number(5.0)]);
    fact_builder.append_row(&[Value::Number(3.0), Value::Number(7.0)]);
    let fact = fact_builder.finalize();

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("DimProduct", dim))
        .expect("add dim table");
    model
        .add_table(Table::from_columnar("FactSales", fact))
        .expect("add fact table");
    model
        .add_relationship(Relationship {
            name: "FactSales_Product".to_string(),
            from_table: "FactSales".to_string(),
            from_column: "ProductId".to_string(),
            to_table: "DimProduct".to_string(),
            to_column: "ProductId".to_string(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .expect("add relationship");
    model
        .add_measure("Total Sales", "SUM(FactSales[Amount])")
        .expect("add measure");

    let total = model
        .evaluate_measure("Total Sales", &FilterContext::empty())
        .expect("evaluate measure");
    assert_eq!(total, formula_dax::Value::from(42.0));

    let measures = vec![PivotMeasure::new("Total Sales", "[Total Sales]").expect("pivot measure")];
    let group_by = vec![GroupByColumn::new("DimProduct", "Category")];
    let result = pivot(
        &model,
        "FactSales",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .expect("pivot");
    assert_eq!(
        result.rows,
        vec![
            vec![formula_dax::Value::from("A"), formula_dax::Value::from(35.0)],
            vec![formula_dax::Value::from("B"), formula_dax::Value::from(7.0)],
        ]
    );

    storage1
        .save_data_model(workbook.id, &model)
        .expect("save data model");
    storage1.persist().expect("persist encrypted");
    drop(storage1);

    let on_disk = std::fs::read(&path).expect("read encrypted file");
    assert!(on_disk.starts_with(b"FMLENC01"));

    let storage2 = Storage::open_encrypted_path(&path, key_provider).expect("reopen storage");
    let loaded = storage2
        .load_data_model(workbook.id)
        .expect("load data model");

    let total2 = loaded
        .evaluate_measure("Total Sales", &FilterContext::empty())
        .expect("evaluate after reload");
    assert_eq!(total2, total);
    let result2 = pivot(
        &loaded,
        "FactSales",
        &group_by,
        &measures,
        &FilterContext::empty(),
    )
    .expect("pivot after reload");
    assert_eq!(result2, result);
}

#[test]
fn encrypted_data_model_round_trip_columnar_calculated_columns() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("workbook.formula");
    let key_provider = Arc::new(InMemoryKeyProvider::default());

    let storage1 = Storage::open_encrypted_path(&path, key_provider.clone()).expect("open storage");
    let workbook = storage1
        .create_workbook("Book", None)
        .expect("create workbook");

    let options = TableOptions {
        page_size_rows: 2,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let dim_schema = vec![
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Category".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut dim_builder = ColumnarTableBuilder::new(dim_schema, options);
    dim_builder.append_row(&[Value::Number(1.0), Value::String(Arc::<str>::from("A"))]);
    dim_builder.append_row(&[Value::Number(2.0), Value::String(Arc::<str>::from("A"))]);
    dim_builder.append_row(&[Value::Number(3.0), Value::String(Arc::<str>::from("B"))]);
    let dim = dim_builder.finalize();

    let fact_schema = vec![
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut fact_builder = ColumnarTableBuilder::new(fact_schema, options);
    fact_builder.append_row(&[Value::Number(1.0), Value::Number(10.0)]);
    fact_builder.append_row(&[Value::Number(1.0), Value::Number(20.0)]);
    fact_builder.append_row(&[Value::Number(2.0), Value::Number(5.0)]);
    fact_builder.append_row(&[Value::Number(3.0), Value::Number(7.0)]);
    let fact = fact_builder.finalize();

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("DimProduct", dim))
        .expect("add dim table");
    model
        .add_table(Table::from_columnar("FactSales", fact))
        .expect("add fact table");
    model
        .add_relationship(Relationship {
            name: "FactSales_Product".to_string(),
            from_table: "FactSales".to_string(),
            from_column: "ProductId".to_string(),
            to_table: "DimProduct".to_string(),
            to_column: "ProductId".to_string(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .expect("add relationship");

    model
        .add_calculated_column("FactSales", "Double Amount", "[Amount] * 2")
        .expect("add numeric calc col");
    model
        .add_calculated_column(
            "FactSales",
            "Category From Dim",
            "RELATED(DimProduct[Category])",
        )
        .expect("add related calc col");
    model
        .add_measure("Total Double Amount", "SUM(FactSales[Double Amount])")
        .expect("add measure");

    let total_before = model
        .evaluate_measure("Total Double Amount", &FilterContext::empty())
        .expect("evaluate before save");
    assert_eq!(total_before, formula_dax::Value::from(84.0));

    let pivot_measures = vec![PivotMeasure::new("Total Double Amount", "[Total Double Amount]")
        .expect("pivot measure")];
    let pivot_group_by = vec![GroupByColumn::new("DimProduct", "Category")];
    let pivot_before = pivot(
        &model,
        "FactSales",
        &pivot_group_by,
        &pivot_measures,
        &FilterContext::empty(),
    )
    .expect("pivot before save");
    assert_eq!(
        pivot_before.rows,
        vec![
            vec![
                formula_dax::Value::from("A"),
                formula_dax::Value::from(70.0)
            ],
            vec![
                formula_dax::Value::from("B"),
                formula_dax::Value::from(14.0)
            ],
        ],
        "pivot should see calculated column values before persistence"
    );

    // Also ensure we can group by the *calculated column itself* (not the related dim column).
    // This exercises the columnar group-by fast path using the calculated column's persisted
    // encoded chunks.
    let pivot_group_by_calc = vec![GroupByColumn::new("FactSales", "Category From Dim")];
    let pivot_calc_before = pivot(
        &model,
        "FactSales",
        &pivot_group_by_calc,
        &pivot_measures,
        &FilterContext::empty(),
    )
    .expect("pivot by calculated column before save");
    assert_eq!(
        pivot_calc_before.rows, pivot_before.rows,
        "grouping by the calculated column should match grouping by the related dim column"
    );

    storage1
        .save_data_model(workbook.id, &model)
        .expect("save data model");
    storage1.persist().expect("persist encrypted");
    drop(storage1);

    let on_disk = std::fs::read(&path).expect("read encrypted file");
    assert!(on_disk.starts_with(b"FMLENC01"));

    let storage2 = Storage::open_encrypted_path(&path, key_provider).expect("reopen storage");
    let schema = storage2
        .load_data_model_schema(workbook.id)
        .expect("schema-only load");
    assert_eq!(schema.calculated_columns.len(), 2);
    let fact_schema = schema
        .tables
        .iter()
        .find(|t| t.name == "FactSales")
        .expect("fact table schema");
    assert!(
        fact_schema
            .columns
            .iter()
            .any(|c| c.name == "Double Amount" && c.column_type == ColumnType::Number),
        "expected schema-only load to include Double Amount as a typed table column"
    );
    assert!(
        fact_schema
            .columns
            .iter()
            .any(|c| c.name == "Category From Dim" && c.column_type == ColumnType::String),
        "expected schema-only load to include Category From Dim as a typed table column"
    );
    assert!(
        schema
            .calculated_columns
            .iter()
            .any(|c| c.table == "FactSales" && c.name == "Double Amount" && c.expression == "[Amount] * 2"),
        "expected schema-only load to include Double Amount definition"
    );
    assert!(
        schema
            .calculated_columns
            .iter()
            .any(|c| c.table == "FactSales"
                && c.name == "Category From Dim"
                && c.expression == "RELATED(DimProduct[Category])"),
        "expected schema-only load to include Category From Dim definition"
    );

    let loaded = storage2
        .load_data_model(workbook.id)
        .expect("load data model");
    assert_eq!(loaded.calculated_columns().len(), 2);
    assert!(
        loaded
            .calculated_columns()
            .iter()
            .any(|c| c.table == "FactSales" && c.name == "Double Amount" && c.expression == "[Amount] * 2"),
        "expected Double Amount to be registered after load"
    );
    assert!(
        loaded
            .calculated_columns()
            .iter()
            .any(|c| c.table == "FactSales"
                && c.name == "Category From Dim"
                && c.expression == "RELATED(DimProduct[Category])"),
        "expected Category From Dim to be registered after load"
    );

    let total_after = loaded
        .evaluate_measure("Total Double Amount", &FilterContext::empty())
        .expect("evaluate after reload");
    assert_eq!(total_after, total_before);

    let pivot_after = pivot(
        &loaded,
        "FactSales",
        &pivot_group_by,
        &pivot_measures,
        &FilterContext::empty(),
    )
    .expect("pivot after reload");
    assert_eq!(
        pivot_after, pivot_before,
        "pivot should see the same calculated column values after persistence"
    );
    let pivot_calc_after = pivot(
        &loaded,
        "FactSales",
        &pivot_group_by_calc,
        &pivot_measures,
        &FilterContext::empty(),
    )
    .expect("pivot by calculated column after load");
    assert_eq!(
        pivot_calc_after, pivot_calc_before,
        "pivot grouped by calculated column should round-trip through persistence"
    );

    let fact_after = loaded.table("FactSales").expect("fact table");
    let double_after: Vec<formula_dax::Value> = (0..fact_after.row_count())
        .map(|row| fact_after.value(row, "Double Amount").expect("double amount"))
        .collect();
    assert_eq!(
        double_after,
        vec![
            formula_dax::Value::from(20.0),
            formula_dax::Value::from(40.0),
            formula_dax::Value::from(10.0),
            formula_dax::Value::from(14.0),
        ]
    );
    let category_after: Vec<formula_dax::Value> = (0..fact_after.row_count())
        .map(|row| {
            fact_after
                .value(row, "Category From Dim")
                .expect("category from dim")
        })
        .collect();
    assert_eq!(
        category_after,
        vec![
            formula_dax::Value::from("A"),
            formula_dax::Value::from("A"),
            formula_dax::Value::from("A"),
            formula_dax::Value::from("B"),
        ]
    );

    let col_table = fact_after.columnar_table().expect("columnar backend");
    let idx_double = col_table
        .schema()
        .iter()
        .position(|c| c.name == "Double Amount")
        .expect("double amount column index");
    let idx_category = col_table
        .schema()
        .iter()
        .position(|c| c.name == "Category From Dim")
        .expect("category column index");
    assert!(
        !col_table.encoded_chunks(idx_double).unwrap().is_empty(),
        "expected at least one chunk for Double Amount"
    );
    assert!(
        !col_table.encoded_chunks(idx_category).unwrap().is_empty(),
        "expected at least one chunk for Category From Dim"
    );
    assert_eq!(
        col_table.get_cell(3, idx_category),
        Value::String(Arc::<str>::from("B"))
    );
}
