use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, FilterContext, GroupByColumn,
    PivotMeasure, Relationship, Table,
};
use formula_storage::Storage;
use std::sync::Arc;
use std::time::Instant;

#[test]
fn data_model_round_trip_columnar() {
    let tmp = tempfile::NamedTempFile::new().expect("tmpfile");
    let path = tmp.path();

    let storage1 = Storage::open_path(path).expect("open storage");
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
            vec![
                formula_dax::Value::from("A"),
                formula_dax::Value::from(35.0)
            ],
            vec![formula_dax::Value::from("B"), formula_dax::Value::from(7.0)],
        ]
    );

    storage1
        .save_data_model(workbook.id, &model)
        .expect("save data model");
    drop(storage1);

    let storage2 = Storage::open_path(path).expect("reopen storage");
    let schema = storage2
        .load_data_model_schema(workbook.id)
        .expect("schema-only load");
    assert_eq!(schema.tables.len(), 2);

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
fn schema_only_load_is_faster_than_full_load() {
    // Use a shared in-memory db so we can simulate reopen without disk IO.
    let uri = "file:data_model_perf?mode=memory&cache=shared";
    let storage1 = Storage::open_uri(uri).expect("open storage");
    let workbook = storage1
        .create_workbook("Book", None)
        .expect("create workbook");

    let rows = 200_000usize;
    let options = TableOptions {
        page_size_rows: 1024,
        cache: PageCacheConfig { max_entries: 8 },
    };
    let schema = vec![
        ColumnSchema {
            name: "Category".to_string(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut builder = ColumnarTableBuilder::new(schema, options);
    let cats = [
        Arc::<str>::from("A"),
        Arc::<str>::from("B"),
        Arc::<str>::from("C"),
        Arc::<str>::from("D"),
    ];
    for i in 0..rows {
        builder.append_row(&[
            Value::String(cats[i % cats.len()].clone()),
            Value::Number((i % 100) as f64),
        ]);
    }
    let table = builder.finalize();

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("Fact", table))
        .expect("add table");
    model
        .add_measure("Total", "SUM(Fact[Amount])")
        .expect("add measure");

    storage1
        .save_data_model(workbook.id, &model)
        .expect("save model");

    let storage2 = Storage::open_uri(uri).expect("open second handle");

    let start = Instant::now();
    let _schema = storage2
        .load_data_model_schema(workbook.id)
        .expect("load schema");
    let schema_dur = start.elapsed();

    let start = Instant::now();
    let _full = storage2.load_data_model(workbook.id).expect("load full");
    let full_dur = start.elapsed();

    eprintln!(
        "schema-only load: {:?}, full load: {:?} (rows={rows})",
        schema_dur, full_dur
    );

    assert!(
        full_dur > schema_dur,
        "expected full load to take longer than schema-only load"
    );

    // Keep the first handle alive so the shared memory DB isn't dropped.
    drop(storage1);
}
