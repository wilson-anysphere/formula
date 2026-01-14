use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, FilterContext, GroupByColumn,
    PivotMeasure, Relationship, Table,
};
use formula_storage::data_model::DataModelChunkKind;
use formula_storage::Storage;
use rusqlite::{params, Connection};
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
fn data_model_round_trip_columnar_calculated_columns() {
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

    // Calculated columns on a *columnar* table. Values must be stored physically in the
    // columnar pages and definitions persisted in `data_model_calculated_columns`.
    model
        .add_calculated_column("FactSales", "Double Amount", "[Amount] * 2")
        .expect("add numeric calculated column");
    model
        .add_calculated_column(
            "FactSales",
            "Category From Dim",
            "RELATED(DimProduct[Category])",
        )
        .expect("add related calculated column");

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

    let fact_before = model.table("FactSales").expect("fact table");
    let double_before: Vec<formula_dax::Value> = (0..fact_before.row_count())
        .map(|row| {
            fact_before
                .value(row, "Double Amount")
                .expect("double amount")
        })
        .collect();
    assert_eq!(
        double_before,
        vec![
            formula_dax::Value::from(20.0),
            formula_dax::Value::from(40.0),
            formula_dax::Value::from(10.0),
            formula_dax::Value::from(14.0),
        ]
    );
    let category_before: Vec<formula_dax::Value> = (0..fact_before.row_count())
        .map(|row| {
            fact_before
                .value(row, "Category From Dim")
                .expect("category from dim")
        })
        .collect();
    assert_eq!(
        category_before,
        vec![
            formula_dax::Value::from("A"),
            formula_dax::Value::from("A"),
            formula_dax::Value::from("A"),
            formula_dax::Value::from("B"),
        ]
    );

    storage1
        .save_data_model(workbook.id, &model)
        .expect("save data model");
    drop(storage1);

    // Verify the calculated column definitions were persisted in the dedicated SQLite table (not
    // inferred on load).
    //
    // Use a short-lived standalone connection to avoid holding locks while we reopen via
    // `Storage::open_path` below.
    let workbook_id_str = workbook.id.to_string();
    {
        let conn = Connection::open(path).expect("open sqlite directly");
        let mut calc_stmt = conn
            .prepare(
                r#"
                SELECT table_name, name, expression
                FROM data_model_calculated_columns
                WHERE workbook_id = ?1
                ORDER BY id
                "#,
            )
            .expect("prepare calculated columns query");
        let mut calc_rows: Vec<(String, String, String)> = calc_stmt
            .query_map(params![&workbook_id_str], |row| {
                Ok((
                    row.get::<_, String>(0)?,
                    row.get::<_, String>(1)?,
                    row.get::<_, String>(2)?,
                ))
            })
            .expect("query calculated columns")
            .map(|row| row.expect("row"))
            .collect();
        calc_rows.sort();
        let mut expected_calc_rows = vec![
            (
                "FactSales".to_string(),
                "Double Amount".to_string(),
                "[Amount] * 2".to_string(),
            ),
            (
                "FactSales".to_string(),
                "Category From Dim".to_string(),
                "RELATED(DimProduct[Category])".to_string(),
            ),
        ];
        expected_calc_rows.sort();
        assert_eq!(calc_rows, expected_calc_rows);

        // Verify the calculated column values were persisted as physical column chunks.
        let mut chunk_stmt = conn
            .prepare(
                r#"
                SELECT c.name, COUNT(ch.id)
                FROM data_model_tables t
                JOIN data_model_columns c ON c.table_id = t.id
                JOIN data_model_chunks ch ON ch.column_id = c.id
                WHERE t.workbook_id = ?1
                  AND t.name = 'FactSales'
                  AND c.name IN ('Double Amount', 'Category From Dim')
                GROUP BY c.name
                ORDER BY c.name
                "#,
            )
            .expect("prepare chunk query");
        let chunk_counts: Vec<(String, i64)> = chunk_stmt
            .query_map(params![&workbook_id_str], |row| {
                Ok((row.get::<_, String>(0)?, row.get::<_, i64>(1)?))
            })
            .expect("query chunk counts")
            .map(|row| row.expect("row"))
            .collect();
        assert_eq!(chunk_counts.len(), 2, "expected chunk rows for both columns");
        for (name, count) in chunk_counts {
            assert!(count > 0, "expected at least one persisted chunk for {name}");
        }
    }

    let storage2 = Storage::open_path(path).expect("reopen storage");
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
            .any(|c| c.table == "FactSales"
                && c.name == "Double Amount"
                && c.expression == "[Amount] * 2"),
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

    // Ensure the chunk streaming API can read calculated columns (used by incremental / lazy loads).
    let mut streamed_double = Vec::new();
    storage2
        .stream_data_model_column_chunks(workbook.id, "FactSales", "Double Amount", |chunk| {
            streamed_double.push((chunk.chunk_index, chunk.kind, chunk.data.len()));
            Ok(())
        })
        .expect("stream Double Amount chunks");
    assert_eq!(streamed_double.len(), 2);
    assert_eq!(
        streamed_double
            .iter()
            .map(|(idx, _, _)| *idx)
            .collect::<Vec<_>>(),
        vec![0, 1]
    );
    assert!(
        streamed_double
            .iter()
            .all(|(_, kind, size)| *kind == DataModelChunkKind::Float && *size > 0),
        "expected Double Amount chunks to be non-empty float chunks"
    );

    let mut streamed_category = Vec::new();
    storage2
        .stream_data_model_column_chunks(workbook.id, "FactSales", "Category From Dim", |chunk| {
            streamed_category.push((chunk.chunk_index, chunk.kind, chunk.data.len()));
            Ok(())
        })
        .expect("stream Category From Dim chunks");
    assert_eq!(streamed_category.len(), 2);
    assert_eq!(
        streamed_category
            .iter()
            .map(|(idx, _, _)| *idx)
            .collect::<Vec<_>>(),
        vec![0, 1]
    );
    assert!(
        streamed_category
            .iter()
            .all(|(_, kind, size)| *kind == DataModelChunkKind::Dict && *size > 0),
        "expected Category From Dim chunks to be non-empty dict chunks"
    );

    let loaded = storage2
        .load_data_model(workbook.id)
        .expect("load data model");
    assert_eq!(loaded.calculated_columns().len(), 2);
    assert!(
        loaded
            .calculated_columns()
            .iter()
            .any(|c| c.table == "FactSales"
                && c.name == "Double Amount"
                && c.expression == "[Amount] * 2"),
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
        .expect("evaluate after load");
    assert_eq!(total_after, total_before);

    let pivot_after = pivot(
        &loaded,
        "FactSales",
        &pivot_group_by,
        &pivot_measures,
        &FilterContext::empty(),
    )
    .expect("pivot after load");
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

    let fact_after = loaded.table("FactSales").expect("loaded fact");
    assert!(
        fact_after.columns().iter().any(|c| c == "Double Amount"),
        "expected Double Amount column to be persisted in the table"
    );
    assert!(
        fact_after
            .columns()
            .iter()
            .any(|c| c == "Category From Dim"),
        "expected Category From Dim column to be persisted in the table"
    );

    let double_after: Vec<formula_dax::Value> = (0..fact_after.row_count())
        .map(|row| {
            fact_after
                .value(row, "Double Amount")
                .expect("double amount")
        })
        .collect();
    assert_eq!(double_after, double_before);

    let category_after: Vec<formula_dax::Value> = (0..fact_after.row_count())
        .map(|row| {
            fact_after
                .value(row, "Category From Dim")
                .expect("category from dim")
        })
        .collect();
    assert_eq!(category_after, category_before);

    // Ensure the values are physically stored in the columnar chunks.
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
