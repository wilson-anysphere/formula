use criterion::{black_box, criterion_group, criterion_main, BenchmarkId, Criterion};
use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, FilterContext, GroupByColumn, PivotMeasure,
    Relationship, Table,
};
use std::sync::Arc;
use std::time::Duration;

fn bench_rows() -> usize {
    std::env::var("FORMULA_DAX_PIVOT_BENCH_ROWS")
        .ok()
        .and_then(|v| v.replace('_', "").parse::<usize>().ok())
        .filter(|&v| v >= 1_000_000 && v <= 5_000_000)
        .unwrap_or(1_000_000)
}

fn build_star_schema_model(rows: usize) -> DataModel {
    // Cardinalities picked to create a non-trivial relationship index, but keep the number
    // of pivot groups small (region x category).
    let customers = 50_000usize;
    let products = 10_000usize;
    let regions = 10usize;
    let categories = 20usize;

    let options = TableOptions {
        page_size_rows: 65_536,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let region_values: Vec<Arc<str>> = (0..regions)
        .map(|i| Arc::<str>::from(format!("Region_{i:02}")))
        .collect();
    let category_values: Vec<Arc<str>> = (0..categories)
        .map(|i| Arc::<str>::from(format!("Category_{i:02}")))
        .collect();

    let mut model = DataModel::new();

    // Regions (snowflake dimension used via Customers -> Regions).
    let regions_schema = vec![
        ColumnSchema {
            name: "RegionId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Name".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut regions_builder = ColumnarTableBuilder::new(regions_schema, options);
    for region_id in 0..regions {
        regions_builder.append_row(&[
            formula_columnar::Value::Number(region_id as f64),
            formula_columnar::Value::String(region_values[region_id].clone()),
        ]);
    }
    model
        .add_table(Table::from_columnar("Regions", regions_builder.finalize()))
        .unwrap();

    // Customers dimension.
    let customers_schema = vec![
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "RegionId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Region".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut customers_builder = ColumnarTableBuilder::new(customers_schema, options);
    for customer_id in 0..customers {
        let region_id = customer_id % regions;
        let region = region_values[region_id].clone();
        customers_builder.append_row(&[
            formula_columnar::Value::Number(customer_id as f64),
            formula_columnar::Value::Number(region_id as f64),
            formula_columnar::Value::String(region),
        ]);
    }
    model
        .add_table(Table::from_columnar("Customers", customers_builder.finalize()))
        .unwrap();

    // Products dimension.
    let products_schema = vec![
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Category".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut products_builder = ColumnarTableBuilder::new(products_schema, options);
    for product_id in 0..products {
        let category = category_values[product_id % categories].clone();
        products_builder.append_row(&[
            formula_columnar::Value::Number(product_id as f64),
            formula_columnar::Value::String(category),
        ]);
    }
    model
        .add_table(Table::from_columnar("Products", products_builder.finalize()))
        .unwrap();

    // Sales fact table.
    //
    // This table intentionally includes both:
    // - foreign keys (`CustomerId`, `ProductId`) used for star-schema pivots
    // - denormalized attributes (`Region`, `Category`) used for the base-table columnar group-by
    let sales_schema = vec![
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "ProductId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Region".to_string(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "Category".to_string(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Quantity".to_string(),
            column_type: ColumnType::Number,
        },
    ];

    let mut sales_builder = ColumnarTableBuilder::new(sales_schema, options);
    for i in 0..rows {
        let customer_id = i % customers;
        // Mix the product id to avoid perfectly-aligned stripes across dimensions.
        let product_id = (i.wrapping_mul(13)) % products;

        let region = region_values[customer_id % regions].clone();
        let category = category_values[product_id % categories].clone();

        let amount = (i % 100) as f64;
        let quantity = (i % 7 + 1) as f64;

        sales_builder.append_row(&[
            formula_columnar::Value::Number(customer_id as f64),
            formula_columnar::Value::Number(product_id as f64),
            formula_columnar::Value::String(region),
            formula_columnar::Value::String(category),
            formula_columnar::Value::Number(amount),
            formula_columnar::Value::Number(quantity),
        ]);
    }

    model
        .add_table(Table::from_columnar("Sales", sales_builder.finalize()))
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "Sales_Customers".into(),
            from_table: "Sales".into(),
            from_column: "CustomerId".into(),
            to_table: "Customers".into(),
            to_column: "CustomerId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "Customers_Regions".into(),
            from_table: "Customers".into(),
            from_column: "RegionId".into(),
            to_table: "Regions".into(),
            to_column: "RegionId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "Sales_Products".into(),
            from_table: "Sales".into(),
            from_column: "ProductId".into(),
            to_table: "Products".into(),
            to_column: "ProductId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total Sales", "SUM(Sales[Amount])").unwrap();
    model.add_measure("Total Qty", "SUM(Sales[Quantity])").unwrap();
    model
        .add_measure("Avg Price", "DIVIDE([Total Sales], [Total Qty])")
        .unwrap();
    model
        .add_measure(
            "Adjusted Sales",
            "([Total Sales] * 1.07) + ([Total Qty] * 0.10)",
        )
        .unwrap();
    // A deliberately unplannable measure (the pivot planner does not support IF) that forces the
    // star-schema pivot to fall back to `pivot_row_scan`.
    model
        .add_measure(
            "Total Sales If",
            "IF(TRUE(), [Total Sales], BLANK())",
        )
        .unwrap();

    model
}

fn bench_pivot_star_schema(c: &mut Criterion) {
    // Optional pivot path tracing.
    //
    // Set `FORMULA_DAX_PIVOT_TRACE=1` when running the bench to print the chosen execution paths
    // once per process (e.g. `columnar_group_by`, `columnar_star_schema_group_by`,
    // `planned_row_group_by`, `row_scan`).
    std::env::remove_var("FORMULA_DAX_PIVOT_DISABLE_STAR_SCHEMA");

    let rows = bench_rows();
    let model = build_star_schema_model(rows);

    let measures = vec![
        PivotMeasure::new("Total Sales", "[Total Sales]").unwrap(),
        PivotMeasure::new("Total Qty", "[Total Qty]").unwrap(),
        PivotMeasure::new("Avg Price", "[Avg Price]").unwrap(),
        PivotMeasure::new("Adjusted Sales", "[Adjusted Sales]").unwrap(),
    ];
    let row_scan_measures = vec![PivotMeasure::new("Total Sales If", "[Total Sales If]").unwrap()];

    let base_group_by = vec![
        GroupByColumn::new("Sales", "Region"),
        GroupByColumn::new("Sales", "Category"),
    ];
    let star_group_by = vec![
        GroupByColumn::new("Customers", "Region"),
        GroupByColumn::new("Products", "Category"),
    ];
    let snowflake_group_by = vec![
        GroupByColumn::new("Regions", "Name"),
        GroupByColumn::new("Products", "Category"),
    ];
    let base_region_group_by = vec![GroupByColumn::new("Sales", "Region")];
    let star_region_group_by = vec![GroupByColumn::new("Customers", "Region")];

    // Sanity check: the denormalized and star-schema pivots should match.
    //
    // Note: the first two column headers differ (`Sales[...]` vs `Customers[...]` / `Products[...]`),
    // but the grouped rows (keys + measures) should be identical.
    let base_result =
        pivot(&model, "Sales", &base_group_by, &measures, &FilterContext::empty()).unwrap();
    let star_result =
        pivot(&model, "Sales", &star_group_by, &measures, &FilterContext::empty()).unwrap();
    assert_eq!(base_result.rows, star_result.rows);
    assert_eq!(base_result.columns.len(), star_result.columns.len());
    assert_eq!(
        &base_result.columns[base_group_by.len()..],
        &star_result.columns[star_group_by.len()..]
    );

    let snowflake_result =
        pivot(&model, "Sales", &snowflake_group_by, &measures, &FilterContext::empty()).unwrap();
    assert_eq!(base_result.rows, snowflake_result.rows);
    assert_eq!(base_result.columns.len(), snowflake_result.columns.len());
    assert_eq!(
        &base_result.columns[base_group_by.len()..],
        &snowflake_result.columns[snowflake_group_by.len()..]
    );

    // Sanity check: a star-schema pivot that falls back to row-scan should still match the
    // denormalized version (which uses the columnar group key fast path + per-group measure eval).
    let base_if_result = pivot(
        &model,
        "Sales",
        &base_region_group_by,
        &row_scan_measures,
        &FilterContext::empty(),
    )
    .unwrap();
    let star_if_result = pivot(
        &model,
        "Sales",
        &star_region_group_by,
        &row_scan_measures,
        &FilterContext::empty(),
    )
    .unwrap();
    assert_eq!(base_if_result.rows, star_if_result.rows);
    assert_eq!(base_if_result.columns.len(), star_if_result.columns.len());
    assert_eq!(
        &base_if_result.columns[base_region_group_by.len()..],
        &star_if_result.columns[star_region_group_by.len()..]
    );

    // Sanity check: disabling the star-schema fast path should still produce identical rows, but
    // use a slower planned row-group-by fallback.
    std::env::set_var("FORMULA_DAX_PIVOT_DISABLE_STAR_SCHEMA", "1");
    let star_result_row_scan =
        pivot(&model, "Sales", &star_group_by, &measures, &FilterContext::empty()).unwrap();
    std::env::remove_var("FORMULA_DAX_PIVOT_DISABLE_STAR_SCHEMA");
    assert_eq!(base_result.rows, star_result_row_scan.rows);

    let mut group = c.benchmark_group("pivot_star_schema");
    group.sample_size(10);
    group.measurement_time(Duration::from_secs(10));

    group.bench_with_input(BenchmarkId::new("base_table_group_by", rows), &rows, |b, _| {
        b.iter(|| {
            let result =
                pivot(&model, "Sales", &base_group_by, &measures, &FilterContext::empty()).unwrap();
            black_box(result);
        })
    });

    group.bench_with_input(BenchmarkId::new("dimension_group_by", rows), &rows, |b, _| {
        b.iter(|| {
            let result =
                pivot(&model, "Sales", &star_group_by, &measures, &FilterContext::empty()).unwrap();
            black_box(result);
        })
    });

    group.bench_with_input(
        BenchmarkId::new("snowflake_dimension_group_by", rows),
        &rows,
        |b, _| {
            b.iter(|| {
                let result = pivot(
                    &model,
                    "Sales",
                    &snowflake_group_by,
                    &measures,
                    &FilterContext::empty(),
                )
                .unwrap();
                black_box(result);
            })
        },
    );

    std::env::set_var("FORMULA_DAX_PIVOT_DISABLE_STAR_SCHEMA", "1");
    group.bench_with_input(
        BenchmarkId::new("dimension_group_by_row_scan", rows),
        &rows,
        |b, _| {
            b.iter(|| {
                let result =
                    pivot(&model, "Sales", &star_group_by, &measures, &FilterContext::empty())
                        .unwrap();
                black_box(result);
            })
        },
    );
    std::env::remove_var("FORMULA_DAX_PIVOT_DISABLE_STAR_SCHEMA");

    group.finish();

    let mut row_scan_group = c.benchmark_group("pivot_star_schema_row_scan");
    // Row-scan is much slower than the planned/columnar paths; keep samples small so the bench
    // remains usable on developer machines.
    // Criterion requires `sample_size >= 10`, so we use the minimum and rely on a shorter
    // measurement window to keep runtime reasonable.
    row_scan_group.sample_size(10);
    row_scan_group.measurement_time(Duration::from_secs(5));

    row_scan_group.bench_with_input(
        BenchmarkId::new("dimension_group_by_region_unplannable", rows),
        &rows,
        |b, _| {
            b.iter(|| {
                let result = pivot(
                    &model,
                    "Sales",
                    &star_region_group_by,
                    &row_scan_measures,
                    &FilterContext::empty(),
                )
                .unwrap();
                black_box(result);
            })
        },
    );

    row_scan_group.finish();
}

criterion_group!(benches, bench_pivot_star_schema);
criterion_main!(benches);
