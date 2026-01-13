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

    // Customers dimension.
    let customers_schema = vec![
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Region".to_string(),
            column_type: ColumnType::String,
        },
    ];
    let mut customers_builder = ColumnarTableBuilder::new(customers_schema, options);
    for customer_id in 0..customers {
        let region = region_values[customer_id % regions].clone();
        customers_builder.append_row(&[
            formula_columnar::Value::Number(customer_id as f64),
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

    model
}

fn bench_pivot_star_schema(c: &mut Criterion) {
    // Enable optional pivot path tracing (prints to stderr once per pivot path).
    //
    // This is useful for validating that the bench is exercising:
    // - the columnar fast path for base-table group-bys
    // - the relationship-aware (row-scan) path for dimension group-bys
    std::env::set_var("FORMULA_DAX_PIVOT_TRACE", "1");

    let rows = bench_rows();
    let model = build_star_schema_model(rows);

    let measures = vec![
        PivotMeasure::new("Total Sales", "[Total Sales]").unwrap(),
        PivotMeasure::new("Total Qty", "[Total Qty]").unwrap(),
        PivotMeasure::new("Avg Price", "[Avg Price]").unwrap(),
        PivotMeasure::new("Adjusted Sales", "[Adjusted Sales]").unwrap(),
    ];

    let base_group_by = vec![
        GroupByColumn::new("Sales", "Region"),
        GroupByColumn::new("Sales", "Category"),
    ];
    let star_group_by = vec![
        GroupByColumn::new("Customers", "Region"),
        GroupByColumn::new("Products", "Category"),
    ];

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

    group.finish();
}

criterion_group!(benches, bench_pivot_star_schema);
criterion_main!(benches);
