use criterion::{black_box, criterion_group, criterion_main, BenchmarkId, Criterion};
use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{
    Cardinality, CrossFilterDirection, DataModel, FilterContext, Relationship, Table, Value,
};
use std::sync::Arc;
use std::time::Duration;

fn bench_rows() -> usize {
    std::env::var("FORMULA_DAX_REL_BENCH_ROWS")
        .ok()
        .and_then(|v| v.replace('_', "").parse::<usize>().ok())
        // Keep the default benchmark model size within a range that's large enough to be meaningful,
        // but bounded to avoid accidental "set it to 100M rows and OOM" mistakes.
        .filter(|&v| v >= 100_000 && v <= 5_000_000)
        .unwrap_or(1_000_000)
}

fn build_chain_model(rows: usize, bidirectional: bool) -> DataModel {
    let dim2_regions = 100usize;
    let dim1_customers = 100_000usize;

    let options = TableOptions {
        page_size_rows: 65_536,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let mut model = DataModel::new();

    // Dim2: Regions
    let region_names: Vec<Arc<str>> = (0..dim2_regions)
        .map(|i| Arc::<str>::from(format!("Region_{i:03}")))
        .collect();

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
    for region_id in 0..dim2_regions {
        regions_builder.append_row(&[
            formula_columnar::Value::Number(region_id as f64),
            formula_columnar::Value::String(region_names[region_id].clone()),
        ]);
    }
    model
        .add_table(Table::from_columnar("Regions", regions_builder.finalize()))
        .unwrap();

    // Dim1: Customers (many) -> Regions (one)
    let customers_schema = vec![
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "RegionId".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut customers_builder = ColumnarTableBuilder::new(customers_schema, options);
    for customer_id in 0..dim1_customers {
        let region_id = customer_id % dim2_regions;
        customers_builder.append_row(&[
            formula_columnar::Value::Number(customer_id as f64),
            formula_columnar::Value::Number(region_id as f64),
        ]);
    }
    model
        .add_table(Table::from_columnar("Customers", customers_builder.finalize()))
        .unwrap();

    // Fact: Sales (many) -> Customers (one)
    let sales_schema = vec![
        ColumnSchema {
            name: "CustomerId".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut sales_builder = ColumnarTableBuilder::new(sales_schema, options);
    for i in 0..rows {
        let customer_id = (i.wrapping_mul(13)) % dim1_customers;
        let amount = (i % 100) as f64;
        sales_builder.append_row(&[
            formula_columnar::Value::Number(customer_id as f64),
            formula_columnar::Value::Number(amount),
        ]);
    }
    model
        .add_table(Table::from_columnar("Sales", sales_builder.finalize()))
        .unwrap();

    let dir = if bidirectional {
        CrossFilterDirection::Both
    } else {
        CrossFilterDirection::Single
    };

    model
        .add_relationship(Relationship {
            name: "Customers_Regions".into(),
            from_table: "Customers".into(),
            from_column: "RegionId".into(),
            to_table: "Regions".into(),
            to_column: "RegionId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: dir,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "Sales_Customers".into(),
            from_table: "Sales".into(),
            from_column: "CustomerId".into(),
            to_table: "Customers".into(),
            to_column: "CustomerId".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: dir,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total Sales", "SUM(Sales[Amount])").unwrap();
    // COUNTROWS avoids scanning an additional value column, so it can act as a more direct proxy
    // for resolve_row_sets + relationship propagation overhead.
    model.add_measure("Sales Rows", "COUNTROWS(Sales)").unwrap();

    model
}

fn bench_relationship_propagation(c: &mut Criterion) {
    let rows = bench_rows();

    let model_single = build_chain_model(rows, false);
    let model_both = build_chain_model(rows, true);

    // Filter on the far (Dim2) table.
    let filter = FilterContext::empty().with_column_equals("Regions", "Name", Value::from("Region_005"));

    // Sanity check: filter results should be identical regardless of relationship direction.
    let single_value = model_single.evaluate_measure("Total Sales", &filter).unwrap();
    let both_value = model_both.evaluate_measure("Total Sales", &filter).unwrap();
    assert_eq!(single_value, both_value);

    let mut group = c.benchmark_group("relationship_propagation");
    group.sample_size(10);
    group.measurement_time(Duration::from_secs(10));

    group.bench_with_input(BenchmarkId::new("single_direction", rows), &rows, |b, _| {
        b.iter(|| {
            let value = model_single.evaluate_measure("Total Sales", &filter).unwrap();
            black_box(value);
        })
    });

    group.bench_with_input(BenchmarkId::new("bidirectional", rows), &rows, |b, _| {
        b.iter(|| {
            let value = model_both.evaluate_measure("Total Sales", &filter).unwrap();
            black_box(value);
        })
    });

    group.finish();

    let single_rows = model_single.evaluate_measure("Sales Rows", &filter).unwrap();
    let both_rows = model_both.evaluate_measure("Sales Rows", &filter).unwrap();
    assert_eq!(single_rows, both_rows);

    let mut rows_group = c.benchmark_group("relationship_propagation_countrows");
    rows_group.sample_size(10);
    rows_group.measurement_time(Duration::from_secs(5));

    rows_group.bench_with_input(BenchmarkId::new("single_direction", rows), &rows, |b, _| {
        b.iter(|| {
            let value = model_single.evaluate_measure("Sales Rows", &filter).unwrap();
            black_box(value);
        })
    });

    rows_group.bench_with_input(BenchmarkId::new("bidirectional", rows), &rows, |b, _| {
        b.iter(|| {
            let value = model_both.evaluate_measure("Sales Rows", &filter).unwrap();
            black_box(value);
        })
    });

    rows_group.finish();
}

criterion_group!(benches, bench_relationship_propagation);
criterion_main!(benches);
