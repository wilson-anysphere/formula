use criterion::{black_box, criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions,
};
use formula_dax::{
    pivot, Cardinality, CrossFilterDirection, DataModel, FilterContext, GroupByColumn,
    PivotMeasure, Relationship, Table, Value,
};
use std::time::Duration;

fn bench_rows() -> usize {
    std::env::var("FORMULA_DAX_PIVOT_SCAN_BENCH_ROWS")
        .ok()
        .and_then(|v| v.replace('_', "").parse::<usize>().ok())
        .filter(|&v| v >= 250_000 && v <= 5_000_000)
        .unwrap_or(1_000_000)
}

fn build_model(rows: usize) -> DataModel {
    // Keep the dimension table small so filter resolution is cheap, while the fact table is large
    // enough that materializing all allowed fact rows into a `Vec<usize>` would be expensive.
    let dim_keys = 10_000usize;
    let buckets = 10usize;

    let options = TableOptions {
        page_size_rows: 65_536,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let mut model = DataModel::new();

    // Dimension table.
    let dim_schema = vec![
        ColumnSchema {
            name: "Key".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Bucket".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut dim_builder = ColumnarTableBuilder::new(dim_schema, options);
    for key in 0..dim_keys {
        dim_builder.append_row(&[
            formula_columnar::Value::Number(key as f64),
            formula_columnar::Value::Number((key % buckets) as f64),
        ]);
    }
    model
        .add_table(Table::from_columnar("Dim", dim_builder.finalize()))
        .unwrap();

    // Fact table.
    let fact_schema = vec![
        ColumnSchema {
            name: "DimKey".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut fact_builder = ColumnarTableBuilder::new(fact_schema, options);
    for i in 0..rows {
        let dim_key = (i % dim_keys) as f64;
        let amount = (i % 100) as f64;
        fact_builder.append_row(&[
            formula_columnar::Value::Number(dim_key),
            formula_columnar::Value::Number(amount),
        ]);
    }
    model
        .add_table(Table::from_columnar("Fact", fact_builder.finalize()))
        .unwrap();

    model
        .add_relationship(Relationship {
            name: "Fact_Dim".into(),
            from_table: "Fact".into(),
            from_column: "DimKey".into(),
            to_table: "Dim".into(),
            to_column: "Key".into(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .unwrap();

    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();

    model
}

fn bench_pivot_scan_filtered(c: &mut Criterion) {
    // Force the scan path by disabling the star-schema fast path.
    //
    // Note: the base table is columnar (so we can cheaply build large row counts), but the pivot
    // execution path is the scan-based `planned_row_group_by` fallback.
    std::env::set_var("FORMULA_DAX_PIVOT_DISABLE_STAR_SCHEMA", "1");

    let rows = bench_rows();
    let model = build_model(rows);

    let group_by = vec![GroupByColumn::new("Dim", "Bucket")];
    let measures = vec![PivotMeasure::new("Total", "[Total]").unwrap()];

    // Filter on the (small) dimension table that keeps most fact rows. This triggers row set
    // resolution without materially reducing the fact table scan size.
    let dim_keys = 10_000usize;
    let keep = dim_keys * 95 / 100;
    let keep_values: Vec<Value> = (0..keep).map(|k| Value::from(k as f64)).collect();
    let filtered = FilterContext::empty().with_column_in("Dim", "Key", keep_values);

    let mut group = c.benchmark_group("pivot_scan_filtered");
    group.sample_size(10);
    group.measurement_time(Duration::from_secs(10));
    group.throughput(Throughput::Elements(rows as u64));

    group.bench_with_input(BenchmarkId::new("unfiltered", rows), &rows, |b, _| {
        b.iter(|| {
            let result = pivot(
                &model,
                "Fact",
                &group_by,
                &measures,
                &FilterContext::empty(),
            )
            .unwrap();
            black_box(result);
        })
    });

    group.bench_with_input(
        BenchmarkId::new("filtered_keep_most", rows),
        &rows,
        |b, _| {
            b.iter(|| {
                let result = pivot(&model, "Fact", &group_by, &measures, &filtered).unwrap();
                black_box(result);
            })
        },
    );

    group.finish();
}

criterion_group!(benches, bench_pivot_scan_filtered);
criterion_main!(benches);
