use criterion::{black_box, criterion_group, criterion_main, Criterion};
use formula_dax::{DataModel, FilterContext, Table, Value};
use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use std::sync::Arc;

fn build_vec_model(rows: usize) -> DataModel {
    let mut model = DataModel::new();
    let mut fact = Table::new("Fact", vec!["Cat", "Amount"]);
    for i in 0..rows {
        let cat = match i % 10 {
            0 => "A",
            1 => "B",
            2 => "C",
            3 => "D",
            4 => "E",
            5 => "F",
            6 => "G",
            7 => "H",
            8 => "I",
            _ => "J",
        };
        fact.push_row(vec![Value::from(cat), Value::from((i % 100) as f64)])
            .unwrap();
    }
    model.add_table(fact).unwrap();
    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    model
        .add_measure("Distinct Cats", "DISTINCTCOUNT(Fact[Cat])")
        .unwrap();
    model
}

fn build_columnar_model(rows: usize) -> DataModel {
    let mut model = DataModel::new();

    let schema = vec![
        ColumnSchema {
            name: "Cat".to_string(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let options = TableOptions {
        page_size_rows: 65_536,
        cache: PageCacheConfig { max_entries: 8 },
    };
    let mut builder = ColumnarTableBuilder::new(schema, options);

    for i in 0..rows {
        let cat = match i % 10 {
            0 => "A",
            1 => "B",
            2 => "C",
            3 => "D",
            4 => "E",
            5 => "F",
            6 => "G",
            7 => "H",
            8 => "I",
            _ => "J",
        };
        builder.append_row(&[
            formula_columnar::Value::String(Arc::<str>::from(cat)),
            formula_columnar::Value::Number((i % 100) as f64),
        ]);
    }

    model
        .add_table(Table::from_columnar("Fact", builder.finalize()))
        .unwrap();
    model.add_measure("Total", "SUM(Fact[Amount])").unwrap();
    model
        .add_measure("Distinct Cats", "DISTINCTCOUNT(Fact[Cat])")
        .unwrap();
    model
}

fn bench_columnar_vs_vec(c: &mut Criterion) {
    let rows = 200_000;
    let vec_model = build_vec_model(rows);
    let columnar_model = build_columnar_model(rows);

    c.bench_function("vec_sum_unfiltered", |b| {
        b.iter(|| {
            let v = vec_model
                .evaluate_measure("Total", &FilterContext::empty())
                .unwrap();
            black_box(v)
        })
    });

    c.bench_function("columnar_sum_unfiltered", |b| {
        b.iter(|| {
            let v = columnar_model
                .evaluate_measure("Total", &FilterContext::empty())
                .unwrap();
            black_box(v)
        })
    });

    c.bench_function("vec_distinctcount_unfiltered", |b| {
        b.iter(|| {
            let v = vec_model
                .evaluate_measure("Distinct Cats", &FilterContext::empty())
                .unwrap();
            black_box(v)
        })
    });

    c.bench_function("columnar_distinctcount_unfiltered", |b| {
        b.iter(|| {
            let v = columnar_model
                .evaluate_measure("Distinct Cats", &FilterContext::empty())
                .unwrap();
            black_box(v)
        })
    });
}

criterion_group!(benches, bench_columnar_vs_vec);
criterion_main!(benches);

